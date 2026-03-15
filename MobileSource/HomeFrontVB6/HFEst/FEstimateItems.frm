VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FEstimateItems 
   BackColor       =   &H8000000B&
   Caption         =   "Budgets"
   ClientHeight    =   10680
   ClientLeft      =   6270
   ClientTop       =   1800
   ClientWidth     =   13605
   Icon            =   "FEstimateItems.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   13605
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   2  'Blackness
      Height          =   240
      Left            =   16905
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   139
      Top             =   945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Frame frmJob 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   270
      TabIndex        =   55
      Top             =   705
      Width           =   14685
      Begin HFEst.VBCombo cboCommunity 
         Height          =   240
         Left            =   1395
         TabIndex        =   40
         Top             =   855
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   423
         Enabled         =   0   'False
      End
      Begin HFEst.VBCombo cboPhase 
         Height          =   240
         Left            =   1395
         TabIndex        =   41
         Top             =   1110
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   423
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtAddress 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1395
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   39
         Top             =   600
         Width           =   2865
      End
      Begin HFEst.Slider Slider2 
         Height          =   1380
         Left            =   8010
         Top             =   240
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   2434
         Max             =   13425
      End
      Begin VB.TextBox txtNotes 
         BorderStyle     =   0  'None
         Height          =   1380
         Left            =   8070
         MaxLength       =   8000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Text            =   "FEstimateItems.frx":000C
         Top             =   245
         Width           =   4905
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   1395
         MaxLength       =   50
         TabIndex        =   38
         Top             =   360
         Width           =   2865
      End
      Begin VB.TextBox txtJob 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   1395
         MaxLength       =   12
         TabIndex        =   37
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txtLot 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   13290
         MaxLength       =   10
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtBlock 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   13290
         MaxLength       =   10
         TabIndex        =   43
         Top             =   750
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtLotPlan 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   13305
         MaxLength       =   10
         TabIndex        =   44
         Top             =   1020
         Visible         =   0   'False
         Width           =   930
      End
      Begin HFEst.VBCombo cboModel 
         Height          =   240
         Left            =   1395
         TabIndex        =   42
         Top             =   1365
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   423
         Enabled         =   0   'False
      End
      Begin VSFlex8Ctl.VSFlexGrid gProperties 
         Height          =   1380
         Left            =   4410
         TabIndex        =   45
         Top             =   240
         Width           =   3495
         _cx             =   1990858773
         _cy             =   1990855042
         Appearance      =   2
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
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   15
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FEstimateItems.frx":000E
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
         OutlineCol      =   1
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
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   141
         Top             =   615
         Width           =   570
      End
      Begin VB.Label Label15lbp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot/Block/Plan"
         Height          =   195
         Left            =   13245
         TabIndex        =   138
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblModelDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Description"
         Height          =   195
         Left            =   1365
         TabIndex        =   132
         Top             =   1650
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   1380
         Index           =   1
         Left            =   8055
         Top             =   225
         Width           =   4905
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   1380
         Index           =   0
         Left            =   4395
         Top             =   225
         Width           =   3525
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         Left            =   8070
         TabIndex        =   62
         Top             =   30
         Width           =   510
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Properties"
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
         Left            =   4410
         TabIndex        =   61
         Top             =   30
         Width           =   870
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   60
         Top             =   390
         Width           =   795
      End
      Begin VB.Label lblJobNumber 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quote Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   59
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label lblCommunity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area/Project"
         Height          =   195
         Left            =   390
         TabIndex        =   58
         Top             =   870
         Width           =   900
      End
      Begin VB.Label Label1120 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Left            =   840
         TabIndex        =   57
         Top             =   1365
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phase"
         Height          =   195
         Left            =   840
         TabIndex        =   56
         Top             =   1110
         Width           =   450
      End
   End
   Begin VB.Frame frmBilling 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   7545
      Left            =   4965
      TabIndex        =   63
      Top             =   11040
      Visible         =   0   'False
      Width           =   15135
      Begin VSFlex8Ctl.VSFlexGrid gContract 
         Height          =   3795
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   10275
         _cx             =   1990870732
         _cy             =   1990859302
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FEstimateItems.frx":00C2
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
         OutlineCol      =   1
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         Begin VB.Label lblMoveLine 
            BackColor       =   &H8000000D&
            Height          =   75
            Index           =   0
            Left            =   0
            TabIndex        =   116
            Top             =   0
            Visible         =   0   'False
            Width           =   75
         End
      End
      Begin VB.Frame frmBillingTasks 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6405
         Left            =   10440
         TabIndex        =   124
         Top             =   0
         Width           =   4275
         Begin VB.Label lblTask 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Create Retainage/Holdback Invoice"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   6
            Left            =   780
            TabIndex        =   134
            Top             =   2280
            Width           =   2580
         End
         Begin VB.Image imgTask 
            Height          =   360
            Index           =   6
            Left            =   360
            Picture         =   "FEstimateItems.frx":03F4
            Top             =   2190
            Width           =   360
         End
         Begin VB.Image imgTask 
            Height          =   360
            Index           =   5
            Left            =   360
            Picture         =   "FEstimateItems.frx":0ADE
            Stretch         =   -1  'True
            Top             =   3120
            Width           =   360
         End
         Begin VB.Label lblTask 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post the selected invoice to accounting"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   5
            Left            =   780
            TabIndex        =   131
            Top             =   3150
            Width           =   2805
         End
         Begin VB.Image imgTask 
            Height          =   360
            Index           =   4
            Left            =   360
            Picture         =   "FEstimateItems.frx":1068
            Stretch         =   -1  'True
            Top             =   2730
            Width           =   360
         End
         Begin VB.Label lblTask 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview the selected invoice"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   4
            Left            =   780
            TabIndex        =   130
            Top             =   2760
            Width           =   2040
         End
         Begin VB.Image imgTask 
            Height          =   360
            Index           =   3
            Left            =   330
            Picture         =   "FEstimateItems.frx":15F2
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lblTask 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edit the job addons"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   3
            Left            =   750
            TabIndex        =   129
            Top             =   840
            Width           =   1740
         End
         Begin VB.Label lblTask 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add items to the schedule of values"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   750
            TabIndex        =   128
            Top             =   1200
            Width           =   2520
         End
         Begin VB.Label lblTask 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Approve requests and create a change orders"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   1
            Left            =   750
            TabIndex        =   127
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label lblTask 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Generate a new invoice"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   2
            Left            =   780
            TabIndex        =   126
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contract and Billing Tasks:"
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
            Left            =   270
            TabIndex        =   125
            Top             =   450
            Width           =   2310
         End
         Begin VB.Image imgTask 
            Height          =   360
            Index           =   0
            Left            =   330
            Picture         =   "FEstimateItems.frx":1CDC
            Top             =   1110
            Width           =   360
         End
         Begin VB.Image imgTask 
            Height          =   360
            Index           =   1
            Left            =   330
            Picture         =   "FEstimateItems.frx":23C6
            Top             =   1500
            Width           =   360
         End
         Begin VB.Image imgTask 
            Height          =   360
            Index           =   2
            Left            =   360
            Picture         =   "FEstimateItems.frx":2AB0
            Top             =   1890
            Width           =   360
         End
      End
      Begin HFEst.Slider Slider3 
         Height          =   60
         Left            =   120
         Top             =   4140
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   106
         Orientation     =   1
         Max             =   15090
      End
      Begin VSFlex8Ctl.VSFlexGrid gInvoices 
         Height          =   2835
         Left            =   120
         TabIndex        =   121
         Top             =   4590
         Width           =   10275
         _cx             =   1990870732
         _cy             =   1990857609
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FEstimateItems.frx":319A
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
         OutlineCol      =   1
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         Begin VB.Label lblMoveLine 
            BackColor       =   &H8000000D&
            Height          =   75
            Index           =   1
            Left            =   0
            TabIndex        =   122
            Top             =   0
            Visible         =   0   'False
            Width           =   75
         End
      End
      Begin VB.Label lblInvoices 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice History"
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
         Left            =   120
         TabIndex        =   123
         Top             =   4290
         Width           =   1290
      End
      Begin VB.Image imgShortcut 
         Height          =   240
         Index           =   7
         Left            =   1860
         Picture         =   "FEstimateItems.frx":344D
         ToolTipText     =   "More Info..."
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule of Values"
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
         Left            =   120
         TabIndex        =   64
         Top             =   120
         Width           =   1665
      End
   End
   Begin VB.Frame frmPurchasing 
      BorderStyle     =   0  'None
      Caption         =   "frmPurchasing"
      Height          =   8385
      Left            =   60
      TabIndex        =   65
      Top             =   2610
      Width           =   21495
      Begin HFEst.Slider Slider 
         Height          =   5355
         Left            =   5370
         Top             =   180
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   9446
         Max             =   13425
      End
      Begin VB.Frame frmEstItems 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6705
         Left            =   6795
         TabIndex        =   70
         Top             =   870
         Width           =   14565
         Begin VB.Frame frmPOIndex 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Height          =   3030
            Left            =   3915
            TabIndex        =   98
            Top             =   2430
            Visible         =   0   'False
            Width           =   12585
            Begin VB.CommandButton cmdApprovePO 
               Caption         =   "Approve PO"
               Height          =   285
               Left            =   6150
               TabIndex        =   34
               Top             =   2280
               Width           =   1425
            End
            Begin VB.CheckBox chkIncludeDocuments 
               Alignment       =   1  'Right Justify
               Caption         =   "Attach Docs"
               Height          =   315
               Left            =   6120
               TabIndex        =   35
               Top             =   2595
               Width           =   1215
            End
            Begin VB.CheckBox chkHidePrice 
               Alignment       =   1  'Right Justify
               Caption         =   "Hide Price"
               Height          =   255
               Left            =   6120
               TabIndex        =   30
               Top             =   870
               Width           =   1215
            End
            Begin VB.CheckBox chkHideQty 
               Alignment       =   1  'Right Justify
               Caption         =   "Hide Qty"
               Height          =   255
               Left            =   6120
               TabIndex        =   31
               Top             =   1125
               Width           =   1215
            End
            Begin VB.TextBox txtRetainagePercent 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
               Height          =   230
               Left            =   7140
               MaxLength       =   50
               TabIndex        =   33
               Text            =   "99%"
               Top             =   1635
               Width           =   375
            End
            Begin VB.CheckBox chkTotalOnly 
               Alignment       =   1  'Right Justify
               Caption         =   "Total Only"
               Height          =   255
               Left            =   6120
               TabIndex        =   32
               Top             =   1365
               Width           =   1215
            End
            Begin VB.TextBox txtVendor 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   120
               Width           =   4155
            End
            Begin VB.TextBox txtPONumber 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   3090
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   20
               Top             =   375
               Width           =   2445
            End
            Begin VB.TextBox txtPOIndex 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   19
               Top             =   375
               Width           =   1695
            End
            Begin VB.TextBox txtPODescription 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   100
               TabIndex        =   21
               Top             =   630
               Width           =   4155
            End
            Begin VB.TextBox txtFOB 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   24
               Top             =   1395
               Width           =   4155
            End
            Begin VB.TextBox txtTerms 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   2000
               TabIndex        =   25
               Top             =   1650
               Width           =   4155
            End
            Begin VB.TextBox txtShipVia 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   23
               Top             =   1140
               Width           =   4155
            End
            Begin VB.TextBox txtOrderedBy 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   26
               Top             =   1905
               Width           =   4155
            End
            Begin VB.TextBox txtStandardText 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   2000
               TabIndex        =   22
               Top             =   885
               Width           =   4155
            End
            Begin VB.TextBox txtDeliveryAddress 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   250
               TabIndex        =   28
               Top             =   2415
               Width           =   4155
            End
            Begin VB.TextBox txtDeliveryRecipient 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   27
               Top             =   2160
               Width           =   4155
            End
            Begin VB.TextBox txtDateIssued 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   240
               Left            =   1380
               TabIndex        =   29
               Text            =   "Aug 31, 2007"
               Top             =   2670
               Width           =   2070
            End
            Begin VB.Label lblApproval 
               AutoSize        =   -1  'True
               Caption         =   "Approved by DANS: Jan 14, 2014"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   6000
               TabIndex        =   135
               Top             =   2040
               Width           =   2415
            End
            Begin VB.Label lblPOHistoryReport 
               Caption         =   "Open the PO history report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   6000
               TabIndex        =   112
               Top             =   120
               Width           =   2355
            End
            Begin VB.Label Label155 
               Caption         =   "Retainage"
               Height          =   255
               Left            =   6150
               TabIndex        =   111
               Top             =   1635
               Width           =   975
            End
            Begin VB.Label lblPOHistory 
               Caption         =   "Posted to accounting Printed Jun 14"
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   6000
               TabIndex        =   110
               Top             =   360
               Width           =   1665
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vendor"
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
               Left            =   480
               TabIndex        =   109
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label1g 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Standard Text"
               Height          =   195
               Left            =   330
               TabIndex        =   108
               Top             =   900
               Width           =   1005
            End
            Begin VB.Label Label1d 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Description"
               Height          =   195
               Left            =   540
               TabIndex        =   107
               Top             =   630
               Width           =   795
            End
            Begin VB.Label lblPOIndex 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Purchase Order"
               Height          =   195
               Left            =   225
               TabIndex        =   106
               Top             =   390
               Width           =   1110
            End
            Begin VB.Label Label125 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Free On Board"
               Height          =   195
               Left            =   300
               TabIndex        =   105
               Top             =   1410
               Width           =   1035
            End
            Begin VB.Label Label124 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Terms"
               Height          =   195
               Left            =   900
               TabIndex        =   104
               Top             =   1665
               Width           =   435
            End
            Begin VB.Label Label123 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Ship Via"
               Height          =   195
               Left            =   750
               TabIndex        =   103
               Top             =   1140
               Width           =   585
            End
            Begin VB.Label Label122 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Ordered By"
               Height          =   195
               Left            =   540
               TabIndex        =   102
               Top             =   1920
               Width           =   795
            End
            Begin VB.Label Label120 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Email/Fax"
               Height          =   195
               Left            =   630
               TabIndex        =   101
               Top             =   2430
               Width           =   705
            End
            Begin VB.Label Label1r 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Recipient"
               Height          =   195
               Left            =   660
               TabIndex        =   100
               Top             =   2190
               Width           =   675
            End
            Begin VB.Label Label1i 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Issued"
               Height          =   195
               Left            =   870
               TabIndex        =   99
               Top             =   2700
               Width           =   465
            End
            Begin VB.Image cmdEditVendor 
               Height          =   240
               Left            =   1110
               Picture         =   "FEstimateItems.frx":39D7
               Top             =   120
               Width           =   240
            End
            Begin VB.Image cmdBrowse 
               Height          =   240
               Index           =   3
               Left            =   5550
               Picture         =   "FEstimateItems.frx":3F61
               Top             =   1635
               Width           =   240
            End
            Begin VB.Image cmdBrowse 
               Height          =   240
               Index           =   4
               Left            =   5550
               Picture         =   "FEstimateItems.frx":40AB
               Top             =   885
               Width           =   240
            End
            Begin VB.Image cmdBrowse 
               Height          =   240
               Index           =   5
               Left            =   5550
               Picture         =   "FEstimateItems.frx":41F5
               Top             =   630
               Width           =   240
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid gItems 
            Height          =   2655
            Left            =   1095
            TabIndex        =   36
            Top             =   4440
            Width           =   7695
            _cx             =   1990866181
            _cy             =   1990857291
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   0   'False
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
            BackColorSel    =   14336431
            ForeColorSel    =   -2147483630
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
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
            SelectionMode   =   0
            GridLines       =   4
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   130
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FEstimateItems.frx":433F
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
            OutlineCol      =   1
            Ellipsis        =   0
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   1
            OwnerDraw       =   2
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   1
            OleDropMode     =   1
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   3
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
            Begin VB.PictureBox picWarningMessages 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   2040
               ScaleHeight     =   585
               ScaleWidth      =   3765
               TabIndex        =   113
               Top             =   960
               Visible         =   0   'False
               Width           =   3795
               Begin VB.TextBox lblWarningMessages 
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   315
                  Left            =   360
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   114
                  Text            =   "FEstimateItems.frx":5575
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.Image imgTipIcon 
                  Height          =   240
                  Left            =   60
                  Top             =   60
                  Width           =   240
               End
            End
            Begin VB.Image imgInfo 
               Height          =   240
               Left            =   0
               Picture         =   "FEstimateItems.frx":557B
               Top             =   480
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image imgWarning 
               Height          =   240
               Left            =   0
               Picture         =   "FEstimateItems.frx":5B05
               Top             =   240
               Visible         =   0   'False
               Width           =   240
            End
         End
         Begin VB.Frame frmAssembly 
            BorderStyle     =   0  'None
            Height          =   2790
            Left            =   60
            TabIndex        =   71
            Top             =   645
            Visible         =   0   'False
            Width           =   12585
            Begin VB.TextBox txtOther 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               TabIndex        =   9
               Top             =   1365
               Width           =   4875
            End
            Begin VB.TextBox txtStyle 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               TabIndex        =   7
               Top             =   1110
               Width           =   1815
            End
            Begin VB.TextBox txtFinish 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   3210
               TabIndex        =   8
               Top             =   1110
               Width           =   3045
            End
            Begin VB.TextBox txtEstimatorNotes 
               BorderStyle     =   0  'None
               Height          =   870
               Left            =   5055
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   1860
               Width           =   4800
            End
            Begin VB.TextBox txtHFComments 
               BorderStyle     =   0  'None
               Height          =   870
               Left            =   180
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Top             =   1860
               Width           =   4800
            End
            Begin VB.TextBox txtLocation 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   3210
               TabIndex        =   6
               Top             =   855
               Width           =   3045
            End
            Begin VB.TextBox txtColor 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               TabIndex        =   5
               Top             =   855
               Width           =   1815
            End
            Begin VB.TextBox txtQuantity 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   4410
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   3
               Top             =   345
               Width           =   915
            End
            Begin VB.TextBox txtJCExtra 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   3060
               MaxLength       =   10
               TabIndex        =   1
               Top             =   90
               Width           =   1395
            End
            Begin VB.TextBox txtHFCategory 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   2
               Top             =   345
               Width           =   2115
            End
            Begin VB.TextBox txtHFDescription 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               MaxLength       =   200
               TabIndex        =   4
               Top             =   600
               Width           =   4155
            End
            Begin VB.TextBox txtHFOption 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   0
               Top             =   90
               Width           =   1695
            End
            Begin VB.Frame frmLockBudgets 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   1245
               Left            =   10830
               TabIndex        =   79
               Top             =   1050
               Width           =   1455
               Begin VB.Image imgLockBudgets 
                  Height          =   480
                  Index           =   0
                  Left            =   510
                  Picture         =   "FEstimateItems.frx":608F
                  Top             =   600
                  Width           =   480
               End
               Begin VB.Label lblBudgetLock 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Budgets are open. Changes are permitted."
                  ForeColor       =   &H8000000D&
                  Height          =   765
                  Left            =   120
                  TabIndex        =   80
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.Image imgLockBudgets 
                  Height          =   480
                  Index           =   1
                  Left            =   510
                  Picture         =   "FEstimateItems.frx":6959
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   480
               End
            End
            Begin VB.Frame frmChangeRequest 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   915
               Left            =   6990
               TabIndex        =   72
               Top             =   90
               Visible         =   0   'False
               Width           =   5295
               Begin VB.TextBox txtChangeRequestedBy 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   1290
                  MaxLength       =   50
                  TabIndex        =   10
                  Top             =   -15
                  Width           =   1005
               End
               Begin VB.TextBox txtChangeRequestedDate 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   1290
                  MaxLength       =   20
                  TabIndex        =   11
                  Text            =   "June 4, 2010"
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.TextBox txtChangeApprovedBy 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   3915
                  MaxLength       =   50
                  TabIndex        =   13
                  Top             =   -15
                  Width           =   1005
               End
               Begin VB.TextBox txtChangeApprovedDate 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   3915
                  MaxLength       =   20
                  TabIndex        =   14
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.TextBox txtPrice 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   3915
                  MaxLength       =   20
                  TabIndex        =   15
                  Top             =   570
                  Width           =   1005
               End
               Begin HFEst.VBCombo cboChangeRequestStatus 
                  Height          =   240
                  Left            =   1290
                  TabIndex        =   12
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   1244
                  _ExtentY        =   423
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Requested By"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   78
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Date Requested"
                  Height          =   195
                  Left            =   75
                  TabIndex        =   77
                  Top             =   240
                  Width           =   1170
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Approved By"
                  Height          =   195
                  Left            =   2955
                  TabIndex        =   76
                  Top             =   0
                  Width           =   915
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Date Approved"
                  Height          =   195
                  Left            =   2790
                  TabIndex        =   75
                  Top             =   240
                  Width           =   1080
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Status"
                  Height          =   195
                  Left            =   795
                  TabIndex        =   74
                  Top             =   600
                  Width           =   450
               End
               Begin VB.Image cmdBrowse 
                  Height          =   240
                  Index           =   7
                  Left            =   2310
                  Picture         =   "FEstimateItems.frx":7223
                  Top             =   240
                  Width           =   240
               End
               Begin VB.Image cmdBrowse 
                  Height          =   240
                  Index           =   8
                  Left            =   4935
                  Picture         =   "FEstimateItems.frx":736D
                  Top             =   240
                  Width           =   240
               End
               Begin VB.Image cmdBrowse 
                  Height          =   240
                  Index           =   9
                  Left            =   4935
                  Picture         =   "FEstimateItems.frx":74B7
                  Top             =   570
                  Width           =   240
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Selling Price"
                  Height          =   195
                  Left            =   3000
                  TabIndex        =   73
                  Top             =   570
                  Width           =   870
               End
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Other"
               Height          =   195
               Left            =   885
               TabIndex        =   136
               Top             =   1365
               Width           =   390
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Style/Finish"
               Height          =   195
               Left            =   105
               TabIndex        =   133
               Top             =   1110
               Width           =   1185
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Estimator Notes"
               Height          =   195
               Left            =   5055
               TabIndex        =   120
               Top             =   1650
               Width           =   1110
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Comments"
               Height          =   195
               Left            =   180
               TabIndex        =   119
               Top             =   1650
               Width           =   735
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Color/Location"
               Height          =   195
               Left            =   165
               TabIndex        =   118
               Top             =   870
               Width           =   1170
            End
            Begin VB.Image imgShortcut 
               Height          =   240
               Index           =   6
               Left            =   5820
               Picture         =   "FEstimateItems.frx":7601
               Stretch         =   -1  'True
               ToolTipText     =   "File Attachments..."
               Top             =   60
               Width           =   240
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Attachments"
               Height          =   195
               Left            =   4920
               TabIndex        =   117
               Top             =   90
               Width           =   885
            End
            Begin VB.Label lblQuantity 
               AutoSize        =   -1  'True
               Caption         =   "lblQuantity"
               ForeColor       =   &H80000011&
               Height          =   195
               Left            =   6060
               TabIndex        =   86
               Top             =   540
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblUOM 
               AutoSize        =   -1  'True
               Caption         =   "lblUOM"
               ForeColor       =   &H80000011&
               Height          =   195
               Left            =   6270
               TabIndex        =   85
               Top             =   270
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label Label144 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Quantity"
               Height          =   195
               Left            =   3780
               TabIndex        =   84
               Top             =   360
               Width           =   585
            End
            Begin VB.Label Label1c 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Category"
               Height          =   195
               Left            =   705
               TabIndex        =   83
               Top             =   360
               Width           =   630
            End
            Begin VB.Label Label1dd 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Description"
               Height          =   195
               Left            =   540
               TabIndex        =   82
               Top             =   615
               Width           =   795
            End
            Begin VB.Label lblAssemblyExtra 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Asmbly/Extra"
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
               Left            =   210
               TabIndex        =   81
               Top             =   90
               Width           =   1125
            End
            Begin VB.Image cmdBrowse 
               Height          =   240
               Index           =   6
               Left            =   4470
               Picture         =   "FEstimateItems.frx":7B8B
               Top             =   90
               Width           =   240
            End
         End
         Begin VB.Frame frmChangeOrder 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   1875
            Left            =   315
            TabIndex        =   87
            Top             =   165
            Visible         =   0   'False
            Width           =   12585
            Begin VB.TextBox txtCONumber 
               BorderStyle     =   0  'None
               Height          =   230
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   92
               Top             =   90
               Width           =   1695
            End
            Begin VB.TextBox txtCODateApproved 
               BorderStyle     =   0  'None
               Height          =   230
               Left            =   5880
               MaxLength       =   100
               TabIndex        =   91
               Top             =   330
               Width           =   2445
            End
            Begin VB.TextBox txtCOApprovedBy 
               BorderStyle     =   0  'None
               Height          =   230
               Left            =   5880
               TabIndex        =   90
               Top             =   90
               Width           =   2445
            End
            Begin VB.TextBox txtCOComments 
               BorderStyle     =   0  'None
               Height          =   1140
               Left            =   1380
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   89
               Top             =   630
               Width           =   11085
            End
            Begin VB.TextBox txtCODate 
               BorderStyle     =   0  'None
               Height          =   230
               Left            =   1380
               MaxLength       =   10
               TabIndex        =   88
               Top             =   330
               Width           =   2445
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Change Order"
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
               Left            =   150
               TabIndex        =   97
               Top             =   90
               Width           =   1185
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Date Approved"
               Height          =   195
               Index           =   0
               Left            =   4755
               TabIndex        =   96
               Top             =   330
               Width           =   1080
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Comments"
               Height          =   195
               Left            =   600
               TabIndex        =   95
               Top             =   630
               Width           =   735
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Approved By"
               Height          =   195
               Left            =   4920
               TabIndex        =   94
               Top             =   90
               Width           =   915
            End
            Begin VB.Image cmdBrowse 
               Height          =   240
               Index           =   0
               Left            =   3840
               Picture         =   "FEstimateItems.frx":7CD5
               Top             =   330
               Width           =   240
            End
            Begin VB.Image cmdBrowse 
               Height          =   240
               Index           =   1
               Left            =   8340
               Picture         =   "FEstimateItems.frx":7E1F
               Top             =   330
               Width           =   240
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Date"
               Height          =   195
               Left            =   990
               TabIndex        =   93
               Top             =   330
               Width           =   345
            End
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid gImport 
         Height          =   1395
         Left            =   930
         TabIndex        =   115
         Top             =   4170
         Visible         =   0   'False
         Width           =   2145
         _cx             =   1990856392
         _cy             =   1990855069
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
         Height          =   6615
         Left            =   -30
         TabIndex        =   48
         Top             =   90
         Width           =   5355
         _cx             =   1990862054
         _cy             =   1990864276
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FEstimateItems.frx":7F69
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         Ellipsis        =   1
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
         Begin VB.Image imgShortcut 
            Height          =   240
            Index           =   8
            Left            =   1590
            Picture         =   "FEstimateItems.frx":80C4
            Stretch         =   -1  'True
            ToolTipText     =   "Expand All"
            Top             =   15
            Width           =   240
         End
         Begin VB.Image imgShortcut 
            Height          =   240
            Index           =   0
            Left            =   2805
            Picture         =   "FEstimateItems.frx":864E
            ToolTipText     =   "More Info..."
            Top             =   15
            Width           =   240
         End
         Begin VB.Image imgShortcut 
            Height          =   240
            Index           =   1
            Left            =   2550
            Picture         =   "FEstimateItems.frx":8BD8
            Stretch         =   -1  'True
            ToolTipText     =   "File Attachments..."
            Top             =   15
            Width           =   240
         End
         Begin VB.Image imgShortcut 
            Height          =   240
            Index           =   3
            Left            =   1905
            Picture         =   "FEstimateItems.frx":9162
            Stretch         =   -1  'True
            ToolTipText     =   "WBS Codes..."
            Top             =   15
            Width           =   240
         End
         Begin VB.Image imgShortcut 
            Height          =   240
            Index           =   2
            Left            =   2220
            Picture         =   "FEstimateItems.frx":96EC
            Stretch         =   -1  'True
            ToolTipText     =   "Contacts..."
            Top             =   15
            Width           =   240
         End
         Begin VB.Image imgShortcut 
            Height          =   240
            Index           =   5
            Left            =   1275
            Picture         =   "FEstimateItems.frx":9C76
            Stretch         =   -1  'True
            ToolTipText     =   "Reload"
            Top             =   15
            Width           =   240
         End
      End
      Begin VB.Frame frmRFPs 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   10440
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   8505
         Begin HFEst.Slider Slider1 
            Height          =   60
            Left            =   120
            Top             =   2340
            Width           =   7710
            _ExtentX        =   13600
            _ExtentY        =   106
            Orientation     =   1
            Max             =   13425
         End
         Begin VB.TextBox txtRFPDescription 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   230
            Left            =   990
            MaxLength       =   250
            TabIndex        =   49
            Top             =   60
            Width           =   5415
         End
         Begin VB.TextBox txtRFPComments 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   645
            Left            =   1005
            MaxLength       =   8000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   360
            Width           =   6705
         End
         Begin VSFlex8Ctl.VSFlexGrid gBids 
            Height          =   1215
            Left            =   120
            TabIndex        =   51
            Top             =   1110
            Width           =   7695
            _cx             =   1990866181
            _cy             =   1990854751
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   0   'False
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
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
            SelectionMode   =   3
            GridLines       =   0
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FEstimateItems.frx":A200
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
            OutlineCol      =   1
            Ellipsis        =   0
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   1
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
         Begin VSFlex8Ctl.VSFlexGrid gBidItems 
            Height          =   2655
            Left            =   630
            TabIndex        =   52
            Top             =   3030
            Width           =   7695
            _cx             =   1990866181
            _cy             =   1990857291
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   0   'False
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
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
            SelectionMode   =   0
            GridLines       =   4
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   64
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FEstimateItems.frx":A309
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
            OutlineCol      =   1
            Ellipsis        =   0
            ExplorerBar     =   7
            PicturesOver    =   -1  'True
            FillStyle       =   1
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Attachments"
            Height          =   195
            Left            =   6570
            TabIndex        =   69
            Top             =   90
            Width           =   885
         End
         Begin VB.Image imgShortcut 
            Height          =   240
            Index           =   4
            Left            =   7470
            Picture         =   "FEstimateItems.frx":AADC
            Stretch         =   -1  'True
            ToolTipText     =   "File Attachments..."
            Top             =   60
            Width           =   240
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   90
            Width           =   795
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Comments"
            Height          =   195
            Left            =   180
            TabIndex        =   67
            Top             =   360
            Width           =   735
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000010&
            Height          =   645
            Index           =   2
            Left            =   990
            Top             =   345
            Width           =   6705
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   1058
      ButtonWidth     =   1746
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   25
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpenJob"
                  Text            =   "Job"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpenProject"
                  Text            =   "Project"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "saveasquote"
                  Text            =   "Save as New Quote"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "saveasjob"
                  Text            =   "Save as New Job"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "converttojob"
                  Text            =   "Convert Quote to Job"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s1"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "View"
            Key             =   "View"
            Style           =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save As"
            Key             =   "SaveAssembly"
            Object.ToolTipText     =   "Save to Model/Option Library"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Snap Shot"
            Key             =   "snapshots"
            Style           =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "OneTime"
            Key             =   "TakeoffOneTime"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Item"
            Key             =   "TakeoffItem"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Assembly"
            Key             =   "TakeoffAssembly"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "PlanSwift"
            Key             =   "TakeoffPlanSwift"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Pipeline"
            Key             =   "TakeoffPipeline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Custom"
            Key             =   "TakeoffCustom"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Loc/wbs"
            Key             =   "takeoffsettings"
            Style           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Preview"
            Key             =   "Preview"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Rfrsh Costs"
            Key             =   "RePrice"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "New RFQ"
            Key             =   "NewRFQ"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Generate"
            Key             =   "Generate"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Send PO's"
            Key             =   "SendPOs"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Send RFQ's"
            Key             =   "SendRFQs"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s4"
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "PO's"
            Key             =   "ViewPOs"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Budgets"
            Key             =   "ViewBudgets"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkCancelledItems 
         Caption         =   "Hide Cancelled Items"
         Height          =   180
         Left            =   16980
         TabIndex        =   142
         Top             =   255
         Width           =   1920
      End
      Begin VB.CheckBox chkToolbarCaptions 
         Caption         =   "Captions"
         Height          =   195
         Left            =   16980
         TabIndex        =   137
         Top             =   30
         Width           =   1920
      End
      Begin VB.Timer Timer1 
         Left            =   14190
         Top             =   60
      End
   End
   Begin VB.Label Label25 
      Caption         =   "picbox for ---> combining icons"
      Height          =   540
      Left            =   15600
      TabIndex        =   140
      Top             =   900
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "FEstimateItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FEstimateItems::"
Public EventTraps As Collection

Private mFunnyFlag As Boolean ' used to ignore the first mouse move event see gitems_mousemove and gitems_rowcolchanged
    
Private Type ViewDefs
    Name         As String
    KeyFlds      As String
    SortFlds     As String
    DisplayFlds  As String
    IconKeys     As String
End Type



Private mReadOnly   As Boolean
    
    
    
Private mDirty                  As Boolean
Private mViews()                As ViewDefs

Private mTakeoffSettings As String 'csv list of location,wbs1,wbs2...  set in ftakeoffsettings and AddItem()
Private mUseChangeRequests As Boolean
Private Enum BidItemStatuses
    bsProposed
    bsAccepted
    bsDeclined
End Enum


'number of grouped columns
Private GroupedColumns As Long


'conditional formatting
Private AcceptedForeColor As Long
Private AcceptedBackColor As Long
Private AcceptedStyle     As String
Private DeclinedForeColor As Long
Private DeclinedBackColor As Long
Private DeclinedStyle     As String
Private MinForeColor      As Long
Private MinBackColor      As Long
Private MinStyle          As String
Private MaxForeColor      As Long
Private MaxBackColor      As Long
Private MaxStyle          As String


Public Enum FCEMode
    fceQuote = 1
    fceBudget
    fcePO
    fceContract
    fceBids
    fceBilling
End Enum


'job property rows
Private mprop_PermitNumber As Integer
Private mprop_PermitDate As Integer
Private mprop_ShellTemplate As Integer
Private mprop_UnitTemplate As Integer
Private mprop_ConstStart As Integer
Private mprop_IntacctDepartment As Integer
Private mprop_TarionBuilderNumber  As Integer
Private mprop_TarionEnrollmentNumber As Integer
Private mprop_LabTax As Integer
Private mprop_MatTax As Integer
Private mprop_SubTax As Integer
Private mprop_EquTax As Integer
Private mprop_OvrTax As Integer
Private mprop_OthTax As Integer
Private mprop_ARTax As Integer
Private mprop_UserFlds As Integer


Private mMode           As FCEMode
Private mJob            As String
Private mCommunity      As String
Private mCommunityPhase As String
Private mRFP            As Long
Private mTakeoffSystemBidID  As Long
Private mTakeoffSystemProjectName As String

Private mEstAssemblyID          As Long 'gets set when the assembly header frame is shown. is used if jcextra is changed
Private mAssemblyChanged        As Boolean
Private mCustomerNo             As String 'these get set when the assembly tree is double clicked. see loaditems()
Private mHFLocation             As Long
Private mChangeOrder            As String

Private VarianceCatDef As String

Private mViewIndex As Long
Private mLock As Boolean 'stupid flag see txtDateDue.Change,keydown,keyup
Private mSplitting As Boolean

Private mTimerTask As String 'stupid menus


'these are for takeoffs - identifys the selected row
Private mCurrentAssemblyType    As AssemblyTypes
Private mCurrentJob             As String
Private mCurrentJCExtra         As String
Private mCurrentChangeOrder     As String
Private mCurrentCommunity       As String
Private mCurrentCommunityPhase  As String
Private mCurrentAssembly        As String
Private mCurrentAssemblyDesc    As String
Public mCurrentEstAssemblyID    As Long
Private mCurrentJobDesc         As String
Private mCurrentModel           As String
Private mCurrentOptionID        As String
Private mCurrentBudgetsLocked   As Boolean
Private mCancelTakeoff          As Boolean

'used purely for debugging
Private mViewQuery As String

Private Const mcASMBlY_ADDCO = 0
Private Const mcASMBlY_ADDCR = 1
Private Const mcASMBlY_ADD = 2
Private Const mcASMBlY_DELETE = 4
Private Const mcASMBlY_RENAME = 5
Private Const mcASMBlY_FINALIZE = 7
Private Const mcASMBlY_ATTACHQUOTE = 8
Private Const mcASMBlY_COPYFROMJOB = 9
Private Const mcASMBlY_PRINT = 11
Private Const mcITEM_COPYITEMS = 0
Private Const mcITEM_SPLITITEMS = 1
Private Const mcITEM_SUBSTITUEITEM = 2
Private Const mcITEM_REMOVEITEMS = 3
Private Const mcITEM_MODIFYQTY = 4
Private Const mcITEM_VIEWFILES = 6
Private Const mcITEM_CANCELBUDGETS = 8
Private Const mcITEM_SAVEONETIMETODB = 10
Private Const mcITEM_UPDATEPRICELIST = 11
Private Const mcITEM_COMPAREPRICES = 12
Private Const mcITEM_FORMATTING = 13
                                
Private Const mcRFP_DELETE = 0
                                
                                
Private Const mcPO_PREVIEW = 0
Private Const mcPO_PRINT = 1
Private Const mcPO_SEND = 2
Private Const mcPO_REPRICE = 4
Private Const mcPO_CANCEL = 6
Private Const mcPO_EDITVENDOR = 8
Private Const mcPO_CHANGE = 9
                                
Private Const mcINV_PREVIEW = 0
Private Const mcINV_PRINT = 1
Private Const mcINV_NEW = 3
Private Const mcINV_EDIT = 4
Private Const mcINV_DELETE = 5
Private Const mcINV_VOID = 6
Private Const mcINV_POST = 8
                                
                                
Private Const mcVENDOR_EDIT = 0
                                
Private Const mcBIDITEM_ACCEPTED = 0
Private Const mcBIDITEM_DECLINED = 1
Private Const mcBIDITEM_COMMENTS = 2
Private Const mcBIDITEM_EXPORT = 4
Private Const mcBIDITEM_IMPORT = 5
Private Const mcBIDITEM_FORMATTING = 7
Private Const mcBIDITEM_VENDOR = 8

Private Const mcRFPITEM_ADD = 0
Private Const mcRFPITEM_REMOVE = 1


Private Const mcBID_ADD = 0
Private Const mcBID_REMOVE = 1
Private Const mcBID_EXPORT = 3
Private Const mcBID_IMPORT = 4
Private Const mcBID_VENDOR = 6




Private Sub LoadProperties()
On Error GoTo eh
    
    Dim s As String
    Dim rs As ADODB.Recordset
    Dim r As Long
    Dim Category As String
    Dim fieldname As String
    Dim fields As Recordset
    Dim data   As Recordset
    Dim TaxGroups As String
    Dim IntacctDepartments As String
    Dim ScheduleTemplates As String
    
   
    If PostPOQtyToAccounting Then
        HFApp.Options.ValueByName("UseVarianceReporting") = "False"
        gItems.TextMatrix(0, gItems.ColIndex("VarianceJCCategory")) = ""
        gItems.ColHidden(gItems.ColIndex("VarianceJCCategory")) = True
        gItems.TextMatrix(0, gItems.ColIndex("VarianceJCCategoryDesc")) = ""
        gItems.ColHidden(gItems.ColIndex("VarianceJCCategoryDesc")) = True
    Else
        If gItems.TextMatrix(0, gItems.ColIndex("VarianceJCCategory")) = "" Then gItems.TextMatrix(0, gItems.ColIndex("VarianceJCCategory")) = "VarianceJCCategory"
        If gItems.TextMatrix(0, gItems.ColIndex("VarianceJCCategoryDesc")) = "" Then gItems.TextMatrix(0, gItems.ColIndex("VarianceJCCategoryDesc")) = "VarianceJCCategoryDesc"
    End If


On Error Resume Next
    TaxGroups = ""
    Set rs = HFApp.SqlExec("select taxgroup,description from TaxGroups where DivisionID = " & HFApp.DivisionID & " order by 1")
    While Not rs.EOF
        TaxGroups = TaxGroups & "|" & rs(0) & vbTab & rs(1)
        rs.MoveNext
    Wend
    TaxGroups = Mid(TaxGroups, 2)
    
    IntacctDepartments = ""
    Set rs = HFApp.SqlExec("select id,name from IntacctDepartments where DivisionID = " & HFApp.DivisionID & " order by 1")
    While Not rs.EOF
        IntacctDepartments = IntacctDepartments & "|" & rs(0) & vbTab & rs(1)
        rs.MoveNext
    Wend
    IntacctDepartments = Mid(IntacctDepartments, 2)
    

    ScheduleTemplates = HFApp.Options.ValueByName("ScheduleTemplates")
    
    s = ""
    s = s & "select" & vbCrLf
    s = s & "  j.shellscheduletemplate" & vbCrLf
    s = s & " ,j.scheduletemplate" & vbCrLf
    s = s & " ,j.permitnumber" & vbCrLf
    s = s & " ,j.permitreceiveddate" & vbCrLf
    s = s & " ,j.start_date" & vbCrLf
    s = s & " ,j.LabourTaxGroup" & vbCrLf
    s = s & " ,j.MaterialTaxGroup" & vbCrLf
    s = s & " ,j.SubcontractTaxGroup" & vbCrLf
    s = s & " ,j.EquipmentTaxGroup" & vbCrLf
    s = s & " ,j.OverheadTaxGroup" & vbCrLf
    s = s & " ,j.OtherTaxGroup" & vbCrLf
    s = s & " ,j.ARTaxGroup" & vbCrLf
    s = s & " ,j.IntacctDepartment" & vbCrLf
    s = s & " ,j.TarionEnrollmentNumber" & vbCrLf
    s = s & " ,isnull(isnull(p.TarionBuilderNumber,c.TarionBuilderNumber),j.TarionBuilderNumber) TarionBuilderNumber" & vbCrLf
    s = s & " ,case when isnull(p.TarionBuilderNumber,'')<>'' then 1" & vbCrLf
    s = s & "       when isnull(c.TarionBuilderNumber,'')<>'' then 1" & vbCrLf
    s = s & "       else 0 end TarionBuilderReadOnly" & vbCrLf
    s = s & "from tblJobs j" & vbCrLf
    s = s & "left join tbllocality c on c.area=j.community" & vbCrLf
    s = s & "left join CommunityPhase p on p.community=j.community and p.communityphase=j.communityphase" & vbCrLf
    s = s & "where j.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "and j.job_no=" & DbQuote(Str, mJob)
    Set rs = HFApp.SqlExec(s)
    With gProperties
        r = -1
        .Rows = 0
        
        
        'add built in properties first
        ' -- permit, permitdate, scheduletemplate, startdate
        Category = "Scheduling"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = FMain.SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_PermitNumber = r
        .TextMatrix(r, .ColIndex("name")) = "Permit Number"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("PermitNumber")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_PermitDate = r
        .TextMatrix(r, .ColIndex("name")) = "Permit Received Date"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("PermitReceivedDate")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = DateTime
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        
        
If IsMultiFamily Then
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_ShellTemplate = r
        .TextMatrix(r, .ColIndex("name")) = "Shell Schedule Template"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("ShellScheduleTemplate")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = ScheduleTemplates
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_UnitTemplate = r
        .TextMatrix(r, .ColIndex("name")) = "Unit Schedule Template"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("ScheduleTemplate")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = ScheduleTemplates
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
Else
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_UnitTemplate = r
        .TextMatrix(r, .ColIndex("name")) = "Schedule Template"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("ScheduleTemplate")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = ScheduleTemplates
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
End If
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_ConstStart = r
        .TextMatrix(r, .ColIndex("name")) = "Construction Start Date"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("Start_Date")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = DateTime
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True

        
        
'add built in properties first
' -- Tarion
If HFApp.Options.ValueByName("EnableTarionFields") = "true" Then
        Category = "Tarion"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = FMain.SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True

        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_TarionBuilderNumber = r
        .TextMatrix(r, .ColIndex("name")) = "Builder No"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("TarionBuilderNumber")
        .TextMatrix(r, .ColIndex("category")) = Category
        
        If "" & rs("TarionBuilderReadOnly") = 1 Then
            .TextMatrix(r, .ColIndex("DataType")) = "READONLY"
            .Cell(flexcpForeColor, r, .ColIndex("name")) = vbGrayText
            .Cell(flexcpForeColor, r, .ColIndex("value")) = vbGrayText
        Else
            .TextMatrix(r, .ColIndex("DataType")) = Str
            .Cell(flexcpForeColor, r, .ColIndex("name")) = vbWindowText
            .Cell(flexcpForeColor, r, .ColIndex("value")) = vbWindowText
        End If
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True

        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_TarionEnrollmentNumber = r
        .TextMatrix(r, .ColIndex("name")) = "Job Enrollment No"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("TarionEnrollmentNumber")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True

End If

        
If HFApp.Options(AccountingSystem) = asIntacct Then
        

        Category = "Intacct"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = FMain.SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_IntacctDepartment = r
        .TextMatrix(r, .ColIndex("name")) = "Department"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("IntacctDepartment")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = IntacctDepartments
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
End If
        
        ' -- tax groups
        Category = "Tax Groups"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = FMain.SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True
        
        
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_LabTax = r
        .TextMatrix(r, .ColIndex("name")) = "Labour"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("LabourTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_MatTax = r
        .TextMatrix(r, .ColIndex("name")) = "Material"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("MaterialTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
                    
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_SubTax = r
        .TextMatrix(r, .ColIndex("name")) = "Subcontract"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("SubContractTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_EquTax = r
        .TextMatrix(r, .ColIndex("name")) = "Equipment"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("EquipmentTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_OvrTax = r
        .TextMatrix(r, .ColIndex("name")) = "Overhead"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("OverheadTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_OthTax = r
        .TextMatrix(r, .ColIndex("name")) = "Other"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("OtherTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True

        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_ARTax = r
        .TextMatrix(r, .ColIndex("name")) = "Accounts Receivable"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("ARTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
On Error GoTo eh
            
            
        'now load user defined fields
        mprop_UserFlds = r + 1
        Set fields = HFApp.SqlExec("SELECT * FROM JobCustomFieldDefs ORDER BY Category,Name")
        Set data = HFApp.SqlExec("SELECT * FROM JobCustomFields WHERE divisionid=" & HFApp.DivisionID & " and Job_No=" & DbQuote(Str, mJob))
        
        While Not fields.EOF
            
            If Category <> "" & fields("category") Then
                Category = "" & fields("category")
                r = r + 1
                .AddItem ""
                .TextMatrix(r, .ColIndex("category")) = Category
                .TextMatrix(r, .ColIndex("name")) = IIf(Category = "", "unclassified", Category)
                Set .Cell(flexcpPicture, r, .ColIndex("name")) = FMain.SmallIcons.ListImages("category").Picture
                .RowOutlineLevel(r) = 0
                .IsSubtotal(r) = True
            End If
            
            r = r + 1
            .AddItem ""
            fieldname = "" & fields("Name")
            .TextMatrix(r, .ColIndex("category")) = Category
            .TextMatrix(r, .ColIndex("name")) = fieldname
            .TextMatrix(r, .ColIndex("PickList")) = "" & fields("PickList")
            Select Case data(fieldname).Type
                Case adInteger, adTinyInt, adSmallInt, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt: .TextMatrix(r, .ColIndex("DataType")) = NumInt
                Case adDouble, adSingle, adDecimal, adNumeric:                                                                           .TextMatrix(r, .ColIndex("DataType")) = Num
                Case adCurrency:                                                                                                         .TextMatrix(r, .ColIndex("DataType")) = Cur
                Case adBoolean:                                                                                                          .TextMatrix(r, .ColIndex("DataType")) = Bit
                Case adDate, adDBDate, adDBTime, adDBTimeStamp:                                                                          .TextMatrix(r, .ColIndex("DataType")) = DateTime
                Case adVarChar, adBSTR, adChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adVariant:                           .TextMatrix(r, .ColIndex("DataType")) = Str
            End Select
            .TextMatrix(r, .ColIndex("length")) = data(fieldname).DefinedSize
            On Error Resume Next
            .TextMatrix(r, .ColIndex("value")) = data(fieldname).value
            On Error GoTo eh
            
            .RowOutlineLevel(r) = 1
            .IsSubtotal(r) = True
            
            fields.MoveNext
        Wend
        
        
        
        
        Call .AutoSize(0, .Cols - 1)
        Call .Outline(0)
        
        

        
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadProperties")
End Sub






Private Sub cboChangeRequestStatus_Change()
    mAssemblyChanged = True
    Dirty = True
End Sub

Private Sub cboChangeRequestStatus_Click()
    mAssemblyChanged = True
    Dirty = True
End Sub

Private Sub cboChangeRequestStatus_GotFocus()
    SelectAll cboChangeRequestStatus
End Sub

Private Sub cboCommunity_Click()
    Dirty = True
End Sub


Private Sub cboPhase_Change()
    Dirty = True
End Sub

Private Sub cboPhase_Click()
    Dirty = True
End Sub



Private Sub chkCancelledItems_Click()
    Call ShowCancelledItems
End Sub

Private Sub chkToolbarCaptions_Click()
    Call ShowToolbarCaptions(Toolbar, chkToolbarCaptions.value = vbChecked)
    Call Form_Resize
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim s As String
    
    
    If Not cmdBrowse(Index).Enabled Then Exit Sub
    
    Select Case Index
        
                        
        Case 9:    mTimerTask = "ShowAssemblyAddons"
                   Timer1.Interval = 10
                   Timer1.Enabled = True
                    
        Case 0:    If Not txtCODate.Locked Then Call DCalendar.Popup(txtCODate)
        Case 1:    If Not txtCODateApproved.Locked Then Call DCalendar.Popup(txtCODateApproved)
        
        
        Case 7:    Call DCalendar.Popup(Me.txtChangeRequestedDate):  mAssemblyChanged = True:  Dirty = True
        Case 8:    Call DCalendar.Popup(Me.txtChangeApprovedDate):   mAssemblyChanged = True:  Dirty = True
        
        Case 3:    s = txtTerms.Text
                   If FComments.Edit(s, txtTerms, , "Terms", 2000) Then txtTerms.Text = s
        Case 4:    s = txtStandardText.Text
                   If FComments.Edit(s, txtStandardText, , "Standard Text", 2000) Then txtStandardText.Text = s
        Case 5:    s = txtPODescription.Text
                   If FComments.Edit(s, txtPODescription, , "Description", 100) Then txtPODescription.Text = s
        Case 6:    txtJCExtra.SetFocus
                   Select Case HFApp.Options(AccountingSystem)
                       Case asTimberline
                           If FPickList.Choose(HFApp.Databases(dbAccounting), "Extra", "select Extra,Description from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(mJob)), txtJCExtra.Text) Then
                               txtJCExtra.Text = FPickList.SelectedItem("Extra")
                               Dirty = True
                           End If
                       Case asMasterBuilder
                           If FPickList.Choose(HFApp.Databases(dbAccounting), "Job Phase", "select phsnum Phase,phsnme Description from jobphs where recnum=" & DbQuote(Num, HFApp.FormatJob(mJob)), txtJCExtra.Text) Then
                               txtJCExtra.Text = FPickList.SelectedItem("Phase")
                               Dirty = True
                           End If
                   End Select
    End Select
End Sub



Private Sub cmdEditVendor_Click()
    Call HFApp.EditVendor(txtVendor.Tag)
End Sub

Private Sub Form_Activate()
    'dont remove. this ensures job header is kept visible
    Call Form_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim s As String
    Select Case True
    
        Case KeyCode = vbKeyR And Shift <> 0
            If mMode = fcePO Then
                s = ""
                s = s & "SELECT DISTINCT v.Community+'.'+v.CommunityPhase ""Comm/Phase"",v.Assembly,v.Model,v.POVendor+' -- '+v.POVendorName Vendor,v.EstPhase + '/' + v.EstItem+' -- '+v.ItemDesc Item,v.PORate ""Current""" & vbCrLf
                s = s & "      ,dbo.Purch_GetItemRate(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor,getdate()," & HFApp.DivisionID & ") PriceList" & vbCrLf
                s = s & "      ,case dbo.Purch_GetItemRateQuality(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor,getdate()," & HFApp.DivisionID & ") when 4  then '1 phase quote' when 5  then '2 community quote' when 6  then '3 global quote' when 9  then '4 phase rate' when 10 then '5 community rate'when 11 then '6 global rate'when 12 then '7 list price' end Type" & vbCrLf
                s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                s = s & " WHERE v.PODeleted=0" & vbCrLf
                s = s & "   AND v.POGenBatch=0" & vbCrLf
            Else
                s = ""
                s = s & "SELECT DISTINCT v.Community+'.'+v.CommunityPhase ""Comm/Phase"",v.Assembly,v.Model,v.BudgetVendor+' -- '+v.BudgetVendorName Vendor,v.EstPhase + '/' + v.EstItem+' -- '+v.ItemDesc Item,v.BudgetRate ""Current""" & vbCrLf
                s = s & "      ,dbo.Purch_GetItemRate(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.BudgetVendor,getdate()," & HFApp.DivisionID & ") PriceList" & vbCrLf
                s = s & "      ,case dbo.Purch_GetItemRateQuality(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.BudgetVendor,getdate()," & HFApp.DivisionID & ") when 4  then '1 phase quote' when 5  then '2 community quote' when 6  then '3 global quote' when 9  then '4 phase rate' when 10 then '5 community rate'when 11 then '6 global rate'when 12 then '7 list price' end Type" & vbCrLf
                s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                s = s & " WHERE v.BudgetDeleted=0" & vbCrLf
                s = s & "   AND v.BudgetGenerated=0" & vbCrLf
            End If
            s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "i.DivisionID")
            If Shift = vbCtrlMask + vbAltMask Then
                Call FDBGrid.ShowForm("Item Prices", s, "Item Prices", "", True)
            ElseIf Shift = vbCtrlMask + vbAltMask + vbShiftMask Then
                Call MsgBox(s, vbInformation, "Item Prices")
            End If
            
        Case KeyCode = vbKeyF11
            Shell "calc.exe"
            
        Case KeyCode = vbKeyS And Shift = vbCtrlMask
            Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
            
        Case KeyCode = vbKeyO And Shift = vbCtrlMask
            Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
            
    End Select
End Sub







Private Sub gAssemblies_OLECompleteDrag(Effect As Long)
    With gAssemblies
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = vbWindowBackground
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, 0) = vbWindowText
    End With
End Sub

Private Sub gAssemblies_OLEDragDrop(data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim s As String
    Dim i As Long
    
    
    If Effect = vbDropEffectCopy Then
    
        If Shift = vbCtrlMask Then
            'copy
            If Dirty Then Call SaveData(False)
            i = gAssemblies.MouseRow
            s = ""
            s = s & "insert into estimateitems(EstAssemblyID,POIndex,Phase,Item,Job,JCExtra,JCCostCode,JCCategory,SortOrder,Description,Comments,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,PONumber,ExcludeFromPO,OriginalJCCategory,BudgetGenerated,BudgetPostingBatch,POGenBatch,BudgetDeleted,PODeleted,BudgetOverridden,POOverridden,Assembly,AssemblyDescription,Model)" & vbCrLf
            s = s & "select " & DbQuote(Num, gAssemblies.Cell(flexcpText, i, 1)) & ",POIndex,Phase,Item,Job,JCExtra,JCCostCode,JCCategory,SortOrder,Description,Comments,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,'',ExcludeFromPO,OriginalJCCategory,0,0,0,0,0,POOverridden,POOverridden,Assembly,AssemblyDescription,Model" & vbCrLf
            s = s & "  from estimateitems" & vbCrLf
            s = s & " where estitemid in(" & Mid(data.GetData(vbCFText), 2) & ")"
            Call HFApp.SqlExec(s, dbHomefront)
        Else
            'move
            If Dirty Then Call SaveData(False)
            i = gAssemblies.MouseRow
            
            s = ""
            s = s & "update estimateitems" & vbCrLf
            s = s & "   set estassemblyid=" & DbQuote(Num, gAssemblies.Cell(flexcpText, i, 1)) & vbCrLf
            s = s & " where estitemid in(" & Mid(data.GetData(vbCFText), 2) & ")"
            Call HFApp.SqlExec(s, dbHomefront)
                    
            With gItems
                For i = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    gItems.RemoveItem (i)
                Next
            End With
        End If
    End If
        
    With gAssemblies
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = vbWindowBackground
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, 0) = vbWindowText
    End With
    
End Sub

Private Sub gAssemblies_OLEDragOver(data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
Dim r As Long
    With gAssemblies
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = vbWindowBackground
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, 0) = vbWindowText
        
        r = .MouseRow
        If r < 0 Then
            Effect = vbDropEffectNone
            Exit Sub
        End If
        
        If Left(data.GetData(vbCFText), 1) <> Chr(2) Then
            Effect = vbDropEffectNone
            Exit Sub
        End If
                
        If .Cell(flexcpData, r, 1) = "EstAssemblyID" Then
            Effect = vbDropEffectCopy
            .Cell(flexcpBackColor, r, 0) = .BackColorSel 'vbHighlight
            .Cell(flexcpForeColor, r, 0) = .ForeColorSel 'vbHighlightText
        Else
            Effect = vbDropEffectNone
        End If
        
    End With
End Sub



Private Sub gBidItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call UpdateBidTotals
End Sub

Private Sub gBidItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    If ReadOnly Then
        Cancel = True
        Exit Sub
    End If
    
    With gBidItems
        For i = Min(.Col, .ColSel) To Max(.Col, .ColSel)
            If .ColData(i) <> "VendorRate" Then
                Cancel = True
                Exit Sub
            End If
        Next
    End With
    
End Sub


Private Sub gBidItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    With gBidItems
    If Button = vbRightButton Then
        If .MouseCol < 0 Or .MouseRow < 0 Then Exit Sub
        
        Call SetCtrlFocus(gBidItems)
        Call GridGotFocus(gBidItems)
        If .MouseRow = 0 Then
            .Col = .MouseCol
            If .ColData(.Col) = "VendorRate" Then
                .Col = .MouseCol
                .Row = 1
                .RowSel = .Rows - 1
                FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_ACCEPTED).checked = .Cell(flexcpData, .Row, .Col) = bsAccepted
                FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_COMMENTS).Enabled = False
                FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_DECLINED).checked = .Cell(flexcpData, .Row, .Col) = bsDeclined
                If Not ReadOnly Then
                    PopupMenu FMain.mnuEstimateBidItemsGrid
                End If
            Else
                Call FMain.ShowColumnMenu(gBidItems)
                .ColHidden(.ColIndex("ItemDesc")) = False
            End If
        ElseIf .ColData(.MouseCol) = "VendorRate" Then
            If Not GridCellSelected(gBidItems, .MouseRow, .MouseCol) Then Call gBidItems.Select(.MouseRow, .MouseCol)
            FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_COMMENTS).Enabled = True
            FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_ACCEPTED).Enabled = .Col = .ColSel
            FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_ACCEPTED).checked = .Cell(flexcpData, .Row, .Col) = bsAccepted
            FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_DECLINED).checked = .Cell(flexcpData, .Row, .Col) = bsDeclined
            If Not ReadOnly Then
                PopupMenu FMain.mnuEstimateBidItemsGrid
            End If
        Else
            If Not GridCellSelected(gBidItems, .MouseRow, .MouseCol) Then Call gBidItems.Select(.MouseRow, .MouseCol)
            If Not ReadOnly Then
                PopupMenu FMain.mnuEstimateRFPItemsGrid
            End If
        End If
    End If
    End With
End Sub

Private Sub gBidItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Cancel As Boolean
    
    With gBidItems
    Select Case True
    
        Case KeyCode = vbKeyF4 And Shift = 0
            Call mnuEstimateBidItemsGridSub_Click(mcBIDITEM_COMMENTS)
        
        Case KeyCode = vbKeyD And Shift = vbCtrlMask
            FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_DECLINED).checked = .Cell(flexcpData, .Row, .Col) = bsDeclined
            Call mnuEstimateBidItemsGridSub_Click(mcBIDITEM_DECLINED)
            
        Case KeyCode = vbKeyA And Shift = vbCtrlMask
            FMain.mnuEstimateBidItemsGridSub(mcBIDITEM_ACCEPTED).checked = .Cell(flexcpData, .Row, .Col) = bsAccepted
            Call mnuEstimateBidItemsGridSub_Click(mcBIDITEM_ACCEPTED)
    
        Case KeyCode = vbKeyDelete And Shift = 0
            Call gBidItems_BeforeEdit(0, 0, Cancel)
            If Not Cancel Then
                gBidItems.Text = ""
                If .Cell(flexcpData, .Row, .Col) = bsAccepted Then .Cell(flexcpData, .Row, .Col) = bsProposed
                Call gBidItems_ValidateEdit(0, 0, Cancel)
                Call gBidItems_AfterEdit(0, gBidItems.Col)
            End If
    End Select
    End With
End Sub

Private Sub gBidItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim qty  As Double
    Dim rate As Double
    
    With gBidItems
        If .MouseRow < 0 Or .MouseCol < 0 Then
            .ToolTipText = ""
            Exit Sub
        End If
            
        qty = .ValueMatrix(.MouseRow, .ColIndex("POQty"))
        rate = .ValueMatrix(.MouseRow, .MouseCol)
        
        If .ColData(.MouseCol) = "VendorRate" And rate <> 0 Then
            .ToolTipText = qty & " @ " & format(rate, "#,##0.00") & " = " & format(rate * qty, "currency")
        Else
            .ToolTipText = ""
        End If
    End With
        
End Sub

Private Sub gBidItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gBidItems
    Select Case True
        Case Trim(.EditText) = ""
            .Cell(flexcpData, Row, Col) = 0
            If .Cell(flexcpData, Row, Col) = bsAccepted Then .Cell(flexcpData, Row, Col) = bsProposed
        Case Else
            .EditText = format(Val(.EditText), "#,##0.00")
    End Select
    Dirty = True
    End With
End Sub

Private Sub gBids_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gBids
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "Vendor", "VendorName", "Contact", "Fax", "Email", "SentDate", "Status", "Amount", "Contact", "DelMethod"
                Cancel = True
            Case "ExpiryDate", "Comments"
                .ComboList = "|..."
            Case Else
        End Select
    End With
End Sub

Private Sub gBids_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Dim b As Boolean
    
    With gBids
    If Button = vbRightButton Then
        Call SetCtrlFocus(gBids)
        Call GridGotFocus(gBids)
        
        If .MouseRow = 0 Then
            .Col = .MouseCol
            Call FMain.ShowColumnMenu(gBids)
        Else
            If .MouseCol > -1 And .MouseRow > -1 Then
                If Not GridCellSelected(gBids, .MouseRow, .MouseCol) Then Call .Select(.MouseRow, .MouseCol)
            End If
            
            On Error Resume Next
            b = False
            b = .Row > 0 And .TextMatrix(.Row, .ColIndex("Vendor")) <> ""
            FMain.mnuEstimateBidsGridSub(mcBID_ADD).Enabled = Not ReadOnly
            FMain.mnuEstimateBidsGridSub(mcBID_EXPORT).Enabled = b And Not ReadOnly
            FMain.mnuEstimateBidsGridSub(mcBID_IMPORT).Enabled = b And Not ReadOnly
            FMain.mnuEstimateBidsGridSub(mcBID_REMOVE).Enabled = b And Not ReadOnly
            FMain.mnuEstimateBidsGridSub(mcBID_VENDOR).Enabled = b
            
            PopupMenu FMain.mnuEstimateBidsGrid
        End If
    End If
    End With

End Sub

Private Sub gBids_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gBids
    
    Select Case .ColKey(Col)
        
        Case "Comments"
            s = .Text
            If FComments.Edit(s, gBids) Then
                .Text = s
            End If
            
        Case "ExpiryDate"
            Call DCalendar.Popup(gBids, .RowPos(.Row) + .RowHeight(.Row), .colPos(.Col))
            .RowData(Row) = "DIRTY"
            Call UpdateBidTotals
        
    End Select
    End With
End Sub

Private Sub gBids_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gBids
        Select Case .ColKey(Col)
            Case "ExpiryDate"
                If .EditText <> "" And Not IsDate(.EditText) Then
                    Cancel = True
                Else
                    .EditText = format(.EditText, "mmm d, yyyy")
                End If
        End Select
        .RowData(Row) = "DIRTY"
        Call UpdateBidTotals
    End With
    Dirty = True
End Sub




Private Sub gContract_AfterMoveColumn(ByVal Col As Long, Position As Long)
    lblMoveLine(0).Visible = False
End Sub

Private Sub gContract_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    lblMoveLine(0).Visible = False
End Sub

Private Sub gContract_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    lblMoveLine(0).Visible = False
End Sub

Private Sub gContract_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContract
    .ComboList = ""
    Select Case .ColKey(Col)
        Case "Description":   .EditMaxLength = 150
        Case "Price", "HoldBackRate":        .EditMaxLength = 0
        Case "TaxGroup":        .ComboList = "|..."
        Case "RevenueAccount":  .ComboList = "|..."
        Case "RevenueAccountDescription":  .ComboList = "..."
        Case Else:            Cancel = True
    End Select
    End With
End Sub

Private Sub gContract_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    
    With gContract
        Select Case .ColKey(Col)
            Case "TaxGroup"
                s = "SELECT TaxGroup,Description,GroupRate FROM TaxGroups where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gContract) Then
                    .Cell(flexcpText, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("TaxGroup")
                    Dirty = True
                End If
                
            Case "RevenueAccount", "RevenueAccountDescription"
                s = "SELECT Account,Description FROM GLAccounts where DivisionID =" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Accounts", s, gContract) Then
                    .Cell(flexcpText, .Row, .ColIndex("RevenueAccount"), .RowSel, .ColIndex("RevenueAccount")) = FPickList.SelectedItem("Account")
                    .Cell(flexcpText, .Row, .ColIndex("RevenueAccountDescription"), .RowSel, .ColIndex("RevenueAccountDescription")) = FPickList.SelectedItem("description")
                    Dirty = True
                End If
                
        End Select
    End With
End Sub

Private Sub gContract_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And Shift <> 0 Then
    With gContract
        If .Row >= 1 And .Row < .Rows - 1 Then
            If .TextMatrix(.Row, .ColIndex("Billed")) = 0 Then
                .RowHidden(.Row) = True
                Dirty = True
            End If
        End If
    End With
    End If
End Sub

Private Sub gContract_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If gContract.MouseRow = 0 Then
            Call FMain.ShowColumnMenu(gContract)
        ElseIf gContract.MouseRow = gContract.Rows - 1 Then
        Else
            PopupMenu FMain.mnuDelete
        End If
    End If
End Sub

Private Sub gContract_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContract
        If Row = .Rows - 1 Then
            .RowData(Row) = "New"
            .AddItem ""
        End If
    End With
End Sub

Private Sub gContract_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As Recordset
    With gContract
        Select Case .ColKey(Col)
            Case "Description"
            Case "TaxGroup"
               If HFApp.SqlExec("Select isnull(count(*),0) from TaxGroups where DivisionID = " & HFApp.DivisionID & " and TaxGroup = " & DbQuote(Str, .EditText))(0) = 0 Then
                    MsgBox "Invalid Tax Group"
                    Cancel = True
               End If
            
            Case "RevenueAccount"
                Set rs = HFApp.SqlExec("Select account,description from GLAccounts where DivisionID = " & HFApp.DivisionID & " and Account= " & DbQuote(Str, .EditText))
                If rs.EOF Then
                    MsgBox "Invalid Account"
                    Cancel = True
                Else
                    .EditText = Val(.EditText)
                    .Cell(flexcpText, Min(.Row, .RowSel), .ColIndex("RevenueAccountDescription"), Max(.Row, .RowSel), .ColIndex("RevenueAccountDescription")) = "" & rs("description")
                End If
            
            Case "Price"
                .EditText = Round(Val(.EditText), 2)
        End Select
        If .RowData(Row) = "" Then .RowData(Row) = "Dirty"
        Dirty = True
    End With
End Sub




Private Sub gInvoices_DblClick()
Dim s As String
With gInvoices
    If .ValueMatrix(.Row, 0) <> 0 Then
        Screen.MousePointer = vbHourglass
        s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\ARInvoice.rpt")
        If Not FileExists(s) Then
            MsgBox "file not found" & vbCrLf & _
                   "create a report at" & vbCrLf & _
                   s & vbCrLf & vbCrLf & _
                   "Add a numeric parameter named 'Invoice'.", vbInformation, App.ProductName
        Else
            'Set f = New FRptViewer
            'Call f.ShowReport(s, True, False, "Invoice", .ValueMatrix(.Row, 0))
            
            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            Call c.ParameterValue("Invoice", .ValueMatrix(.Row, 0))
            On Error GoTo 0
            Call c.PrintPreview("Print Preview")
            
            
            Screen.MousePointer = vbDefault
        End If
    End If
End With
End Sub




Private Sub gInvoices_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And Shift <> 0 Then
        If gInvoices.Row >= 1 Then Call Me.mnuInvoicesSub_Click(mcINV_DELETE)
    End If
End Sub

Private Sub gitems_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
    Dim r As Long
    
    With gItems
    
        If Row < 1 Then Row = 1
    
        r = Row + 1
        While r <= .Rows - 1
            If .IsSubtotal(r) Then
                'rehide hidden subtotal rows
                If .TextMatrix(r, .ColIndex("EstItemID")) = "HIDDEN" Then
                    .RowHidden(r) = True
                End If
            Else
                'hide deleted items
                If ((IsIn(mMode, fceQuote, fceBudget) And .TextMatrix(r, .ColIndex("BudgetDeleted")) = "True") Or ((Not IsIn(mMode, fceQuote, fceBudget)) And .TextMatrix(r, .ColIndex("PODeleted")) = "True")) Then
                    .RowHidden(r) = True
                End If
            End If
            r = r + 1
        Wend
    End With
    Call ShowCancelledItems

End Sub

Private Sub gItems_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call GroupGrid
End Sub

Private Sub gItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error Resume Next
    
    With gItems
        If OldRow > 0 Then
            .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = vbWindowBackground
            Call ColorizeItems(OldRow)
            .Redraw = flexRDDirect
        End If

        If NewRow > 0 Then
            .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = .BackColorSel 'vbHighlight
        End If



        


    End With
End Sub

Private Sub gItems_AfterSort(ByVal Col As Long, Order As Integer)
gItems.Cell(flexcpBackColor, gItems.Row, 0, gItems.Row, gItems.Cols - 1) = vbButtonFace
End Sub

Private Sub gItems_BeforeSort(ByVal Col As Long, Order As Integer)
On Error Resume Next
    gItems.Cell(flexcpBackColor, 1, 0, gItems.Rows - 1, gItems.Cols - 1) = vbWindowBackground
End Sub

Private Sub gItems_DblClick()
    With gItems
    If .MouseRow < 0 Then Exit Sub
        If .IsSubtotal(.MouseRow) Then
            .GetNode(.MouseRow).Expanded = Not .GetNode(.MouseRow).Expanded
        End If
    End With
End Sub

Private Sub gitems_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, done As Boolean)
    With gItems
        
        'default dont draw
        done = True
        
        'do draw the column header row
        If Row = 0 Then
            done = False
            Exit Sub
        End If
        
        'do draw group header rows but only columns to the left
        If .IsSubtotal(Row) And Col >= .RowOutlineLevel(Row) Then
            done = False
            Exit Sub
        End If
        
        'do draw data rows but only the ungrouped columns
        If Not .IsSubtotal(Row) And Not .ColData(Col) = "GROUPED" Then
            done = False
            Exit Sub
        End If
    
        'exception: we always want to see the itemdesc
        If Col = .ColIndex("ItemDesc") Then
            done = False
            Exit Sub
        End If
    
    End With
End Sub

Private Sub gItems_OLEStartDrag(data As VSFlex8Ctl.VSDataObject, AllowedEffects As Long)
    Dim i As Long
    Dim s As String
    With gItems
        For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            s = s & "," & .TextMatrix(i, .ColIndex("EstItemID"))
        Next
    End With
    s = Chr(2) & Mid(s, 2)
    Call data.SetData(s, 1)
End Sub



Private Sub imgShortcut_Click(Index As Integer)
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
        
    If mJob = "" Then Exit Sub
    
    
    Select Case Index
    
        Case 0, 7 'job info
            frmJob.Visible = Not frmJob.Visible
            Call Form_Resize
        
        Case 8
            Call GridExpandALL(Me.gAssemblies)
                    
                    
        Case 1 'job documents
            mTimerTask = "EditAttachments|J~" & mJob & "|Job Documents"
            Timer1.Interval = 10
            Timer1.Enabled = True
            
        Case 2 'contacts
            Call FContacts.ShowForm(mJob)
        
        Case 3 'wbs codes
            Call FWBSCodes.ShowForm(mJob)
            Call LoadWBSDescriptions
    
        Case 4 'rfp documents
            If mRFP <> 0 Then
                mTimerTask = "EditAttachments|J~" & mJob & "|Job Documents|RFP~" & mRFP & "|RFP Documents"
                Timer1.Interval = 10
                Timer1.Enabled = True
            End If
        
        Case 5 'reload
            Call mnuEstimateItemViewsSub_Click(CInt(mViewIndex + 1))
        
        Case 6 'assembly documents
            mTimerTask = "EditAttachments|EstAsm~" & mEstAssemblyID & "|Documents"
            Timer1.Interval = 10
            Timer1.Enabled = True

    End Select
    
    
eh: Exit Sub
End Sub



Private Sub lblJobNumber_Click()
    If txtJob.Text <> "" Then
        If SaveData(True) Then
            Call HFApp.RunTask("EditJob|" & txtJob.Text)
            Call LoadProperties
        End If
    End If
End Sub


Private Sub lblPOHistoryReport_Click()
On Error GoTo eh
    Dim s As String
    s = HFApp.SystemFolder & "System\Reports\Estimating\POHistory.rpt"
    'Call FRptViewer.ShowReport(s, True, True, "Job", mJob, "POIndex", txtPOIndex.Text)
    
    Dim c As New ZybUtil.Crystal
    Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
    On Error Resume Next
    Call c.ParameterValue("DivisionID", HFApp.DivisionID)
    Call c.ParameterValue("Job", mJob)
    Call c.ParameterValue("POIndex", txtPOIndex.Text)
    On Error GoTo eh
    Call c.PrintPreview("Print Preview")

Exit Sub
eh: Call errHandler(SRCFILE & "lblPOHistoryReport_Click")
End Sub





Private Sub imgTask_Click(Index As Integer)
    Call lblTask_Click(Index)
End Sub
Private Sub lblTask_Click(Index As Integer)
    If Not SaveData(False) Then Exit Sub
    Select Case Index
        Case 0 ' add items to schedule of values
            If FCreateContract.Edit(False, mJob) Then Call LoadContract
            
        Case 1 ' approve requests and create change order
            If FCreateContract.Edit(True, mJob) Then Call LoadContract
            
        Case 2 ' generate invoice
            If Not SaveData(False) Then Exit Sub
            If FCreateInvoice.Edit(mJob) Then Call LoadContract
        
        Case 3
            mTimerTask = "ShowJobAddons"
            Timer1.Interval = 10
            Timer1.Enabled = True
            
        Case 4:  Call mnuInvoicesSub_Click(mcINV_PREVIEW)
        Case 5:  Call mnuInvoicesSub_Click(mcINV_POST)
        Case 6 'Create Holdback Invoice
        If Not SaveData(False) Then Exit Sub
            If FCreateHBInvoice.Edit(mJob) Then Call LoadContract
            
    End Select
End Sub


Private Sub Slider1_Move()
    Call Form_Resize
End Sub

Private Sub Slider2_Move()
Call Form_Resize
End Sub

Private Sub Slider3_Move()
    Call Form_Resize
End Sub


Private Sub Timer1_Timer()
    Dim d As Double
    
    Timer1.Interval = 0
    Select Case mTimerTask
    
        Case "ShowJobAddons"
            If SaveData(False) Then Call FAddons.ShowJob(txtJob.Text, 0)
        
        Case "ShowAssemblyAddons"
            If SaveData(False) Then
                Call FAddons.ShowJob(txtJob.Text, mEstAssemblyID)
                
                d = HFApp.SqlExec("select totalamount from addons where basis='Total' and assemblyid=" & DbQuote(Num, mEstAssemblyID))(0)
                txtPrice.Text = format(d, "#,##0.00")
            End If
            
            
        Case Else
            Call HFApp.RunTask(mTimerTask)
    
    End Select
End Sub


Private Sub txtBlock_Change()
    Dirty = True
    
End Sub

Private Sub txtBlock_GotFocus()
    SelectAll txtBlock
End Sub

Private Sub txtChangeApprovedBy_Change()
    mAssemblyChanged = True
    Dirty = True

End Sub

Private Sub txtChangeApprovedBy_GotFocus()
    SelectAll txtChangeApprovedBy
End Sub

Private Sub txtChangeApprovedDate_Change()
    mAssemblyChanged = True
    Dirty = True

End Sub

Private Sub txtChangeApprovedDate_GotFocus()
    SelectAll txtChangeApprovedDate
End Sub

Private Sub txtChangeApprovedDate_Validate(Cancel As Boolean)
    If txtChangeApprovedDate.Text = "" Then Exit Sub
    If IsDate(txtChangeApprovedDate.Text) Then
        txtChangeApprovedDate.Text = format(txtChangeApprovedDate.Text, "medium date")
    Else
        Cancel = True
    End If
End Sub

Private Sub txtChangeRequestedBy_Change()
    mAssemblyChanged = True
    Dirty = True
End Sub

Private Sub txtChangeRequestedBy_GotFocus()
    SelectAll txtChangeRequestedBy
End Sub

Private Sub txtChangeRequestedDate_Change()
    mAssemblyChanged = True
    Dirty = True

End Sub

Private Sub txtChangeRequestedDate_GotFocus()
    SelectAll txtChangeRequestedDate
End Sub

Private Sub txtChangeRequestedDate_Validate(Cancel As Boolean)
    If txtChangeRequestedDate.Text = "" Then Exit Sub
    If IsDate(txtChangeRequestedDate.Text) Then
        txtChangeRequestedDate.Text = format(txtChangeRequestedDate.Text, "medium date")
    Else
        Cancel = True
    End If
End Sub

Private Sub txtCOComments_Change()
    Dirty = True
End Sub
Private Sub txtCOComments_GotFocus()
    SelectAll txtCOComments
End Sub

Private Sub txtCODate_Change()
    Dirty = True
End Sub
Private Sub txtCODate_GotFocus()
    SelectAll txtCODate
End Sub

Private Sub txtCOApprovedBy_Change()
    Dirty = True
End Sub
Private Sub txtCOApprovedBy_GotFocus()
    SelectAll txtCOApprovedBy
End Sub

Private Sub txtCODate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(0)
End Sub

Private Sub txtCODate_Validate(Cancel As Boolean)
    With txtCODate
    If IsDate(.Text) Or Trim(.Text) = "" Then
        .Text = format(.Text, "medium date")
    Else
        Cancel = True
    End If
    End With
End Sub

Private Sub txtCODateApproved_Change()
    Dirty = True
End Sub
Private Sub txtCODateApproved_GotFocus()
    SelectAll txtCODateApproved
End Sub


Private Sub txtCODateApproved_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(1)
End Sub

Private Sub txtCODateApproved_Validate(Cancel As Boolean)
    With txtCODateApproved
    If IsDate(.Text) Or Trim(.Text) = "" Then
        .Text = format(.Text, "medium date")
    Else
        Cancel = True
    End If
    End With
End Sub

Private Sub txtColor_Change()
    Dirty = True
End Sub

Private Sub txtColor_GotFocus()
    SelectAll txtColor
End Sub

Private Sub txtEstimatorNotes_Change()
    Dirty = True
End Sub

Private Sub txtEstimatorNotes_GotFocus()
    SelectAll txtEstimatorNotes
End Sub

Private Sub txtFinish_Change()
    Dirty = True
End Sub

Private Sub txtFinish_GotFocus()
    SelectAll txtFinish
End Sub

Private Sub txtJob_Validate(Cancel As Boolean)
    txtJob.Text = CleanJob(txtJob.Text)
End Sub

Private Sub txtLocation_Change()
    Dirty = True
End Sub

Private Sub txtLocation_GotFocus()
    SelectAll txtLocation
End Sub

Private Sub txtCONumber_GotFocus()
    SelectAll txtCONumber
End Sub

Private Sub txtDateDue_KeyDown(KeyCode As Integer, Shift As Integer)
    mLock = True
End Sub

Private Sub txtDateDue_KeyUp(KeyCode As Integer, Shift As Integer)
    mLock = False
End Sub


Private Sub txtDateIssued_Change()
On Error Resume Next
    If Not mLock Then txtDateIssued.Text = format(txtDateIssued.Text, "mmm d, yyyy")
    Dirty = True
End Sub

Private Sub txtDateIssued_GotFocus()
    SelectAll txtDateIssued
End Sub

Private Sub txtDateIssued_KeyDown(KeyCode As Integer, Shift As Integer)
    mLock = True
End Sub

Private Sub txtDateIssued_KeyUp(KeyCode As Integer, Shift As Integer)
    mLock = False
End Sub

Private Sub txtDateIssued_Validate(Cancel As Boolean)
On Error Resume Next
    If Not IsDate(txtDateIssued.Text) Then
        Cancel = True
    Else
        txtDateIssued.Text = format(txtDateIssued.Text, "mmm d, yyyy")
    End If
End Sub

Private Sub txtDeliveryAddress_GotFocus()
    SelectAll txtDeliveryAddress
End Sub



Private Sub txtDeliveryRecipient_GotFocus()
    SelectAll txtDeliveryRecipient
End Sub


Private Sub txtDescription_Change()
    Dirty = True
End Sub

Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub

Private Sub txtFOB_GotFocus()
    SelectAll txtFOB
End Sub

Private Sub txtHFCategory_Change()
    Dirty = True
    mAssemblyChanged = True
End Sub

Private Sub txtHFCategory_GotFocus()
    SelectAll txtHFCategory
End Sub

Private Sub txtHFComments_GotFocus()
    SelectAll txtHFComments
End Sub

Private Sub txtHFDescription_GotFocus()
    SelectAll txtHFDescription
End Sub

Private Sub txtHFDescription_Validate(Cancel As Boolean)
    Dim r As Long
    With gItems
    For r = 1 To .Rows - 1
        gItems.TextMatrix(r, .ColIndex("HFDescription")) = txtHFDescription.Text
    Next
    End With
End Sub


Private Sub txtJCExtra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(6)
End Sub

Private Sub txtJCExtra_Validate(Cancel As Boolean)
On Error GoTo eh
    Dim s As String
    Dim Description As String
    Dim rs As Recordset
    Dim r As Long
    
    
    If Not mAssemblyChanged Then Exit Sub
    
    If mMode <> fceQuote Then
        Select Case HFApp.Options(AccountingSystem)
            Case asTimberline
                If Trim(txtJCExtra.Text) <> "" Then
                    Set rs = HFApp.SqlExec("select extra from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(mJob)) & " AND extra=" & DbQuote(Str, txtJCExtra), dbAccounting)
                    If rs.EOF Then
                        Description = Left(txtHFDescription.Text, 30)
                        Description = InputBox(vbCrLf & "Extra '" & txtJCExtra.Text & "' could not be found." & vbCrLf & "Do you want to add it?" & vbCrLf & vbCrLf & vbCrLf & "Enter a description for the extra. (30 characters or less)", App.ProductName, Description)
                        If Description = "" Then
                            Cancel = True
                        Else
                            s = ""
                            s = s & "INSERT INTO jcm_master__extra(Job,Extra,Description,Status)" & vbCrLf
                            s = s & "VALUES(" & DbQuote(Str, HFApp.FormatJob(mJob)) & vbCrLf
                            s = s & "      ," & DbQuote(Str, txtJCExtra.Text) & vbCrLf
                            s = s & "      ," & DbQuote(Str, Left(Description, 30)) & vbCrLf
                            s = s & "      ,'In progress')" & vbCrLf
                            Set rs = HFApp.SqlExec(s, dbAccounting)
                        End If
                    Else
                        txtJCExtra.Text = "" & rs(0)
                    End If
                End If
                
            Case asMasterBuilder
                If txtJCExtra.Text <> "" And Not IsNumeric(txtJCExtra.Text) Then
                    MsgBox "Phase must be numeric", vbExclamation, App.ProductName
                    Call SelectAll(txtJCExtra)
                    Cancel = True
                End If
            
                Set rs = HFApp.SqlExec("select phsnum from jobphs where recnum=" & DbQuote(Num, mJob) & " AND phsnum=" & DbQuote(Num, txtJCExtra), dbAccounting)
                If rs.EOF Then
                    Description = Left(txtHFDescription.Text, 30)
                    Description = InputBox(vbCrLf & "Phase '" & txtJCExtra.Text & "' could not be found." & vbCrLf & "Do you want to add it?" & vbCrLf & vbCrLf & vbCrLf & "Enter a description for the phase. (30 characters or less)", App.ProductName, Description)
                    If Description = "" Then
                        Cancel = True
                    Else
                        s = HFApp.XmlMbStart(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
                        s = s & "<JobModRq requestID=""1"">" & vbCrLf
                        s = s & HFApp.XmlMBAdd(n, 10, "ObjectRef", mJob)
                        s = s & "<PhaseAdd>" & vbCrLf
                        s = s & "<ObjectRef>" & HFApp.XmlMBAdd(n, 10, "PhaseID", txtJCExtra.Text) & "</ObjectRef>" & vbCrLf
                        s = s & HFApp.XmlMBAdd(c, 30, "Desc", Description)
                        s = s & HFApp.XmlMBAdd(c, 10, "Unit", lblUOM.Caption)
                        s = s & HFApp.XmlMBAdd(n, 12.2, "Quantity", lblQuantity.Caption)
                        s = s & HFApp.XmlMBAdd(m, 10, "Memo", txtHFComments.Text)
                        s = s & "</PhaseAdd>" & vbCrLf
                        s = s & "</JobModRq>" & vbCrLf
                        s = s & HFApp.XmlMBEnd()
                        Call HFApp.XmlMbSubmit(s, HFApp.Options(MasterBuilderPWD))
                    End If
                Else
                    txtJCExtra.Text = "" & rs(0)
                End If
    
                
        End Select
    End If
    
    
    If Not Cancel Then
        Dirty = True
        With gItems
        For r = 1 To .Rows - 1
            If .ValueMatrix(r, .ColIndex("BudgetGenerated")) = 0 And Trim(.TextMatrix(r, .ColIndex("PONumber"))) = "" Then
                gItems.TextMatrix(r, .ColIndex("JCExtra")) = txtJCExtra.Text
                .RowData(r) = "DIRTY"
            End If
        Next
        End With
    End If
Exit Sub
eh: Call errHandler(SRCFILE & "txtJCExtra_Validate", s)
End Sub

Public Sub ShowForm(Mode As FCEMode, Job As String)
    
    'ProjectBased = False
    mMode = Mode
    mJob = Job
    If Mode = fceQuote Then ProjectBased = False
    If ProjectBased And Mode <> fceQuote And mJob <> "" Then
        mCommunity = "" & HFApp.SqlExec("Select Community from tblJobs where DivisionID = " & HFApp.DivisionID & " and job_no = " & DbQuote(Str, mJob))(0)
        
    End If
    Me.Show
    Me.Tag = Mode & Chr(0) & "" & mJob
    
End Sub

Private Sub chkHidePrice_Click()
    Dirty = True
End Sub
Private Sub chkIncludeDocuments_Click()
    Dirty = True
End Sub

Private Sub chkHideQty_Click()
    Dirty = True
End Sub

Private Sub chkTotalOnly_Click()
    Dirty = True
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim s As String
    Dim t As String
    Dim rs As Recordset
    Dim c As Control
    Dim b As Boolean
    
    
    'configure form
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Me.Caption = Choose(mMode, "Prepare Quote", "Issue Budgets", "Issue PO's", Me.Caption, Me.Caption, Me.Caption)
    VarianceCatDef = HFApp.Options.ValueByName("DefaultVarianceCategory")
    
    
    
    'change to normal colors
    Me.BackColor = vbButtonFace
    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            c.BackColor = vbButtonFace
        End If
        If TypeName(c) = "Slider" Then
            c.Visible = True
        End If
    Next
        
    Call LoadCustomDescriptions

    'get conditional formatting
    AcceptedForeColor = IniGet(AppIni, "Formatting", "AcceptedForeColor", vbWindowText)
    AcceptedBackColor = IniGet(AppIni, "Formatting", "AcceptedBackColor", vbWindowBackground)
    AcceptedStyle = IniGet(AppIni, "Formatting", "AcceptedStyle", "Bold Italic")
    DeclinedForeColor = IniGet(AppIni, "Formatting", "DeclinedForeColor", vbGrayText)
    DeclinedBackColor = IniGet(AppIni, "Formatting", "DeclinedBackColor", vbButtonFace)
    DeclinedStyle = IniGet(AppIni, "Formatting", "DeclinedStyle", "Regular")
    MinForeColor = IniGet(AppIni, "Formatting", "MinForeColor", vbWindowText)
    MinBackColor = IniGet(AppIni, "Formatting", "MinBackColor", 16773869) 'light blue
    MinStyle = IniGet(AppIni, "Formatting", "MinStyle", "Regular")
    MaxForeColor = IniGet(AppIni, "Formatting", "MaxForeColor", vbWindowText)
    MaxBackColor = IniGet(AppIni, "Formatting", "MaxBackColor", 12640511) 'light pink
    MaxStyle = IniGet(AppIni, "Formatting", "MaxStyle", "Regular")
        
    mUseChangeRequests = HFApp.Options.ValueByName("UseChangeRequests") <> "False"
    With cboChangeRequestStatus
        .Clear
        .AddItem "Pending"
        s = HFApp.Options.ValueByName("ChangeRequestStatuses")
        For i = 1 To Parse(s, , "|")
            t = Trim(Parse(s, i, "|"))
            If t <> "" Then .AddItem t
        Next
    End With
        
        
    Toolbar.Buttons("TakeoffPipeline").Visible = HFApp.Options.ValueByName("EstimatingSystem") = esPipeline
    Toolbar.Buttons("TakeoffPlanSwift").Visible = TakeoffSystem = tsPlanSwift
    
    chkCancelledItems.Visible = PostPOQtyToAccounting
    
        
    If HFApp.Options(MultiFamily) Then
        cboModel.Visible = False
        Label1120.Visible = False
        lblModelDesc.Visible = False
    End If
        
    
    Select Case HFApp.Options(AccountingSystem)
        'change terminology - extra to phase
        Case asMasterBuilder:  lblAssemblyExtra.Caption = "Asmbly/Phase"
    End Select
    
    
    Select Case mMode
        Case fceQuote
            lblJobNumber.Caption = "Quote Number"
            Toolbar.Buttons("Save").Style = tbrDropdown
                        
        Case Else
            lblJobNumber.Caption = "Job Number"
            Toolbar.Buttons("Save").Style = tbrDefault
    
            
    End Select
    
    
    
    
    
    cmdBrowse(6).Visible = HFApp.Options(AccountingSystem) <> AccountingSystems.asNone
    gItems.Rows = 1
    
    On Error Resume Next
    Call IniGetForm(Me)
    chkToolbarCaptions.value = IIf(CBool(Parse(Toolbar.Tag, 2)), vbChecked, vbUnchecked)
    chkCancelledItems.value = HFApp.Options.ValueByName("CancelledItems")
    On Error GoTo 0
    
    With gItems
        
        Select Case mMode
            Case fceQuote
                Toolbar.Buttons("SendPOs").Visible = False
                Toolbar.Buttons("Preview").Visible = False
                Toolbar.Buttons("ViewBudgets").Visible = False
                Toolbar.Buttons("ViewPOs").Visible = False
                Toolbar.Buttons("s4").Visible = False
                Toolbar.Buttons("Delete").Visible = True
                Toolbar.Buttons("SaveAssembly").Visible = True
                Toolbar.Buttons("RePrice").Visible = True
                Toolbar.Buttons("Generate").Visible = False
                .TextMatrix(0, .ColIndex("POVendor")) = ""
                .TextMatrix(0, .ColIndex("POVendorName")) = ""
                .TextMatrix(0, .ColIndex("POQty")) = ""
                .TextMatrix(0, .ColIndex("PORate")) = ""
                .TextMatrix(0, .ColIndex("POPretax")) = ""
                .TextMatrix(0, .ColIndex("POTaxGroup")) = ""
                .TextMatrix(0, .ColIndex("POTax")) = ""
                .TextMatrix(0, .ColIndex("POJCTax")) = ""
                .TextMatrix(0, .ColIndex("PONJCTax")) = ""
                .TextMatrix(0, .ColIndex("POTotal")) = ""
            
            Case fceBudget
                Toolbar.Buttons("Delete").Visible = False
                Toolbar.Buttons("SendPOs").Visible = False
                Toolbar.Buttons("Preview").Visible = False
                Toolbar.Buttons("ViewBudgets").Visible = False
                Toolbar.Buttons("ViewPOs").Visible = True
                Toolbar.Buttons("s4").Visible = True
                Toolbar.Buttons("SaveAssembly").Visible = True
                Toolbar.Buttons("RePrice").Visible = True
                Toolbar.Buttons("Generate").Visible = True
                .TextMatrix(0, .ColIndex("POVendor")) = ""
                .TextMatrix(0, .ColIndex("POVendorName")) = ""
                .TextMatrix(0, .ColIndex("POQty")) = ""
                .TextMatrix(0, .ColIndex("PORate")) = ""
                .TextMatrix(0, .ColIndex("POPretax")) = ""
                .TextMatrix(0, .ColIndex("POTaxGroup")) = ""
                .TextMatrix(0, .ColIndex("POTax")) = ""
                .TextMatrix(0, .ColIndex("POJCTax")) = ""
                .TextMatrix(0, .ColIndex("PONJCTax")) = ""
                .TextMatrix(0, .ColIndex("POTotal")) = ""
                
            Case fcePO
                If HFApp.UserPermission("IssuePOs") = False Then
                    Toolbar.Buttons("TakeoffOneTime").Visible = False
                    Toolbar.Buttons("TakeoffItem").Visible = False
                    Toolbar.Buttons("TakeoffAssembly").Visible = False
                    Toolbar.Buttons("TakeoffCustom").Visible = False
                    Toolbar.Buttons("ViewBudgets").Visible = False
                    Toolbar.Buttons("RePrice").Visible = False
                    Toolbar.Buttons("NewRFQ").Visible = False
                End If
                Toolbar.Buttons("Delete").Visible = False
                Toolbar.Buttons("ViewPOs").Visible = False
                Toolbar.Buttons("SendPOs").Visible = HFApp.UserPermission("SendPO") And HFApp.Options.ValueByName("BuildProCompanyCode") = ""
                .TextMatrix(0, .ColIndex("BudgetVendor")) = ""
                .TextMatrix(0, .ColIndex("BudgetVendorName")) = ""
                .TextMatrix(0, .ColIndex("BudgetQty")) = ""
                .TextMatrix(0, .ColIndex("BudgetRate")) = ""
                .TextMatrix(0, .ColIndex("BudgetPretax")) = ""
                .TextMatrix(0, .ColIndex("BudgetTaxGroup")) = ""
                .TextMatrix(0, .ColIndex("BudgetTax")) = ""
                .TextMatrix(0, .ColIndex("BudgetJCTax")) = ""
                .TextMatrix(0, .ColIndex("BudgetNJCTax")) = ""
                .TextMatrix(0, .ColIndex("BudgetTotal")) = ""
        End Select
        
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "" Then .ColHidden(i) = True
        Next
        
        
    End With
    Call LoadViews
    mViewIndex = Val(IniGet(AppIni, "Options", "EstimateItemsView" & mMode))
    Select Case mMode
        Case fceContract: mViewIndex = ViewIndex("Contract")
        Case fceBilling:  mViewIndex = ViewIndex("Contract")
        Case fceBids:     mViewIndex = ViewIndex("RFQ's")
    End Select
    
    
    Call ClearGroups
    Call IniGetGrid(Me, gItems, , , mMode & mViewIndex, True)
    Call IniGetGrid(Me, gBids)
    Call IniGetGrid(Me, gAssemblies)
    Call IniGetGrid(Me, gContract)
    Call IniGetGrid(Me, gInvoices)
    Call IniGetGrid(Me, gBidItems)
    Call mnuEstimateItemViewsSub_Click(mViewIndex + 1)
    
    If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
        Label1120.Caption = "Template"
    Else
        Label1120.Caption = "Model"
    End If
    
    frmJob.Visible = IniGet(AppIni, "Options", "EstimateItemsView.JobInfo." & mMode, False)
    If mMode = fceQuote Then frmJob.Visible = True
    gItems.ColHidden(gItems.ColIndex("Selected")) = False
    gItems.ColHidden(gItems.ColIndex("WarningMessages")) = False
    Me.Show
    
    'load static picklists
    Call LoadComboBox(cboCommunity, HFApp.Databases(dbHomefront), "SELECT c.Description,c.Area,0 FROM tblLocality c join divisioncommunities d on c.Area = d.Community where d.DivisionID = " & HFApp.DivisionID & " ORDER BY 1")
    Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), "select distinct isnull(model,'') ,'',0 from tbldbassemblymaster where DivisionID = " & HFApp.DivisionID & " and isnull(inactive,0)=0 and assemblytype=0")
    
    
    'enabled/disable snapshots
    Toolbar.Buttons("snapshots").Visible = mMode = fceBudget And (InIde() Or HFApp.Options.ValueByName("BudgetSnapshotsEnabled") = "True")
    With gItems
    If HFApp.Options.ValueByName("BudgetSnapshotsEnabled") = "True" Then
        .TextMatrix(0, .ColIndex("Budget1Rate")) = "Initial Budget Rate"
        .TextMatrix(0, .ColIndex("Budget1Qty")) = "Initial Budget Qty"
        .TextMatrix(0, .ColIndex("Budget2Rate")) = "Confirmed Budget Rate"
        .TextMatrix(0, .ColIndex("Budget2Qty")) = "Confirmed Budget Qty"
        .TextMatrix(0, .ColIndex("Budget3Rate")) = "Modified Budget Rate"
        .TextMatrix(0, .ColIndex("Budget3Qty")) = "Modified Budget Qty"
    Else
        .TextMatrix(0, .ColIndex("Budget1Rate")) = ""
        .TextMatrix(0, .ColIndex("Budget1Qty")) = ""
        .TextMatrix(0, .ColIndex("Budget2Rate")) = ""
        .TextMatrix(0, .ColIndex("Budget2Qty")) = ""
        .TextMatrix(0, .ColIndex("Budget3Rate")) = ""
        .TextMatrix(0, .ColIndex("Budget3Qty")) = ""
        
        .ColHidden(.ColIndex("Budget1Rate")) = True
        .ColHidden(.ColIndex("Budget1Qty")) = True
        .ColHidden(.ColIndex("Budget2Rate")) = True
        .ColHidden(.ColIndex("Budget2Qty")) = True
        .ColHidden(.ColIndex("Budget3Rate")) = True
        .ColHidden(.ColIndex("Budget3Qty")) = True
    End If
    End With
    
    Me.Caption = Choose(mMode, "Prepare Quote", "Issue Budgets", "Issue PO's") & " - " & mJob
    Call LoadJob(mJob)
    Call GroupGrid
    
End Sub


Public Sub mnuEstimateItemViewsSub_Click(Index As Integer)
    If Not SaveData(True) Then Exit Sub
    
    
    Select Case mViews(Index - 1).Name
        Case "Addons"
            mTimerTask = "ShowJobAddons"
            Timer1.Interval = 10
            Timer1.Enabled = True
        
        Case "Contract"
            mViewIndex = Index - 1
            frmBilling.Visible = True
            frmPurchasing.Visible = False
            Call LoadContract
        
        Case Else
            frmBilling.Visible = False
            frmPurchasing.Visible = True
            Call IniPutGrid(Me, gItems, , mMode & mViewIndex)
            
            
            mViewIndex = Index - 1
            
            Call ClearGroups
            Call IniGetGrid(Me, gItems, , , mMode & mViewIndex, True)
            gItems.ColHidden(gItems.ColIndex("Selected")) = False
            gItems.ColHidden(gItems.ColIndex("WarningMessages")) = False
            
            gItems.OLEDragMode = IIf(mViewIndex = 0, flexOLEDragAutomatic, flexOLEDragManual)
            If ReadOnly Then gItems.OLEDragMode = flexOLEDragManual
            
            frmRFPs.Visible = mViews(mViewIndex).Name = "RFQ's"
            frmEstItems.Visible = mViews(mViewIndex).Name <> "RFQ's"
            Slider1.Visible = True
            Call Clear
            Call LoadAssemblies(True)
    
    End Select
    
End Sub

Private Sub CancelBudgets()
On Error GoTo eh
    Dim List As String
    Dim s As String
    Dim r As Long
    
    If Not SaveData(True) Then Exit Sub
    
    With gItems
        List = ""
        For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            If Not .IsSubtotal(r) Then
                List = List & "," & .TextMatrix(r, .ColIndex("EstItemID"))
            End If
        Next
        List = Mid(List, 2)
        
        If List = "" Then Exit Sub
        
        s = ""
        s = s & "INSERT INTO CancelledBudgets(EstItemID,DivisionID,Job,JCExtra,JCCostCode,JCCategory,Qty,UOM,Pretax,Tax,TaxGroup,OrigBatch,PostBatch,ChangeOrder,UStmp,TStmp)" & vbCrLf
        s = s & "SELECT i.EstItemID" & vbCrLf
        s = s & "      ," & HFApp.DivisionID
        s = s & "      ,i.Job" & vbCrLf
        s = s & "      ,i.JCExtra" & vbCrLf
        s = s & "      ,i.JCCostCode" & vbCrLf
        s = s & "      ,i.JCCategory" & vbCrLf
        s = s & "      ,i.BudgetQty*-1" & vbCrLf
        s = s & "      ,i.OrderUOM" & vbCrLf
        s = s & "      ,i.BudgetPretax*-1" & vbCrLf
        s = s & "      ,i.BudgetJCTax*-1" & vbCrLf
        s = s & "      ,i.BudgetTaxGroup" & vbCrLf
        s = s & "      ,i.BudgetPostingBatch OrigBatch" & vbCrLf
        s = s & "      ,0 PostBatch" & vbCrLf
        s = s & "      ,a.ChangeOrder" & vbCrLf
        s = s & "      ,'ADMIN' UStmp" & vbCrLf
        s = s & "      ,GETDATE() TStmp" & vbCrLf
        s = s & "  FROM EstimateItems i" & vbCrLf
        s = s & "  LEFT OUTER JOIN estimateassemblies a on i.estassemblyid=a.estassemblyid" & vbCrLf
        s = s & " WHERE i.BudgetDeleted=0" & vbCrLf
        s = s & "   AND i.BudgetPostingBatch<>0" & vbCrLf
        s = s & "   AND i.EstItemID IN(" & List & ")"
        Call HFApp.SqlExec(s)
        
        s = ""
        s = s & "UPDATE EstimateItems" & vbCrLf
        s = s & "   SET BudgetGenerated=0" & vbCrLf
        s = s & "      ,BudgetPostingBatch=0" & vbCrLf
        s = s & " WHERE BudgetDeleted=0" & vbCrLf
        s = s & "   AND EstItemID IN(" & List & ")"
        Call HFApp.SqlExec(s)
        
        For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            .TextMatrix(r, .ColIndex("BudgetPostingBatch")) = "0"
            .TextMatrix(r, .ColIndex("BudgetGenerated")) = "False"
            Call ColorizeItems(r)
        Next
        
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "CancelBudgets", s)
End Sub

Private Sub GenerateBudgets()
    If Not SaveData(True) Then Exit Sub
    
    'Moved to bottom of the procedure because it was failing for Jorgensen saying there were two customers with the same name, then it would not continue to the rest of the code. The customers were not duplicated in simply either
    'If HFApp.Options(AccountingSystem) = asSimply Then Call HFApp.WriteJobToAccounting(mJob)
    
    Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetGenerated=1 WHERE BudgetDeleted=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems i WHERE " & WhereClause & ")")
    Call ClearWhereClause
    Dirty = False
    Call RefreshTree
    Call LoadItems
    If HFApp.Options(AccountingSystem) = asSimply Then Call HFApp.WriteJobToAccounting(mJob)
    
End Sub

Private Sub Toolbar_ButtonDropDown(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim i As Long
    Dim ParentMenu As Long
    
    Select Case Button.Key
        Case "snapshots"
            Call ShowSnapshotsMenu(Button)
            
        Case "View"
            With FMain.PopMenu
                ParentMenu = .MenuIndex("mnuEstimateItemViews")
                Call .ClearSubMenusOfItem(ParentMenu)
                For i = 0 To UBound(mViews)
                    .AddItem mViews(i).Name, "EstimateItemViews" & i, , i, ParentMenu, , i = mViewIndex
                Next
            End With
            PopupMenu FMain.mnuEstimateItemViews, , Button.Left, Button.Top + Button.Height
        
        Case "Open"

    End Select
    Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonDropDown")
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim bRePrice As Boolean
    Dim i As Long
    Dim rs As Recordset
    Dim s As String
    Dim SalesPrice As Double
    Dim r As Long
    Dim lastRow As Long
    Dim f As FEstimateItems
    Dim CostBasis As CostBasisTypes
    Dim EffectiveDate As Date
    Dim ApplyVarianceCat As Boolean
    Dim RefreshOverridden As Boolean

    Select Case Button.Key
        
        Case "Delete":          Call DeleteQuote
        
        Case "SendPOs":         Call SendingWizard("PO", mJob, SelectedPOs)
        Case "SendRFQs":        Call SendingWizard("RFQ", mJob)
                                
                                
        Case "Preview":         Call ShowSelectedPOs(True)
        Case "SaveAssembly":    Call SaveDBAssembly
        Case "Save":            Call SaveData(False)
        Case "Open":            Call LoadJob("")
        Case "View":            Call Toolbar_ButtonDropDown(Button)
        
        Case "snapshots"
            Call ShowSnapshotsMenu(Button)
            
        Case "NewRFQ"
            Call SaveData(False)
            Call FRFPWizard.ShowForm(mJob)
            Call mnuEstimateItemViewsSub_Click(1 + mViewIndex)
        
        Case "takeoffsettings":
            If mTakeoffSettings <> "" Then
                mTakeoffSettings = ""
                Button.value = tbrUnpressed
            Else
                mTakeoffSettings = FTakeoffSettings.GetSettings(mJob, mTakeoffSettings)
                Button.value = IIf(mTakeoffSettings = "", tbrUnpressed, tbrPressed)
            End If


        Case "ViewPOs", "ViewBudgets"
            If SaveData(True) Then
                'Set f = FindForm("FEstimateItems", IIf(mMode = fcepo, fceBudget, fcepo) & Chr(0) & mJob)
                'If f Is Nothing Then
                '    Set f = New FEstimateItems
                '    Call f.ShowForm(IIf(mMode = fcepo, fceBudget, fcepo), mJob)
                'Else
                '    If f.WindowState = vbMinimized Then f.WindowState = vbNormal
                '    f.SetFocus
                'End If
                
                
                Unload Me
                
                Set f = New FEstimateItems
                Call f.ShowForm(IIf(mMode = fcePO, fceBudget, fcePO), mJob)
                
            End If
            
        Case "TakeoffPipeline"
            Call ReadFromPipeline
            
        Case "TakeoffOneTime", "TakeoffItem", "TakeoffAssembly", "TakeoffCustom", "TakeoffPlanswift"
    
            If Not frmAssembly.Visible Then
                If ProjectBased Then
                    s = "SELECT COUNT(DISTINCT EstAssemblyID) FROM estimateassemblies join tbljobs on estimateassemblies.job = tbljobs.job_no and estimateassemblies.DivisionID = tbljobs.DivisionID WHERE isquote=0 and community=" & DbQuote(Str, mCommunity) & " and tbljobs.DivisionID = " & HFApp.DivisionID
                Else
                    s = "SELECT COUNT(DISTINCT EstAssemblyID) FROM estimateassemblies WHERE Job=" & DbQuote(Str, mJob) & " and DivisionID = " & HFApp.DivisionID
                End If
                i = Val("" & HFApp.SqlExec(s)(0))
                If ProjectBased Then
                    If i = 1 Then
                        s = "SELECT Job,tbljobs.Description,EstAssemblyID,HFDescription FROM estimateassemblies join tbljobs on estimateassemblies.job = tbljobs.job_no and estimateassemblies.DivisionID = tbljobs.DivisionID  WHERE isquote=0 and Community=" & DbQuote(Str, mCommunity) & " and tbljobs.DivisionID = " & HFApp.DivisionID
                        Set rs = HFApp.SqlExec(s)
                        mEstAssemblyID = Val("" & rs(2))
                        mCurrentAssemblyDesc = "" & rs(3)
                    Else
                        s = ""
                        s = s & "select c.Job_no,c.description Customer,a.ChangeOrder,a.hfDescription Description,a.EstAssemblyID,a.budgetslocked" & vbCrLf
                        s = s & "  from estimateassemblies a" & vbCrLf
                        s = s & "  join tblcustomers c on(a.customer_no=c.customer_no)" & vbCrLf
                        s = s & " where a.DivisionID = " & HFApp.DivisionID & " and c.job_no<>'' and Community=" & DbQuote(Str, mCommunity) & vbCrLf
                        If mMode = fceBudget Then
                            s = s & "   and isnull(budgetslocked,0)=0" & vbCrLf
                        End If
                        s = s & "order by 1,2,3,4" & vbCrLf
                        
                        If FPickList.Choose(HFApp.Databases(dbHomefront), "Assembly", s, , , , , "EstAssemblyID" & IIf(HFApp.Options(MultiFamily), "", ",Customer")) Then
                            mCurrentAssemblyDesc = FPickList.SelectedItem(4)
                            mEstAssemblyID = Val(FPickList.SelectedItem(5))
                        Else
                            mEstAssemblyID = -1
                        End If
                    End If
                
                Else
                    If i = 1 Then
                        s = "SELECT EstAssemblyID,HFDescription FROM estimateassemblies WHERE Job=" & DbQuote(Str, mJob) & " and DivisionID = " & HFApp.DivisionID
                        Set rs = HFApp.SqlExec(s)
                        mEstAssemblyID = Val("" & rs(0))
                        mCurrentAssemblyDesc = "" & rs(1)
                    Else
                        s = ""
                        s = s & "select c.description Customer,a.ChangeOrder,a.hfDescription Description,a.EstAssemblyID,a.budgetslocked" & vbCrLf
                        s = s & "  from estimateassemblies a" & vbCrLf
                        s = s & "  join tblcustomers c on(a.customer_no=c.customer_no)" & vbCrLf
                        s = s & " where job=" & DbQuote(Str, mJob) & " and a.DivisionID = " & HFApp.DivisionID & vbCrLf
                        If mMode = fceBudget Then
                            s = s & "   and isnull(budgetslocked,0)=0" & vbCrLf
                        End If
                        s = s & "order by 1,2,3" & vbCrLf
                        
                        If FPickList.Choose(HFApp.Databases(dbHomefront), "Assembly", s, , , , , "EstAssemblyID" & IIf(HFApp.Options(MultiFamily), "", ",Customer")) Then
                            mCurrentAssemblyDesc = FPickList.SelectedItem(3)
                            mEstAssemblyID = Val(FPickList.SelectedItem(4))
                        Else
                            mEstAssemblyID = -1
                        End If
                    End If
                End If
            End If
            
            If mEstAssemblyID <> -1 Then
                mCurrentEstAssemblyID = mEstAssemblyID
                
                s = ""
                s = s & "SELECT j.community,j.communityphase,a.assembly,j.job_no,a.jcextra,j.description jobdesc,isnull(nullif(a.model,''),j.model) model,a.optionid,a.hfdescription,a.assemblytype,a.budgetslocked,a.changeorder" & vbCrLf
                s = s & "FROM Estimateassemblies a JOIN tblJobs j on(a.job=j.job_no and a.DivisionID = j.DivisionID)" & vbCrLf
                s = s & "WHERE EstAssemblyID=" & DbQuote(Num, mCurrentEstAssemblyID) & " and j.DivisionID = " & HFApp.DivisionID
                Set rs = HFApp.SqlExec(s)
                
                mCurrentAssemblyDesc = "" & rs("HFDescription")
                mCurrentJobDesc = "" & rs("JobDesc")
                mCurrentModel = "" & rs("Model")
                mCurrentOptionID = "" & rs("OptionID")
                mCurrentJob = "" & rs("Job_No")
                mCurrentCommunity = "" & rs("Community")
                mCurrentCommunityPhase = "" & rs("CommunityPhase")
                mCurrentAssembly = "" & rs("Assembly")
                mCurrentJCExtra = "" & rs("JCExtra")
                mCurrentChangeOrder = "" & rs("ChangeOrder")
                mCurrentBudgetsLocked = "" & rs("BudgetsLocked") = "True"
                                
                s = "Call FTakeoff.Takeoff"
                mCurrentAssemblyType = Val("" & rs("AssemblyType"))
                
                'save assembly,model,option back to EstimateAssembly if it doesnt have one
                Dim bReqSave As Boolean
                bReqSave = mCurrentAssembly = ""
                mCancelTakeoff = True
                'remember lastrow so you can find new items for variance reporting
                lastRow = gItems.Rows - 1
                Call FTakeoff.Takeoff(Me, _
                                      mCurrentChangeOrder <> "", _
                                      Mid(Button.Key, 8), _
                                      mTakeoffSystemBidID, _
                                      mTakeoffSystemProjectName, _
                                      mCurrentAssemblyType, _
                                      mCurrentAssemblyDesc, _
                                      mCurrentCommunity, _
                                      mCurrentCommunityPhase, _
                                      mCurrentModel, _
                                      mCurrentOptionID, _
                                      mCurrentAssembly, _
                                      mCurrentJob)
                
                'enforce variance cat
                With gItems
                If mCurrentBudgetsLocked Then
                    If HFApp.Options.ValueByName("UseVarianceReporting") = "False" Then
                        For r = lastRow To .Rows - 1
                            .TextMatrix(r, .ColIndex("VarianceJCCategory")) = .TextMatrix(r, .ColIndex("JCCategory"))
                            .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = .TextMatrix(r, .ColIndex("JCCategoryDesc"))
                        Next
                    Else
                        If FPickList.Choose(HFApp.Databases(dbHomefront), "Variance Category", "Select [Category], [Description] from StandardCategories where isvariance=1 and DivisionID =" & HFApp.DivisionID, VarianceCatDef) Then
                            For r = lastRow + 1 To .Rows - 1
                                .TextMatrix(r, .ColIndex("VarianceJCCategory")) = FPickList.SelectedItem("Category")
                                .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = FPickList.SelectedItem("Description")
                            Next
                        End If
                    End If
                End If
                End With
                
                
                If bReqSave And mCurrentAssembly <> "" And Not mCancelTakeoff Then
                    
                    If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
                        If txtPrice.Text = "0.00" Then
                            On Error Resume Next
                            Select Case mCurrentAssemblyType
                            Case 0
                                 s = "Select top 1 isnull(base_house,0) from tblmodels where DivisionID = " & HFApp.DivisionID & " and model = " & DbQuote(Str, mCurrentModel) & " and Area in (" & DbQuote(Str, mCurrentCommunity) & ",'')" & " and CommunityPhase in (" & DbQuote(Str, mCurrentCommunityPhase) & ",'') order by Area desc,CommunityPhase desc"
                                 SalesPrice = HFApp.SqlExec(s, dbHomefront)(0)
                            Case 2
                                 s = "Select top 1 Price from tblOptions where DivisionID = " & HFApp.DivisionID & " and model = " & DbQuote(Str, mCurrentModel) & " and Opt = " & DbQuote(Str, mCurrentOptionID) & " and Area in (" & DbQuote(Str, mCurrentCommunity) & ",'')" & " and CommunityPhase in (" & DbQuote(Str, mCurrentCommunityPhase) & ",'') order by Area desc,CommunityPhase desc"
                                 SalesPrice = HFApp.SqlExec(s, dbHomefront)(0)
                            Case 3
                                 s = "Select top 1 Price from tblGlobalOptions where DivisionID = " & HFApp.DivisionID & " and Opt = " & DbQuote(Str, mCurrentOptionID) & " and Community in (" & DbQuote(Str, mCurrentCommunity) & ",'')" & " and CommunityPhase in (" & DbQuote(Str, mCurrentCommunityPhase) & ",'') order by Community desc,CommunityPhase desc"
                                 SalesPrice = HFApp.SqlExec(s, dbHomefront)(0)
                            Case 4
                            End Select
                            On Error GoTo eh
                            If SalesPrice <> 0 Then txtPrice.Text = format(SalesPrice, "##,#.00")
                        End If
                    End If
                
                    s = ""
                    s = s & "UPDATE EstimateAssemblies" & vbCrLf
                    s = s & "   SET AssemblyType=" & DbQuote(Num, mCurrentAssemblyType) & vbCrLf
                    s = s & "      ,Assembly=case isnull(Assembly,'') when '' then " & DbQuote(Str, mCurrentAssembly) & " else Assembly end" & vbCrLf
                    s = s & "      ,HFDescription=case isnull(hfdescription,'')" & vbCrLf
                    s = s & "           when 'Manual Estimates' then " & DbQuote(Str, mCurrentAssemblyDesc) & vbCrLf
                    s = s & "           when 'untitled' then " & DbQuote(Str, mCurrentAssemblyDesc) & vbCrLf
                    s = s & "           when '' then " & DbQuote(Str, mCurrentAssemblyDesc) & vbCrLf
                    s = s & "           else HFDescription end" & vbCrLf
                    If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
                        s = s & "      ,SalesRate= case when isnull(SalesRate,0) = 0 then " & DbQuote(Num, SalesPrice) & " else SalesRate end" & vbCrLf
                    End If
                    s = s & "      ,Model=" & DbQuote(Str, mCurrentModel) & vbCrLf
                    s = s & "      ,OptionID=case isnull(OptionID,'') when '' then " & DbQuote(Str, mCurrentOptionID) & " else OptionID end" & vbCrLf
                    
                    s = s & "WHERE EstAssemblyID=" & DbQuote(Num, mCurrentEstAssemblyID)
                    Set rs = HFApp.SqlExec(s)
                    
                    s = ""
                    s = s & "UPDATE EstimateItems" & vbCrLf
                    s = s & "   SET Assembly=" & DbQuote(Str, mCurrentAssembly) & vbCrLf
                    s = s & "      ,Model=" & DbQuote(Str, mCurrentModel) & vbCrLf
                    s = s & "WHERE EstAssemblyID=" & DbQuote(Num, mCurrentEstAssemblyID)
                    Set rs = HFApp.SqlExec(s)
                    
                    'add model and assembly to all items in the list
                    With gItems
                        For i = 1 To .Rows - 1
                            If .ValueMatrix(i, .ColIndex("EstAssemblyID")) = mCurrentEstAssemblyID Then
                                .TextMatrix(i, .ColIndex("Assembly")) = mCurrentAssembly
                                .TextMatrix(i, .ColIndex("Model")) = mCurrentModel
                                .TextMatrix(i, .ColIndex("OptionID")) = mCurrentOptionID
                            End If
                        Next
                    End With
                    
                    'change descriptions
                    If txtHFOption.Text = "" And IsIn(txtHFDescription.Text, "Manual Estimates", "untitled", "") Then
                        txtHFDescription.Text = mCurrentAssemblyDesc
                        With gAssemblies
                            For i = 1 To .Rows - 1
                                If .Cell(flexcpData, i, 1) = "EstAssemblyID" And .Cell(flexcpText, i, 1) = mCurrentEstAssemblyID Then
                                    .Cell(flexcpText, i, 0) = txtHFDescription.Text
                                End If
                            Next
                        End With
                    End If
                    If txtHFOption.Text = "" Then txtHFOption.Text = mCurrentAssembly
                    
                End If
                
                Call GroupGrid
                Call ColorizeItems(-1)
            End If
            
        Case "Generate"
        
            
            If mMode = fcePO Then
                'remove genbatch from any records that are generated but still have no ponumber... something went wrong.
                s = ""
                s = s & "update estimateitems set pogenbatch=0 where estitemid in(" & vbCrLf
                s = s & "SELECT i.estitemid FROM EstimatedItems i WHERE pogenbatch<>0 AND ponumber='' AND " & WhereClause & vbCrLf
                s = s & ")"
                Call HFApp.SqlExec(s)
            End If
            
            s = "SELECT COUNT(*) FROM EstimatedItems i WHERE " & IIf(mMode = fceBudget, "isnull(budgetgenerated,0)=0", "isnull(pogenbatch,0)=0") & " AND " & WhereClause
            If 0 = HFApp.SqlExec(s)(0) Then
                MsgBox "You have nothing selected that can be generated." & vbCrLf & _
                       "Use the check boxes to select some items that have" & vbCrLf & _
                       "not already been generated.", vbExclamation, App.ProductName
                Exit Sub
            End If
            
            If Not SaveData(False) Then Exit Sub
            If Not FGenerate.ValidateData(WhereClause, mMode = fceBudget) Then Exit Sub
            If Not CheckVendorInsurance(WhereClause) Then Exit Sub
            
            'warn if items have no purchase price
            s = "SELECT COUNT(*) FROM EstimatedItems i WHERE ischangerequest=0 and " & IIf(mMode = fceBudget, "budgetdeleted=0 and BudgetRate", "podeleted=0 and PORate") & "=0 AND " & WhereClause
            On Error Resume Next
            i = 0
            i = HFApp.SqlExec(s)(0)
            On Error GoTo eh
            If i > 0 Then If MsgBox("Some of these items have no purchase price ($0.00)." & vbCrLf & vbCrLf & "Are you sure you want to generate " & IIf(mMode = fceBudget, "budgets", "PO's") & "?", vbOKCancel + vbExclamation, App.ProductName) = vbCancel Then Exit Sub
                   
            'okay do it
            If mMode = fceBudget Then
                Call GenerateBudgets
            Else
                Call GeneratePOs
            End If
           
            
            
        Case "RePrice"
            If Not SaveData(True) Then Exit Sub
            If mMode = fcePO Then
                If 0 = HFApp.SqlExec("SELECT COUNT(*) FROM EstimatedItems i WHERE ISNULL(PONumber,'')='' AND " & WhereClause)(0) Then
                    MsgBox "You have nothing selected that can be refreshed." & vbCrLf & _
                           "Use the check boxes to select some items that have" & vbCrLf & _
                           "not already been generated.", vbExclamation, App.ProductName
                Else
                    bRePrice = FEstimateItemsRefreshCosts.ShowForm(CostBasis, EffectiveDate, ApplyVarianceCat, VarianceCatDef, RefreshOverridden)
                    If bRePrice Then
                        
                        'convert costbasis to lookuptype
                        If CostBasis <> 0 Then CostBasis = CostBasis + 1
                        
                        
                        'if posting qty to accounting
                        If PostPOQtyToAccounting Then
                            'create correcting entries
                            s = ""
                            s = s & "--- CREATE REVERSING AND CORRECTING ENTRIES ------------------------" & vbCrLf
                            s = s & "insert estimateitems(" & vbCrLf
                            s = s & "  OriginalItemID, IsReversingItem, IsCorrectingItem, JCCategory  " & vbCrLf
                            s = s & ", PORate, POQty, POPretax, POTaxGroup, POJCTax, POJCTaxRate, PONJCTax, PONJCTaxRate, BudgetDeleted " & vbCrLf
                            s = s & ", EstAssemblyID, POIndex, Phase, Item, Job, JCExtra, JCCostCode, SortOrder, Description, Comments, TakeoffQty, TakeoffUOM, ConversionFactor, OrderUOM, POVendor, PONumber, ExcludeFromPO, OriginalJCCategory, Assembly, AssemblyDescription, Model, Location, Formula, Unit, BillingItemID, SalesQty, DivisionID, RequestDetailID" & vbCrLf
                            s = s & ", WBS01, WBS02, WBS03, WBS04, WBS05, WBS06, WBS07, WBS08, WBS09, WBS10, WBS11, WBS12, WBS13, WBS14, WBS15, WBS16, WBS17, WBS18, WBS19, WBS20, WBS21, WBS22, WBS23, WBS24, WBS25, WBS26, WBS27, WBS28, WBS29, WBS30, WBS31, WBS32, WBS33, WBS34, WBS35, WBS36, WBS37, WBS38, WBS39, WBS40)" & vbCrLf
                            s = s & vbCrLf '=================================================================
                            
                            s = s & "select" & vbCrLf
                            s = s & "  i.EstItemID OriginalItemID, 1 IsReversingItem, 0 IsCorrectingItem" & vbCrLf
                            s = s & ", " & DbQuote(Str, IIf(ApplyVarianceCat, VarianceCatDef, "i.JCCategory")) & " JCCategory" & vbCrLf
                            s = s & ", i.PORate" & vbCrLf
                            s = s & ", -1*i.POQty, -1*i.POPretax, i.POTaxGroup, -1*i.POJCTax ,i.POJCTaxRate, -1*i.PONJCTax, i.PONJCTaxRate, 1 BudgetDeleted" & vbCrLf
                            s = s & ", i.EstAssemblyID, i.POIndex, i.Phase, i.Item, i.Job, i.JCExtra, i.JCCostCode, i.SortOrder, i.Description, i.Comments, i.TakeoffQty, i.TakeoffUOM, i.ConversionFactor, i.OrderUOM, i.POVendor, i.PONumber, i.ExcludeFromPO, i.OriginalJCCategory, i.Assembly, i.AssemblyDescription, i.Model, i.Location, i.Formula, i.Unit, i.BillingItemID, i.SalesQty, i.DivisionID, i.RequestDetailID" & vbCrLf
                            s = s & ",i.WBS01, i.WBS02, i.WBS03, i.WBS04, i.WBS05, i.WBS06, i.WBS07, i.WBS08, i.WBS09, i.WBS10, i.WBS11, i.WBS12, i.WBS13, i.WBS14, i.WBS15, i.WBS16, i.WBS17, i.WBS18, i.WBS19, i.WBS20, i.WBS21, i.WBS22, i.WBS23, i.WBS24, i.WBS25, i.WBS26, i.WBS27, i.WBS28, i.WBS29 ,i.WBS30, i.WBS31, i.WBS32, i.WBS33, i.WBS34, i.WBS35, i.WBS36, i.WBS37, i.WBS38, i.WBS39, i.WBS40" & vbCrLf
                            s = s & "from EstimateItems i " & vbCrLf
                            s = s & "join EstimatedItems v on(i.EstItemID=v.EstItemID)" & vbCrLf
                            s = s & "where v.PODeleted=0" & vbCrLf
                            s = s & "and v.POGenBatch=0" & vbCrLf
                            s = s & "and i.RowType='Original'" & vbCrLf
                            s = s & "and i.porate<> case when ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0) <>0 then ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0)" & vbCrLf
                            s = s & "               when 'True'=" & DbQuote(Str, "" & HFApp.Options(ZeroRateOnRefreshCosts)) & " and i.POOverridden=0 then 0" & vbCrLf
                            s = s & "               else i.PORate end" & vbCrLf
                            If Not RefreshOverridden Then s = s & "   AND i.POOverridden=0" & vbCrLf
                            s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")
                            s = s & "union all" & vbCrLf
                            
                            
                            s = s & "select" & vbCrLf
                            s = s & "  i.EstItemID OriginalItemID, 0 IsReversingItem, 1 IsCorrectingItem" & vbCrLf
                            s = s & ", " & DbQuote(Str, IIf(ApplyVarianceCat, VarianceCatDef, "i.JCCategory")) & " JCCategory" & vbCrLf
                            s = s & ", case when ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0) <>0 then ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0)" & vbCrLf
                            s = s & "  when 'True'=" & DbQuote(Str, "" & HFApp.Options(ZeroRateOnRefreshCosts)) & " and i.POOverridden=0 then 0" & vbCrLf
                            s = s & "  else i.PORate end" & vbCrLf
                            s = s & ", i.POQty, i.POPretax, i.POTaxGroup, i.POJCTax , i.POJCTaxRate, i.PONJCTax, i.PONJCTaxRate, 1 BudgetDeleted" & vbCrLf
                            s = s & ", i.EstAssemblyID, i.POIndex, i.Phase, i.Item, i.Job, i.JCExtra, i.JCCostCode, i.SortOrder, i.Description, i.Comments, i.TakeoffQty, i.TakeoffUOM, i.ConversionFactor, i.OrderUOM, i.POVendor, i.PONumber, i.ExcludeFromPO, i.OriginalJCCategory, i.Assembly, i.AssemblyDescription, i.Model, i.Location, i.Formula, i.Unit, i.BillingItemID, i.SalesQty, i.DivisionID, i.RequestDetailID" & vbCrLf
                            s = s & ",i.WBS01, i.WBS02, i.WBS03, i.WBS04, i.WBS05, i.WBS06, i.WBS07, i.WBS08, i.WBS09, i.WBS10, i.WBS11, i.WBS12, i.WBS13, i.WBS14, i.WBS15, i.WBS16, i.WBS17, i.WBS18, i.WBS19, i.WBS20, i.WBS21, i.WBS22, i.WBS23, i.WBS24, i.WBS25, i.WBS26, i.WBS27, i.WBS28, i.WBS29 ,i.WBS30, i.WBS31, i.WBS32, i.WBS33, i.WBS34, i.WBS35, i.WBS36, i.WBS37, i.WBS38, i.WBS39, i.WBS40" & vbCrLf
                            s = s & "from EstimateItems i " & vbCrLf
                            s = s & "join EstimatedItems v on(i.EstItemID=v.EstItemID)" & vbCrLf
                            s = s & "where v.PODeleted=0" & vbCrLf
                            s = s & "and v.POGenBatch=0" & vbCrLf
                            s = s & "and i.RowType='Original'" & vbCrLf
                            s = s & "and i.porate<> case when ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0) <>0 then ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0)" & vbCrLf
                            s = s & "               when 'True'=" & DbQuote(Str, "" & HFApp.Options(ZeroRateOnRefreshCosts)) & " and i.POOverridden=0 then 0" & vbCrLf
                            s = s & "               else i.PORate end" & vbCrLf
                            If Not RefreshOverridden Then s = s & "   AND i.POOverridden=0" & vbCrLf
                            s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")
                            
                            
                            s = s & vbCrLf '=================================================================
                            s = s & "--- SET CORRECTINGITEMID ON ORIGINAL RECORD ------------------------" & vbCrLf
                            s = s & "update o set CorrectingItemID=c.EstItemID" & vbCrLf
                            s = s & "from estimateitems c" & vbCrLf
                            s = s & "join estimateitems o on c.OriginalItemID=o.EstItemID" & vbCrLf
                            s = s & "join EstimatedItems v on(c.EstItemID=v.EstItemID)" & vbCrLf
                            s = s & "where c.iscorrectingitem=1" & vbCrLf
                            s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")
                            s = s & vbCrLf '=================================================================
                            s = s & "--- SET REVERSINGITEMID ON ORIGINAL RECORD ------------------------" & vbCrLf
                            s = s & "update o set ReversingItemID=r.EstItemID" & vbCrLf
                            s = s & "from estimateitems r" & vbCrLf
                            s = s & "join estimateitems o on r.OriginalItemID=o.EstItemID" & vbCrLf
                            s = s & "join EstimatedItems v on(r.EstItemID=v.EstItemID)" & vbCrLf
                            s = s & "where r.isreversingitem=1" & vbCrLf
                            s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")
                            s = s & vbCrLf '=================================================================
                            s = s & "--- SET CORRECTINGITEMID ON REVERSING RECORD ------------------------" & vbCrLf
                            s = s & "update r set correctingitemid=o.correctingitemid" & vbCrLf
                            s = s & "from estimateitems o" & vbCrLf
                            s = s & "join estimateitems r on o.reversingitemid=r.estitemid" & vbCrLf
                            s = s & "join EstimatedItems v on(r.EstItemID=v.EstItemID)" & vbCrLf
                            s = s & "where 1=1" & vbCrLf
                            s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")

                            Call HFApp.SqlExec(s, dbHomefront, i)
                        
                        
                        End If
                    
                        'get item rates
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET POOverridden=0" & vbCrLf
                        If ApplyVarianceCat Then
                            s = s & "      ,VarianceJCCategory=" & DbQuote(Str, VarianceCatDef) & vbCrLf
                        End If
                        s = s & "      ,PORate = case when ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0) <>0 then ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0)" & vbCrLf
                        s = s & "                     when 'True'=" & DbQuote(Str, "" & HFApp.Options(ZeroRateOnRefreshCosts)) & " and i.POOverridden=0 then 0" & vbCrLf
                        s = s & "                     else i.PORate end" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.PODeleted=0 AND v.POGenBatch=0" & vbCrLf
                        s = s & "   AND i.RowType NOT IN('Original - Cancelled','Original - Reversal')" & vbCrLf
                        If Not RefreshOverridden Then s = s & "   AND i.POOverridden=0" & vbCrLf
                        s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")
                        Call HFApp.SqlExec(s, dbHomefront, i)
                        
                        're-extend prices and taxes
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET POPretax = i.PORate * i.POQty" & vbCrLf
                        s = s & "      ,POJCTax = i.PORate * i.POQty * i.POJCTaxRate/100" & vbCrLf
                        s = s & "      ,PONJCTax = i.PORate * i.POQty * i.PONJCTaxRate/100" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.PODeleted=0 AND v.POGenBatch=0" & vbCrLf
                        If Not RefreshOverridden Then s = s & "   AND i.POOverridden=0" & vbCrLf
                        s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")
                        Call HFApp.SqlExec(s, dbHomefront)
                                                
                        'keep budget values in synch
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET BudgetRate=v.PORate" & vbCrLf
                        s = s & "      ,BudgetPretax = i.PORate * i.BudgetQty" & vbCrLf
                        s = s & "      ,BudgetJCTax = i.PORate * i.BudgetQty * i.BudgetJCTaxRate/100" & vbCrLf
                        s = s & "      ,BudgetNJCTax = i.PORate * i.BudgetQty * i.BudgetNJCTaxRate/100" & vbCrLf
                        s = s & "      ,BudgetOverridden=0" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.PODeleted=0 AND v.POGenBatch=0" & vbCrLf
                        If Not RefreshOverridden Then s = s & "   AND i.POOverridden=0" & vbCrLf
                        s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID") & vbCrLf
                        s = s & "   AND v.BudgetsLocked=0" & vbCrLf
                        s = s & "   AND v.BudgetGenerated=0" & vbCrLf
                        Call HFApp.SqlExec(s, dbHomefront)
                        
                        Dirty = False
                        Call LoadItems
                        MsgBox "Costs have been successfully refreshed.", vbInformation, App.ProductName
                        
                    End If
                End If
            Else
                If 0 = HFApp.SqlExec("SELECT COUNT(*) FROM EstimatedItems i WHERE BudgetGenerated=0 AND " & WhereClause)(0) Then
                    MsgBox "You have nothing selected that can be refreshed." & vbCrLf & _
                           "Use the check boxes to select some items that have" & vbCrLf & _
                           "not already been generated.", vbExclamation, App.ProductName
                Else
                    bRePrice = FEstimateItemsRefreshCosts.ShowForm(CostBasis, EffectiveDate, ApplyVarianceCat, VarianceCatDef, RefreshOverridden)
                    If bRePrice Then
                        
                        'convert costbasis to lookuptype
                        If CostBasis <> 0 Then CostBasis = CostBasis + 1
                    
                        'get item rates
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET BudgetOverridden=0" & vbCrLf
                        If ApplyVarianceCat Then
                            s = s & "      ,VarianceJCCategory=" & DbQuote(Str, VarianceCatDef) & vbCrLf
                        End If
                        s = s & "      ,BudgetRate= case when ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.BudgetVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0) <>0 then ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.BudgetVendor," & DbQuote(Date, EffectiveDate) & "," & HFApp.DivisionID & "),0)" & vbCrLf
                        s = s & "                        when 'True'=" & DbQuote(Str, "" & HFApp.Options(ZeroRateOnRefreshCosts)) & " and i.BudgetOverridden=0 then 0" & vbCrLf
                        s = s & "                        else i.BudgetRate end" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.BudgetDeleted=0 AND v.BudgetGenerated=0" & vbCrLf
                        If Not RefreshOverridden Then s = s & "   AND i.BudgetOverridden=0" & vbCrLf
                        s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")
                        Call HFApp.SqlExec(s, dbHomefront, i)
                        
                        
                        're-extend prices and taxes
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET BudgetPretax = i.BudgetRate * i.BudgetQty" & vbCrLf
                        s = s & "      ,BudgetJCTax = i.BudgetRate * i.BudgetQty * i.BudgetJCTaxRate/100" & vbCrLf
                        s = s & "      ,BudgetNJCTax = i.BudgetRate * i.BudgetQty * i.BudgetNJCTaxRate/100" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.BudgetDeleted=0 AND v.BudgetGenerated=0" & vbCrLf
                        If Not RefreshOverridden Then s = s & "   AND i.BudgetOverridden=0" & vbCrLf
                        s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID")
                        Call HFApp.SqlExec(s, dbHomefront)
                                                
                        'keep po values in synch
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET PORate=v.BudgetRate" & vbCrLf
                        s = s & "      ,POPretax = i.BudgetRate * i.POQty" & vbCrLf
                        s = s & "      ,POJCTax = i.BudgetRate * i.POQty * i.POJCTaxRate/100" & vbCrLf
                        s = s & "      ,PONJCTax = i.BudgetRate * i.POQty * i.PONJCTaxRate/100" & vbCrLf
                        s = s & "      ,POOverridden=0" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.BudgetDeleted=0 AND v.BudgetGenerated=0" & vbCrLf
                        If Not RefreshOverridden Then s = s & "   AND i.BudgetOverridden=0" & vbCrLf
                        s = s & "   AND " & Replace(WhereClause("v"), "v.i.DivisionID", "v.DivisionID") & vbCrLf
                        s = s & "   AND v.POGenBatch=0" & vbCrLf
                        Call HFApp.SqlExec(s, dbHomefront)
                        
                        Dirty = False
                        Call LoadItems
                        MsgBox "Costs have been successfully refreshed." & vbCrLf, vbInformation, App.ProductName
                    End If
                
                End If
            End If
            
    End Select

Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonClick", s)
End Sub


Private Sub ReadFromPipeline()
On Error GoTo eh

    Dim X As New PipelineWrapper.PipelineWrapper
    Dim xml As String
    Dim s As String
    Dim i As Integer
    
    If MsgBox("Import from current Pipeline configuration?", vbQuestion + vbOKCancel, App.ProductName) = vbCancel Then Exit Sub
    If Not SaveData(True) Then Exit Sub
    
    Call X.Connect(HFApp.Options.ValueByName("CGVisionsRootURI"), HFApp.Options.ValueByName("CGVisionsUID"), HFApp.Options.ValueByName("CGVisionsPWD"))
    xml = X.GetJobProducts(mJob)
    
    'clean bad xml
    xml = Replace(xml, Chr(150), "-")
    
    
    'log response
    i = FreeFile()
    s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "Zybertech\" & VB.App.EXEName & "\pipeline.jobbom.xml")
    Call CreatePath("", FilePath(s))
    Open PathAppend(s) For Output As #i
    Print #i, xml
    Close #i
    
    
    
    s = "exec BIM_ImportJobBOM " & DbQuote(Str, HFApp.LoginID) & "," & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, xml)
    Call HFApp.SqlExec(s, dbHomefront)

    'reload
    Call mnuEstimateItemViewsSub_Click(CInt(mViewIndex + 1))

    MsgBox "Download Complete", vbInformation, App.ProductName

Exit Sub
eh: Call errHandler(SRCFILE & "ReadFromPipeline", s)
End Sub
Public Sub AddItem(Assembly As String, _
                   AssemblyDescription As String, _
                   Model As String, _
                   Phase As String, _
                   Item As String, Sequence As Long, _
                   Description As String, _
                   OrderQty As Double, _
                   OrderUOM As String, _
                   TakeoffQty As Double, _
                   TakeoffUOM As String, _
                   ConversionFactor As Double, RoundTo As Double, RoundDir As Long, WastePercent As Long, _
                   JCExtra As String, _
                   JCCostCode As String, _
                   JCCostCodeDesc As String, JCCategory As String, _
                   JCCategoryDesc As String, Vendor As String, _
                   VendorName As String, price As Double, _
                   TaxGroup As String, TaxGroupName As String, _
                   JCTaxRate As Double, NJCTaxRate As Double, _
                   POIndex As String, _
                   Comments As String, Formula As String, SalesQty As Double, Location As String, _
                   WBS, Optional Job As String)
On Error GoTo eh
    
    Dim i As Long
    Dim w As Long
    Dim s As String
    
    Dim rs As Recordset
    Dim POIndexDesc As String
    
    Dim sJobNumber As String

    
    
    mCancelTakeoff = False

    If Job = "" Then
        sJobNumber = "" & HFApp.SqlExec("select job from estimateassemblies where estassemblyid = " & DbQuote(Num, mCurrentEstAssemblyID))(0)
    Else
        sJobNumber = Job
    End If
    Dirty = True
    With gItems
        .AddItem ""
        i = .Rows - 1
        
        On Error Resume Next
        
        
        Set rs = HFApp.SqlExec("select description from tblpoindex where DivisionID = " & HFApp.DivisionID & " and poindex=" & DbQuote(Str, POIndex), dbHomefront)
        POIndexDesc = "" & rs("Description")
        If POIndexDesc = "" Then POIndexDesc = POIndex
        On Error GoTo eh
        
        .RowData(i) = "NEW"
        .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexNoCheckbox
        
        .TextMatrix(i, .ColIndex("Assembly")) = Assembly
        .TextMatrix(i, .ColIndex("AssemblyDescription")) = AssemblyDescription
        .TextMatrix(i, .ColIndex("EstAssemblyID")) = mCurrentEstAssemblyID
        .TextMatrix(i, .ColIndex("Job_No")) = sJobNumber
        .TextMatrix(i, .ColIndex("HFDescription")) = mCurrentAssemblyDesc
        If JCExtra <> "" Then
            .TextMatrix(i, .ColIndex("JCExtra")) = JCExtra
        Else
            'if you have an assembly loaded then default to whats on the screen, not whats in the database
            .TextMatrix(i, .ColIndex("JCExtra")) = IIf(mEstAssemblyID = 0, mCurrentJCExtra, txtJCExtra.Text)
        End If
        .TextMatrix(i, .ColIndex("Sequence")) = Sequence
        .TextMatrix(i, .ColIndex("EstPhase")) = Phase
        .TextMatrix(i, .ColIndex("EstItem")) = Item
        .TextMatrix(i, .ColIndex("ItemDesc")) = Description
        .TextMatrix(i, .ColIndex("ConversionFactor")) = ConversionFactor
        .TextMatrix(i, .ColIndex("RoundTo")) = RoundTo
        .TextMatrix(i, .ColIndex("RoundDir")) = RoundDir
        .TextMatrix(i, .ColIndex("WastePercent")) = WastePercent
        .TextMatrix(i, .ColIndex("JCCostCode")) = JCCostCode
        .TextMatrix(i, .ColIndex("JCCostCodeDesc")) = JCCostCodeDesc
        .TextMatrix(i, .ColIndex("JCCategory")) = JCCategory
        .TextMatrix(i, .ColIndex("OriginalJCCategory")) = JCCategory
        .TextMatrix(i, .ColIndex("JCCategoryDesc")) = JCCategoryDesc
        .TextMatrix(i, .ColIndex("POIndex")) = POIndex
        .TextMatrix(i, .ColIndex("POIndexDescription")) = POIndexDesc
        
        .TextMatrix(i, .ColIndex("SalesQty")) = SalesQty
        
        .TextMatrix(i, .ColIndex("ItemComments")) = Comments
        .TextMatrix(i, .ColIndex("TakeoffQty")) = TakeoffQty
        .TextMatrix(i, .ColIndex("TakeoffUOM")) = TakeoffUOM
        .TextMatrix(i, .ColIndex("OrderUOM")) = OrderUOM
        .TextMatrix(i, .ColIndex("Model")) = Model
        .TextMatrix(i, .ColIndex("OptionID")) = mCurrentOptionID
        .TextMatrix(i, .ColIndex("Formula")) = Formula
        
        .TextMatrix(i, .ColIndex("RoundTo")) = RoundTo
        .TextMatrix(i, .ColIndex("RoundDir")) = RoundDir
        
        
        .TextMatrix(i, .ColIndex("POQty")) = OrderQty
        .TextMatrix(i, .ColIndex("POVendor")) = Vendor
        .TextMatrix(i, .ColIndex("POVendorName")) = VendorName
        .TextMatrix(i, .ColIndex("PORate")) = price
        .TextMatrix(i, .ColIndex("POPretax")) = Round(OrderQty * price, 2)
        .TextMatrix(i, .ColIndex("POTaxGroup")) = TaxGroup
        .TextMatrix(i, .ColIndex("POJCTaxRate")) = JCTaxRate
        .TextMatrix(i, .ColIndex("PONJCTaxRate")) = NJCTaxRate
        .TextMatrix(i, .ColIndex("POJCTax")) = Round(.ValueMatrix(i, .ColIndex("POPretax")) * JCTaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("PONJCTax")) = Round(.ValueMatrix(i, .ColIndex("POPretax")) * NJCTaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("POTax")) = .ValueMatrix(i, .ColIndex("PONJCTax")) + .ValueMatrix(i, .ColIndex("POJCTax"))
        
        
        If mCurrentBudgetsLocked Then
            .TextMatrix(i, .ColIndex("BudgetDeleted")) = "True"
        Else
            .TextMatrix(i, .ColIndex("BudgetDeleted")) = "False"
            .TextMatrix(i, .ColIndex("BudgetQty")) = .TextMatrix(i, .ColIndex("POQty"))
            .TextMatrix(i, .ColIndex("BudgetVendor")) = .TextMatrix(i, .ColIndex("POVendor"))
            .TextMatrix(i, .ColIndex("BudgetVendorName")) = .TextMatrix(i, .ColIndex("POVendorName"))
            .TextMatrix(i, .ColIndex("BudgetRate")) = .TextMatrix(i, .ColIndex("PORate"))
            .TextMatrix(i, .ColIndex("BudgetPretax")) = .TextMatrix(i, .ColIndex("POPretax"))
            .TextMatrix(i, .ColIndex("BudgetTaxGroup")) = .TextMatrix(i, .ColIndex("POTaxGroup"))
            .TextMatrix(i, .ColIndex("BudgetJCTaxRate")) = .TextMatrix(i, .ColIndex("POJCTaxRate"))
            .TextMatrix(i, .ColIndex("BudgetNJCTaxRate")) = .TextMatrix(i, .ColIndex("PONJCTaxRate"))
            .TextMatrix(i, .ColIndex("BudgetJCTax")) = .TextMatrix(i, .ColIndex("POJCTax"))
            .TextMatrix(i, .ColIndex("BudgetNJCTax")) = .TextMatrix(i, .ColIndex("PONJCTax"))
            .TextMatrix(i, .ColIndex("BudgetTax")) = .TextMatrix(i, .ColIndex("POTax"))
        End If

        If Len(mTakeoffSettings) = 0 Then
            .TextMatrix(i, .ColIndex("Location")) = Location
            For w = 1 To 40
                .TextMatrix(i, .ColIndex("WBS" & format(w, "00"))) = WBS(w)
            Next
        Else
            .TextMatrix(i, .ColIndex("Location")) = Parse(mTakeoffSettings, 1, Chr(1))
            For w = 1 To 40
                .TextMatrix(i, .ColIndex("WBS" & format(w, "00"))) = Parse(mTakeoffSettings, w + 1, Chr(1))
            Next
        End If
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "AddItem")
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    
    Select Case ButtonMenu.Key
        
        Case "OpenJob"
            ProjectBased = False
            Call LoadViews
            Call LoadJob("")
        
        Case "OpenProject"
            ProjectBased = True
            ProjectPhaseBased = False
            Call LoadViews
            Call LoadJob("")
            
        Case "saveasquote"
            If mJob = "" Then
                MsgBox "You must save this quote first", vbInformation, App.ProductName
            Else
                Call SaveAs(True)
            End If
            
        Case "saveasjob"
            If mJob = "" Then
                MsgBox "You must save this quote first", vbInformation, App.ProductName
            Else
                Call SaveAs(False)
            End If
            
        Case "converttojob"
            If mJob = "" Then
                MsgBox "You must save this quote first", vbInformation, App.ProductName
            Else
                If (Not ValidateJobNumber(mJob)) Then
                    MsgBox "This quote cannot be converted to a job. The quote number" & vbCrLf & "does not conform to accounting's job numbering rules. Use" & vbCrLf & """Save as New Job"" instead.", vbInformation, App.ProductName
                    Exit Sub
                End If
                Call SaveData(False)
                s = "update tbljobs set isquote=0,quote_no=job_no where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mJob)
                Call HFApp.SqlExec(s)
                mMode = fceBudget
                Call Form_Load
                Call LoadJob(mJob)
            End If
            
            
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonClick", s)
End Sub


Private Sub ShowSelectedPOs(Preview As Boolean)
On Error GoTo eh
    Dim s As String
    Dim p As String
    Dim i As Long
    Dim POs As String
    Dim rs As Recordset
    Dim lastRpt As String
    Dim c As ZybUtil.Crystal
    Dim PrinterName As String
    
    
    
    POs = SelectedPOs
    If POs = "" Then Exit Sub
    
    If Not Preview Then
        'choose printer
        If Not VBPrintDlg(i, eprAll, True, , , True, , , , , , , Me.hwnd, , PrinterName) Then Exit Sub
        PrinterName = Trim(PrinterName)
    End If
        
        
    'check if any of the selected are not approved
    s = ""
    s = s & "SELECT " & vbCrLf
    s = s & " sum(case when status='Pending' then 1 else 0 end) Pending" & vbCrLf
    s = s & ",sum(case when status='Pending' then 0 else 1 end) Approved" & vbCrLf
    s = s & "  FROM POMaster" & vbCrLf
    s = s & " WHERE PONumber IN(" & POs & ")" & vbCrLf
    s = s & " and DivisionID = " & HFApp.DivisionID & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    If rs("Pending") > 0 Then
        If rs("Approved") = 0 Then
            Call MsgBox("These POs cannot be printed until approval is received.", vbInformation, App.ProductName)
            Exit Sub
        Else
            If vbCancel = MsgBox("" & rs("Pending") & " of these POs cannot be printed until approval is received." & vbCrLf & vbCrLf & "Do you want to print the " & rs("Approved") & " POs that are approved?", vbInformation + vbOKCancel, App.ProductName) Then Exit Sub
        End If
    End If
  
    s = ""
    s = s & "SELECT PONumber,POFormat" & vbCrLf
    s = s & "  FROM PurchaseOrders" & vbCrLf
    s = s & " WHERE PONumber IN(" & POs & ")" & vbCrLf
    s = s & " and Status<>'Pending'" & vbCrLf
    s = s & " and DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "ORDER BY POFormat" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    'loop thru this set building csv list of ponumbers in each format then showing rpt
    While Not rs.EOF
        If lastRpt <> "" & rs("POFormat") Then
            If POs <> "" And lastRpt <> "" Then
                s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt")
               
                Set c = New ZybUtil.Crystal
                Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
                On Error Resume Next
                Call c.ParameterValue("DivisionID", HFApp.DivisionID)
                
                POs = Mid(POs, 2)
                For i = 1 To Parse(POs, , "','")
                    p = Parse(POs, i, "','")
                    If Left(p, 1) = "'" Then p = Mid(p, 2)
                    If Right(p, 1) = "'" Then p = Left(p, Len(p) - 1)
                    If p <> "" Then
                        Call c.ParameterValue("PONumber", p)
                    End If
                Next
                
                On Error GoTo eh
                If Preview Then
                    Call c.PrintPreview("Print Preview")
                Else
                    Call c.PrintReport(PrinterName)
                End If
            
            End If
            lastRpt = "" & rs("POFormat")
            POs = ""
        End If
        POs = POs & "," & DbQuote(Str, "" & rs("PONumber"))
        rs.MoveNext
    Wend
    If POs <> "" And lastRpt <> "" Then
        s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt")
        Set c = New ZybUtil.Crystal
        Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
        On Error Resume Next
        Call c.ParameterValue("DivisionID", HFApp.DivisionID)
        
        POs = Mid(POs, 2)
        For i = 1 To Parse(POs, , "','")
            p = Parse(POs, i, "','")
            If Left(p, 1) = "'" Then p = Mid(p, 2)
            If Right(p, 1) = "'" Then p = Left(p, Len(p) - 1)
            If p <> "" Then
                Call c.ParameterValue("PONumber", p)
            End If
        Next
        
        On Error GoTo eh
        If Preview Then
            Call c.PrintPreview("Print Preview")
        Else
            Call c.PrintReport(PrinterName)
        End If
        
    End If
Exit Sub
eh:
    Select Case Err.Number
    Case 438 'File not found.
        MsgBox "Unable to open PO format """ & lastRpt & """." & vbCrLf & vbCrLf & "File not found:" & vbCrLf & PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt"), vbExclamation, App.ProductName
    Case Else: Call errHandler(SRCFILE & "ShowSelectedPOs", s)
    End Select
End Sub

Private Sub Slider_Move()
    Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim X As Long
    Dim c As Control
    
    Slider.Visible = True
    
    chkToolbarCaptions.Move Me.ScaleWidth - chkToolbarCaptions.Width, 30
    chkCancelledItems.Move Me.ScaleWidth - chkCancelledItems.Width, 30 + chkToolbarCaptions.Top + chkToolbarCaptions.Height
    
    '------ job frame -------------------
    frmJob.Move 0, Toolbar.Height, Me.ScaleWidth
    Slider2.Visible = True
    Slider2.Height = frmJob.Height - (5 * margin)
    Slider2.Top = gProperties.Top
    Slider2.Min = 5580
    Slider2.Max = Me.ScaleWidth - 240
    
    
    gProperties.Width = Slider2.Left - gProperties.Left - 2 * margin
    gProperties.Height = Slider2.Height
    Shape1(0).Width = gProperties.Width
    Shape1(0).Height = gProperties.Height
    
    txtNotes.Move Slider2.Left + Slider2.Width + Screen.TwipsPerPixelX, gProperties.Top, frmJob.Width - 2 * margin - Slider2.Left - Slider2.Width, gProperties.Height
    Shape1(1).Left = txtNotes.Left - Screen.TwipsPerPixelX
    Shape1(1).Width = txtNotes.Width
    Shape1(1).Height = txtNotes.Height
    lblNotes.Left = txtNotes.Left
    
    
    
    '------ purchasing frame -------------------
    frmPurchasing.Move 0, frmJob.Top + IIf(frmJob.Visible, frmJob.Height, 0), Me.ScaleWidth, Me.ScaleHeight - frmJob.Top - IIf(frmJob.Visible, frmJob.Height, 0)
    Slider.Min = 960
    Slider.Max = Me.ScaleWidth - 960
    Slider.Move Slider.Left, -15, Slider.Width, frmPurchasing.Height
    gAssemblies.Move -15, Slider.Top, Slider.Left + 15, Slider.Height
    imgShortcut(0).Move Me.ScaleLeft + 2500, 45 'gAssemblies.Width - imgShortcut(0).Width - 2315, 45
    imgShortcut(1).Move imgShortcut(0).Left - imgShortcut(0).Width - 30, imgShortcut(0).Top
    imgShortcut(2).Move imgShortcut(1).Left - imgShortcut(0).Width - 30, imgShortcut(0).Top
    imgShortcut(3).Move imgShortcut(2).Left - imgShortcut(0).Width - 30, imgShortcut(0).Top
    imgShortcut(8).Move imgShortcut(3).Left - imgShortcut(0).Width - 30, imgShortcut(0).Top
    imgShortcut(5).Move imgShortcut(8).Left - imgShortcut(0).Width - 30, imgShortcut(0).Top
    'Call gAssemblies.AutoSize(0, gAssemblies.Cols - 1)
    'Call gAssemblies.AutoSize(4, 6, True)
    'If gAssemblies.ColWidth(0) + 3 * gAssemblies.ColWidth(4) < gAssemblies.Width Then
    '    gAssemblies.ColWidth(0) = gAssemblies.Width - gAssemblies.ColWidth(4) - gAssemblies.ColWidth(5) - gAssemblies.ColWidth(6) - 255
    'End If
    frmEstItems.Move Slider.Left + Slider.Width, Slider.Top, Me.ScaleWidth - Slider.Left - Slider.Width, Slider.Height
        frmAssembly.Move 0, 0, frmEstItems.Width
            txtEstimatorNotes.Width = frmAssembly.Width - txtEstimatorNotes.Left - frmLockBudgets.Width
            frmLockBudgets.Move txtEstimatorNotes.Left + txtEstimatorNotes.Width, frmAssembly.Height - frmLockBudgets.Height
        frmChangeOrder.Move 0, 0, frmEstItems.Width
            txtCOComments.Width = frmChangeOrder.Width - txtCOComments.Left - 360
        frmPOIndex.Move 0, 0, frmEstItems.Width
        X = IIf(frmAssembly.Visible, frmAssembly.Height, 0) + IIf(frmPOIndex.Visible, frmPOIndex.Height, 0) + IIf(frmChangeOrder.Visible, frmChangeOrder.Height, 0)
        gItems.Move 0, X, frmEstItems.Width, frmEstItems.Height - X
    frmRFPs.Move frmEstItems.Left, frmEstItems.Top, frmEstItems.Width, frmEstItems.Height
        Slider1.Min = gBids.Top + 540
        Slider1.Max = frmRFPs.Height - 540
        Slider1.Move 0, Slider1.Top, frmRFPs.Width
        gBids.Move 0, gBids.Top, Slider1.Width, Slider1.Top - gBids.Top
        gBidItems.Move 0, Slider1.Top + Slider1.Height, Slider1.Width, frmRFPs.Height - Slider1.Top - Slider1.Height
    
    
    '------ billing frame -------------------
    frmBilling.Move frmPurchasing.Left, frmPurchasing.Top, frmPurchasing.Width, frmPurchasing.Height
    frmBillingTasks.Move frmBilling.Width - frmBillingTasks.Width, 0
    
    Slider3.Visible = True
    Slider3.Move margin, Slider3.Top, frmBillingTasks.Left - margin
    'Slider3.Min = 5580
    'Slider3.Max = frmBilling.Width - 240
    gContract.Move margin, gContract.Top, frmBillingTasks.Left - margin, Slider3.Top - gContract.Top
    gInvoices.Move margin, Slider3.Top + Slider3.Height + gContract.Top, gContract.Width, frmBilling.Height - (Slider3.Top + Slider3.Height + gContract.Top + margin)
    lblInvoices.Top = gInvoices.Top - lblInvoices.Height - margin
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim u As String

    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    
    'unlock record if is locked by me. this looks wrong but it isnt.
    'RecordLockedBy() only returns other user names. If is locked by me then it returns nothing.
    u = HFApp.RecordLockedBy("BudgetsAndPOs", mJob)
    If u = "" Then Call HFApp.UnLockRecord("BudgetsAndPOs", mJob)
    
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gBids)
    Call IniPutGrid(Me, gAssemblies)
    Call IniPutGrid(Me, gItems, , mMode & mViewIndex)
    
    
    
    Call IniPutGrid(Me, gBidItems)
    Call IniPutGrid(Me, gContract)
    Call IniPutGrid(Me, gInvoices)
    
    Call IniPut(AppIni, "Options", "EstimateItemsView" & mMode, mViewIndex)
    Call IniPut(AppIni, "Options", "EstimateItemsView.JobInfo." & mMode, frmJob.Visible)
    
    Call IniPut(AppIni, "Formatting", "AcceptedForeColor", AcceptedForeColor)
    Call IniPut(AppIni, "Formatting", "AcceptedBackColor", AcceptedBackColor)
    Call IniPut(AppIni, "Formatting", "AcceptedStyle", AcceptedStyle)
    Call IniPut(AppIni, "Formatting", "DeclinedForeColor", DeclinedForeColor)
    Call IniPut(AppIni, "Formatting", "DeclinedBackColor", DeclinedBackColor)
    Call IniPut(AppIni, "Formatting", "DeclinedStyle", DeclinedStyle)
    Call IniPut(AppIni, "Formatting", "MinForeColor", MinForeColor)
    Call IniPut(AppIni, "Formatting", "MinBackColor", MinBackColor)
    Call IniPut(AppIni, "Formatting", "MinStyle", MinStyle)
    Call IniPut(AppIni, "Formatting", "MaxForeColor", MaxForeColor)
    Call IniPut(AppIni, "Formatting", "MaxBackColor", MaxBackColor)
    Call IniPut(AppIni, "Formatting", "MaxStyle", MaxStyle)
        
    HFApp.Options.ValueByName("CancelledItems") = chkCancelledItems.value
    
    Unload FComments
End Sub


Private Sub ResetChildrenChecked(ByVal ParentRow As Long)
    Dim r As Long
    If ParentRow < 0 Then Exit Sub
    With gAssemblies
        'for all my children set checkbox to same as me
        r = .GetNodeRow(ParentRow, flexNTFirstChild)
        While r <> -1
            .Cell(flexcpText, r, 3) = .Cell(flexcpText, ParentRow, 3)
            Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(.Cell(flexcpValue, r, 2), .Cell(flexcpValue, r, 3), .Cell(flexcpValue, r, .ColIndex("IsChangeRequest")), .Cell(flexcpText, r, .ColIndex("POIconType")))
            Call ResetChildrenChecked(r)
            r = .GetNodeRow(r, flexNTNextSibling)
        Wend
    End With
End Sub

Private Sub ResetParentChecked(ByVal ChildRow As Long)
    Dim ParentRow As Long
    Dim r As Long
    With gAssemblies
        If ChildRow = -1 Then Exit Sub
        
        If .Cell(flexcpValue, ChildRow, .ColIndex("IsChangeRequest")) = 1 Then Exit Sub
        
        ParentRow = .GetNodeRow(ChildRow, flexNTParent)
        If ParentRow = -1 Then Exit Sub
        
        .Cell(flexcpText, ParentRow, 3) = .Cell(flexcpText, ChildRow, 3)
        Set .Cell(flexcpPicture, ParentRow, 0) = MultiStateIcon(.Cell(flexcpValue, ParentRow, 2), .Cell(flexcpValue, ParentRow, 3), .Cell(flexcpValue, ParentRow, .ColIndex("IsChangeRequest")), .Cell(flexcpText, ParentRow, .ColIndex("POIconType")))
        
        r = .GetNodeRow(ChildRow, flexNTFirstSibling)
        While r <> -1
            If .Cell(flexcpText, r, 3) <> .Cell(flexcpText, ChildRow, 3) Then
                .Cell(flexcpText, ParentRow, 3) = msSome
                Set .Cell(flexcpPicture, ParentRow, 0) = MultiStateIcon(.Cell(flexcpValue, ParentRow, 2), .Cell(flexcpValue, ParentRow, 3), .Cell(flexcpValue, ParentRow, .ColIndex("IsChangeRequest")), .Cell(flexcpText, ParentRow, .ColIndex("POIconType")))
                r = -1
            Else
                r = .GetNodeRow(r, flexNTNextSibling)
            End If
        Wend
        Call ResetParentChecked(ParentRow)
    End With
End Sub

Private Sub ClearWhereClause()
    Dim r As Long
    
    
    With gAssemblies
        If .Rows > 1 Then
            .Cell(flexcpText, 1, 3, .Rows - 1, 3) = msNone
            For r = 1 To .Rows - 1
                Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(.Cell(flexcpValue, r, 2), msNone, .Cell(flexcpValue, r, .ColIndex("IsChangeRequest")), .Cell(flexcpText, r, .ColIndex("POIconType")))
            Next
        End If
    End With
    
    If gItems.Rows > 1 Then
        gItems.Cell(flexcpChecked, 1, gItems.ColIndex("Selected"), gItems.Rows - 1, gItems.ColIndex("Selected")) = flexUnchecked
    End If
End Sub

Private Function WhereClause(Optional prefix As String)
    Dim i As Long
    Dim s As String
    Dim i2 As Long
    If prefix <> "" Then prefix = prefix & "."
    
    'add items to where clause
    With gItems
        For i = 1 To .Rows - 1
            If Not .IsSubtotal(i) And .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked Then
                s = s & "," & .TextMatrix(i, .ColIndex("EstItemID"))
            End If
        Next
        If s <> "" Then s = " OR " & prefix & "EstItemID IN(" & Mid(s, 2) & ")"
    End With
    
    
    'add groups to where clause intelligently.
    '  if a node is checked then skip over all its children.
    '  They WILL be checked and we don't need to add them to the where clause
    With gAssemblies
        i = 1
        While i < .Rows And i >= 0
            If .Cell(flexcpText, i, 3) = msAll And .RowData(i) <> "" Then
                s = s & " OR (" & Mid(Replace(.RowData(i), " AND ", " AND " & prefix), 6) & ")" & vbCrLf
                i2 = .GetNodeRow(i, flexNTNextSibling)
                If i2 = -1 Then
                    i = i + 1
                Else
                    i = i2
                End If
            Else
                i = i + 1
            End If
        Wend
    End With

    If s = "" Then
        WhereClause = "1=2"
    Else
        WhereClause = "(" & prefix & "EstItemID IS NOT NULL AND " & Mid(s, 5) & ")"
    End If
End Function



Public Function SelectedPOs() As String
'return a csv list of pos that are selected
'if anything other than a po is selected then return nothing.

    Dim i As Long
    Dim s As String
    Dim PONumber As String
    
    'check for selected assemblies
    With gAssemblies
        i = 0
        s = ""
        For i = 1 To .Rows - 1
            If (.Row = i Or .Cell(flexcpText, i, 3) <> msNone) And (.Cell(flexcpData, i, 1) = "PONumber" Or .Cell(flexcpData, i, 1) = "PONumberVendor") Then
                PONumber = Parse(.Cell(flexcpText, i, 1), 1, Chr(2))
                If PONumber <> "" Then s = s & "," & DbQuote(Str, PONumber)
            End If
        Next
        s = Mid(s, 2)
    End With
    
    SelectedPOs = s
End Function


Private Sub GeneratePOs()
On Error GoTo eh
    
    Dim c As Long
    Dim i As Long
    
    Dim s  As String
    Dim rs As Recordset
    Dim POPrefix As String
    Dim PONumber As String
    Dim Batch As Long
    Dim GlobalPO As Boolean
    Dim SQLBatch As String
        
    Dim ItemWhereClause As String
    Dim ItemIDString As String
    Dim CombineJobs As Boolean      'Combine multiple Jobs per vendor to a single purchase order.
    Dim CombinePOIndexes As Boolean 'Combine multiple po indexes per vendor to a single purchase order.
    Dim POPerUnit As Boolean        'Seperate pos for each unit
    
    Dim CCheckPassed As Boolean
    Dim PostExtra As Boolean
    
    PostExtra = HFApp.Options.ValueByName("PostJCExtraAsSubJob") = "True"
    
    'this is not valid anymore. Cant support it and the poindex based ponumbers. Or potentially BuildPro???
    'CombinePOIndexes = HFApp.Options(GroupPurchaseOrdersByVendor)
    CombinePOIndexes = False
    
    SQLBatch = ""
   
    s = "SELECT COUNT(*) FROM EstimatedItems i WHERE PODeleted=0 AND POGenBatch=0 AND " & WhereClause
    If 0 = HFApp.SqlExec(s)(0) Then
        MsgBox "Nothing to generate", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    
    If Not FGenerate.VerifyBudgets(WhereClause) Then Exit Sub
    
    
    Call FProgress.Progress("Generating PO's", "initializing batch...", 1, 2)

    'create new batch
    Call HFApp.SqlExec("INSERT INTO Batches(BatchType,UStmp,TStmp,DivisionID) VALUES('PO Generation'," & DbQuote(Str, HFApp.LoginID) & ",GETDATE()," & HFApp.DivisionID & ")")
    Batch = HFApp.SqlIdentity("Batches")
    SQLBatch = SQLBatch & "UPDATE EstimateItems SET POGenBatch=" & DbQuote(Num, Batch) & " WHERE EstItemID IN(SELECT EstItemID FROM EstimatedItems i WHERE PODeleted=0 AND POGenBatch=0 AND (" & WhereClause & "))" & vbCrLf
    
    ItemWhereClause = " WHERE ei.EstItemID IN(SELECT EstItemID FROM EstimatedItems i WHERE PODeleted=0 AND POGenBatch=0 AND (" & WhereClause & "))"
    ItemIDString = " IN(SELECT EstItemID FROM EstimatedItems i WHERE " & WhereClause & ")"
    Call ClearWhereClause
    
    
    'Get POs
    POPerUnit = False
    CombineJobs = False
    If ProjectBased Then
        If MsgBox("Consolidate Multiple Job's to a single PO per Vendor?", vbYesNo + vbQuestion, "Consolidate PO's") = vbYes Then
            CombineJobs = True
            If CombinePOIndexes Then
                c = HFApp.SqlExec("SELECT count(DISTINCT ISNULL(Community,'')+ISNULL(POVendor,'')+ISNULL(POFormat,'')) FROM EstimatedItems WHERE EstItemID " & ItemIDString, dbHomefront)(0)
                s = "SELECT DISTINCT ei.Community Job,ei.POVendor,ei.POFormat "
            Else
                c = HFApp.SqlExec("SELECT count(DISTINCT ISNULL(Community,'')+ISNULL(POIndex,'')+ISNULL(POVendor,'')) FROM EstimatedItems WHERE EstItemID " & ItemIDString, dbHomefront)(0)
                s = "SELECT DISTINCT ei.Community Job,ei.POIndex,ei.POVendor "
            End If
        Else
            c = HFApp.SqlExec("SELECT count(DISTINCT ISNULL(Job_No,'')+ISNULL(POVendor,'')+ISNULL(POFormat,'')+ISNULL(Unit,'')) FROM EstimatedItems WHERE EstItemID " & ItemIDString, dbHomefront)(0)
            If c > 1 Then
                If MsgBox("POs by unit?", vbYesNo + vbQuestion, "Consolidate PO's") = vbYes Then
                    POPerUnit = True
                    c = HFApp.SqlExec("SELECT count(DISTINCT ISNULL(Job_No,'')+ISNULL(POVendor,'')+ISNULL(POFormat,'')+ISNULL(Unit,'')) FROM EstimatedItems WHERE EstItemID " & ItemIDString, dbHomefront)(0)
                    s = "SELECT DISTINCT ei.Job_No Job,ei.POVendor,ei.POFormat+ei.Unit,ei.Unit "
                ElseIf CombinePOIndexes Then
                    c = HFApp.SqlExec("SELECT count(DISTINCT ISNULL(Job_No,'')+ISNULL(POVendor,'')+ISNULL(POFormat,'')) FROM EstimatedItems WHERE EstItemID " & ItemIDString, dbHomefront)(0)
                    s = "SELECT DISTINCT ei.Job_No Job,ei.POVendor,ei.POFormat "
                Else
                    c = HFApp.SqlExec("SELECT count(DISTINCT ISNULL(Job,'')+ISNULL(POIndex,'')+ISNULL(POVendor,'')) FROM EstimateItems WHERE EstItemID " & ItemIDString, dbHomefront)(0)
                    s = "SELECT DISTINCT ei.Job_NO Job,ei.POIndex,ei.POVendor "
                End If
            Else
                c = HFApp.SqlExec("SELECT count(DISTINCT ISNULL(Job,'')+ISNULL(POIndex,'')+ISNULL(POVendor,'')) FROM EstimateItems WHERE EstItemID " & ItemIDString, dbHomefront)(0)
                s = "SELECT DISTINCT ei.Job_NO Job,ei.POIndex,ei.POVendor "
            End If
        End If
    Else
        If HFApp.Options(MultiFamily) Then
            If HFApp.Options(AccountingSystem) = asIntacct And PostExtra Then
                'not possible if posting extras as subjobs
                POPerUnit = False
            Else
                POPerUnit = MsgBox("Individual POs per unit?", vbYesNo + vbQuestion, "Consolidate PO's") = vbYes
            End If
        End If
        
        'get count
        s = ""
        s = s & "SELECT count(DISTINCT" & vbCrLf
        s = s & "   ISNULL(Job_No,'')" & vbCrLf
        s = s & "  +ISNULL(POVendor,'')" & vbCrLf
        s = s & "  +ISNULL(" & IIf(CombinePOIndexes, "POFormat", "POIndex") & ",'')" & vbCrLf
        If POPerUnit Then
        s = s & "  +ISNULL(unit,'')" & vbCrLf
        End If
        s = s & "  )" & vbCrLf
        s = s & "FROM EstimatedItems" & vbCrLf
        s = s & "WHERE EstItemID " & ItemIDString
        c = HFApp.SqlExec(s, dbHomefront)(0)
        
        'get records
        s = "SELECT DISTINCT ei.Job_No Job,ei.POVendor," & IIf(CombinePOIndexes, "ei.POFormat", "ei.POIndex") & IIf(POPerUnit, ",ei.unit", "") & vbCrLf
        
    End If
    s = s & ", CASE WHEN isnull(jpo.RetainagePercent,0)>isnull(v.HoldbackPercentage,0) THEN isnull(jpo.RetainagePercent,0) ELSE isnull(v.HoldbackPercentage,0) END RetainagePercent" & vbCrLf
    s = s & "      ,PurchDelMethod DeliveryMethod" & vbCrLf
    s = s & "      ,PurchContact DeliveryRecipient,pe.Email" & vbCrLf
    s = s & "      ,CASE when isnull(pe.Email,'')  = ''  then case when PurchDelMethod = 2 THEN PurchFax ELSE '' END else pe.email end DeliveryAddress" & vbCrLf
    s = s & "      ,jpo.POIndex" & vbCrLf
    If Not CombineJobs Then
        s = s & "      ,jpo.Job_No Job" & vbCrLf
    End If
    s = s & "      ,ei.POVendor Vendor" & vbCrLf
    s = s & "      ,jpo.Description" & vbCrLf
    s = s & "      ,jpo.StandardText" & vbCrLf
    s = s & "      ,jpo.HideQty" & vbCrLf
    s = s & "      ,jpo.HidePrice" & vbCrLf
    s = s & "      ,jpo.TotalOnly" & vbCrLf
    s = s & "      ,jpo.ShipVia" & vbCrLf
    s = s & "      ,jpo.FOB" & vbCrLf
    s = s & "      ,jpo.Terms" & vbCrLf
    s = s & "      ,jpo.OrderedBy" & vbCrLf
    s = s & "      ,x.mpo" & vbCrLf
    s = s & "  FROM EstimatedItems ei " & vbCrLf
    s = s & "       left outer JOIN JobPOIndex jpo ON(jpo.DivisionID = ei.DivisionID and ei.Job_No=jpo.Job_No AND ei.POIndex=jpo.POIndex)" & vbCrLf
    s = s & "       left outer JOIN tblPOIndex x ON(x.DivisionID = ei.DivisionID and ei.POIndex=x.POIndex)" & vbCrLf
    s = s & "       left outer JOIN tblVendors v ON(v.DivisionID = ei.DivisionID and ei.POVendor=v.Vendor_id)" & vbCrLf
    s = s & "       LEFT OUTER JOIN dbo.ContactPurchasingEmails pe on (ei.POvendor = pe.VendorCode and pe.DivisionID = ei.DivisionID)" & vbCrLf
    s = s & ItemWhereClause & vbCrLf
    s = s & "ORDER BY 1,2,3"
        
    i = 0
    Call FProgress.Progress(, "Writing purchase orders...", 0, 1)
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        i = i + 1
        Call FProgress.Progress(, , i, c)


        s = " dbo.Purch_CreatePONumber " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, "" & rs("Job")) & "," & DbQuote(Str, "" & rs("poindex"))
        PONumber = "" & HFApp.SqlExec(s)(0)
        If PONumber = "" Then Err.Raise vbObjectError + 33, "", "dbo.Purch_CreatePONumber returned a blank PONumber"


        
        ''this is a new mess
        s = ""
        s = s & "UPDATE EstimateItems" & vbCrLf
        s = s & "   SET PONumber = " & DbQuote(Str, PONumber) & vbCrLf
        s = s & " WHERE DivisionID =" & HFApp.DivisionID & " and EstItemID " & vbCrLf
        s = s & " IN(SELECT EstItemID" & vbCrLf
        s = s & "                      FROM EstimatedItems" & vbCrLf
        s = s & "                     WHERE POGenBatch = " & DbQuote(Str, Batch) & vbCrLf
        If Not CombineJobs Then
            s = s & "                       AND Job_No=" & DbQuote(Str, "" & rs("Job")) & vbCrLf
        End If
        If POPerUnit Then
            s = s & "                       AND unit=" & DbQuote(Str, "" & rs("unit")) & vbCrLf
        End If
        If CombinePOIndexes Then
            s = s & "                       AND POFormat=" & DbQuote(Str, "" & rs("POFormat")) & vbCrLf
        Else
            s = s & "                       AND POIndex=" & DbQuote(Str, "" & rs("POIndex")) & vbCrLf
        End If
        s = s & "                       AND POVendor=" & DbQuote(Str, "" & rs("POVendor")) & ")" & vbCrLf
        
        SQLBatch = SQLBatch & s & vbCrLf
    
    
        s = ""
        s = s & "INSERT INTO POMaster(pogenbatch,ismpo,IncludeDocuments,DivisionID,RetainagePercent,DeliveryMethod,DeliveryRecipient,DeliveryAddress,PONumber,POIndex,Job,Vendor,PODate,Description,StandardText,HideQty,HidePrice,TotalOnly,ShipVia,FOB,Terms,OrderedBy,Summarized,UStmp,TStmp)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Num, Batch) & vbCrLf
        s = s & "      ," & DbQuote(Bit, "" & rs("mpo")) & vbCrLf
        s = s & "      ,1" & vbCrLf
        s = s & "      ," & HFApp.DivisionID & vbCrLf
        s = s & "      ," & DbQuote(Num, Val("" & rs("RetainagePercent"))) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("DeliveryMethod")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("DeliveryRecipient")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("DeliveryAddress")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & PONumber) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("POIndex")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Job")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Vendor")) & vbCrLf
        s = s & "      ,GETDATE()" & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Description")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("StandardText")) & vbCrLf
        s = s & "      ," & DbQuote(Bit, "" & rs("HideQty")) & vbCrLf
        s = s & "      ," & DbQuote(Bit, "" & rs("HidePrice")) & vbCrLf
        s = s & "      ," & DbQuote(Bit, "" & rs("TotalOnly")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipVia")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("FOB")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Terms")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("OrderedBy")) & vbCrLf
        s = s & "      ," & DbQuote(Bit, HFApp.Options(PostSummarizedPOs)) & vbCrLf
        s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "      ,GETDATE()" & ")" & vbCrLf
        SQLBatch = SQLBatch & s & vbCrLf
        
        
        
        rs.MoveNext
    Wend
    
Call FProgress.Progress(, "updating records", 0, 1)
    i = 0
    PONumber = ""
    
'MsgBox 1
    Call FProgress.Progress(, "closing batch", 0, 1)
    
'MsgBox 2
    
    'create POItems
    s = ""
    s = s & "INSERT INTO POItems(PONumber,DivisionID,Description,Comments,OrderQty,OrderUOM,TakeoffQty,TakeoffUOM,Rate,Pretax,TaxGroup,JCTax,JCTaxRate,NJCTax,NJCTaxRate,DontPrint,VarianceQty,variancejccategory,variancepretax,variancejctax,variancenjctax,Job,JCExtra,JCCostCode,JCCategory,EstPhase,EstItem,EstItemID,CustomerChangeOrder,GenBatch,SortOrder,PartNumber)" & vbCrLf
    s = s & "SELECT PONumber" & vbCrLf
    s = s & "      ," & HFApp.DivisionID & vbCrLf
    s = s & "      ,ItemDesc Description" & vbCrLf
    s = s & "      ,ItemComments Comments" & vbCrLf
    s = s & "      ,POQty OrderQty" & vbCrLf
    s = s & "      ,OrderUOM" & vbCrLf
    s = s & "      ,TakeoffQty" & vbCrLf
    s = s & "      ,TakeoffUOM" & vbCrLf
    s = s & "      ,PORate Rate" & vbCrLf
    s = s & "      ,POPretax Pretax" & vbCrLf
    s = s & "      ,POTaxGroup TaxGroup" & vbCrLf
    s = s & "      ,POJCTax JCTax" & vbCrLf
    s = s & "      ,POJCTaxRate JCTaxRate" & vbCrLf
    s = s & "      ,PONJCTax NJCTax" & vbCrLf
    s = s & "      ,PONJCTaxRate NJCTaxRate" & vbCrLf
    s = s & "      ,0 DontPrint" & vbCrLf
    s = s & "      ,varianceqty,variancejccategory,round(variancepretax,2),Round(VarianceJCTax,2),Round(VarianceNJCTax,2)" & vbCrLf
    s = s & "      ,Job_No Job" & vbCrLf
    s = s & "      ,JCExtra" & vbCrLf
    s = s & "      ,JCCostCode" & vbCrLf
    s = s & "      ,JCCategory" & vbCrLf
    s = s & "      ,EstPhase" & vbCrLf
    s = s & "      ,EstItem" & vbCrLf
    s = s & "      ,EstItemID" & vbCrLf
    s = s & "      ,ChangeOrder CustomerChangeOrder" & vbCrLf
    s = s & "      ,POGenBatch" & vbCrLf
    s = s & "      ,SortOrder" & vbCrLf
    s = s & "      ,PartNumber" & vbCrLf
    s = s & "  FROM EstimatedItems" & vbCrLf
    s = s & " WHERE POGenBatch=" & DbQuote(Num, Batch)
    s = s & " order by SortOrder" & vbCrLf
    SQLBatch = SQLBatch & s & vbCrLf
     

'MsgBox 3
    'set line item numbers and description
    If HFApp.Options(PostSummarizedPOs) Then
        s = ""
        s = s & "update poitems" & vbCrLf
        s = s & "set linedescription=isnull(c.description,a.Description)" & vbCrLf
        s = s & "   ,linenumber=(select LineNumber" & vbCrLf
        s = s & "                from (select row_number() over (PARTITION BY ponumber order by ponumber,job,jcextra,jccostcode,jccategory,isnull(taxgroup,'')) LineNumber,ponumber,job,jcextra,jccostcode,jccategory,isnull(taxgroup,'') TaxGroup" & vbCrLf
        s = s & "                      from poitems x where x.ponumber=a.ponumber" & vbCrLf
        s = s & "                      group by ponumber,job,jcextra,jccostcode,jccategory,isnull(taxgroup,'')) b " & vbCrLf
        s = s & "                where a.ponumber=b.ponumber" & vbCrLf
        s = s & "                and isnull(a.taxgroup,'')=isnull(b.taxgroup,'')" & vbCrLf
        s = s & "                and a.job=b.job" & vbCrLf
        s = s & "                and isnull(a.jcextra,'')=isnull(b.jcextra,'')" & vbCrLf
        s = s & "                and a.jccostcode=b.jccostcode" & vbCrLf
        s = s & "                and a.jccategory=b.jccategory)" & vbCrLf
        s = s & "from poitems a" & vbCrLf
        s = s & "left outer join standardcostcodes c on a.DivisionID = c.DivisionID and a.jccostcode=c.costcode" & vbCrLf
        s = s & "where a.DivisionID =" & HFApp.DivisionID & " and a.genBatch=" & DbQuote(Num, Batch) & vbCrLf
        SQLBatch = SQLBatch & s & vbCrLf
'MsgBox 4
    Else
    
        s = ""
        s = s & "update poitems" & vbCrLf
        s = s & "set linedescription=a.description" & vbCrLf
        s = s & "   ,linenumber=(select LineNumber" & vbCrLf
        s = s & "                from (select row_number() over (PARTITION BY DivisionID,ponumber order by ponumber,ItemSeq) LineNumber,ponumber,ItemSeq" & vbCrLf
        s = s & "                      from poitems x where x.ponumber=a.ponumber and x.DivisionID=a.DivisionID" & vbCrLf
        s = s & "                      group by DivisionID,ponumber,ItemSeq) b " & vbCrLf
        s = s & "                where a.ponumber=b.ponumber" & vbCrLf
        s = s & "                and a.itemseq=b.itemseq)" & vbCrLf
        s = s & "from poitems a" & vbCrLf
        s = s & "where a.DivisionID = " & HFApp.DivisionID & " and a.genBatch=" & DbQuote(Num, Batch) & vbCrLf
        SQLBatch = SQLBatch & s & vbCrLf
    End If
    
    
    Call HFApp.SqlExec(SQLBatch)
    
    
 
    
    
    
    'check consistancy
    Dim a As Long
    Dim b As Long
    a = HFApp.SqlExec("select count(*) from estimateitems where pogenbatch=" & DbQuote(Num, Batch))(0)
    b = HFApp.SqlExec("select count(*) from poitems where genbatch=" & DbQuote(Num, Batch))(0)
    CCheckPassed = (a = b And a <> 0)
    If Not CCheckPassed Then 'failed. try again
        Call HFApp.SqlExec(SQLBatch)
        
        're-check consistancy
        a = HFApp.SqlExec("select count(*) from estimateitems where pogenbatch=" & DbQuote(Num, Batch))(0)
        b = HFApp.SqlExec("select count(*) from poitems where genbatch=" & DbQuote(Num, Batch))(0)
        CCheckPassed = (a = b And a <> 0)
        
    End If
        
    If CCheckPassed Then
        
        s = ""
        s = s & "update p set isvar = x.isvar" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "join (select i.divisionid,i.ponumber,case when sum(isnull(cast(c.isvariance as int),0) + isnull(cast(v.isvariance as int),0))>0 then 1 else 0 end isVar" & vbCrLf
        s = s & "      from poitems i " & vbCrLf
        s = s & "      left join standardcategories c on i.divisionid=c.divisionid and i.jccategory=c.category" & vbCrLf
        s = s & "      left join standardcategories v on i.divisionid=v.divisionid and i.variancejccategory=v.category and i.variancepretax<>0" & vbCrLf
        s = s & "      group by i.divisionid,i.ponumber" & vbCrLf
        s = s & "      ) x on p.divisionid=x.divisionid and p.ponumber=x.ponumber" & vbCrLf
        s = s & "where p.pogenbatch=" & DbQuote(Num, Batch)
        Call HFApp.SqlExec(s)
    
    Else
        
        Call WriteException(SRCFILE & "GeneratePOs", "Consistancy check FAILED", SQLBatch)
        
        MsgBox "The batch (" & Batch & ") failed consistency checks and will be rolled back. Please retry the operation. If " & _
               "generation fails twice an error condition may be present in your data. Please contact support.", vbInformation, App.ProductName
                
        s = "exec dbo.Purch_RollbackPOBatch " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Num, Batch)
        Call HFApp.SqlExec(s)
        
        Unload FProgress
        Exit Sub
    
    End If
    
   
    
    
    
'MsgBox 9
    Call SetPOStatuses(Batch)
'MsgBox 10
    Call RefreshTree
'MsgBox 11
    
    Unload FProgress

'MsgBox 12

    s = ""
    If HFApp.Options(AccountingSystem) = asSimply Then
        Call HFApp.WriteJobToAccounting(mJob)
    End If
    
    If HFApp.Options.ValueByName("BuildProSendPOsImmediately") = "true" Then
        Call SendBuildProJobs(mJob)
        Call SendBuildProPOs(mJob, "")
    End If
    
'MsgBox 13

Exit Sub
eh:
    Call errHandler(SRCFILE & "GeneratePOs", s)
    Call WriteException(SRCFILE & "GeneratePOs", Err.Description, SQLBatch)
End Sub

Private Sub cmdApprovePO_Click()
    Dim s As String
    Dim rs As Recordset
    
    
    s = ""
    s = s & "select sum(pretax) + sum(isnull(jctax,0) + isnull(njctax,0))" & vbCrLf
    s = s & "from poitems" & vbCrLf
    s = s & "where ponumber=" & DbQuote(Str, txtPONumber.Text) & vbCrLf
    s = s & "group by ponumber" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If Val("" & rs(0)) > MaxPOAmount Then
        Call MsgBox("This PO exceeds your approval limit.", vbInformation, App.ProductName)
        Exit Sub
    End If
    
    s = ""
    s = s & "update pomaster set status='Approved'" & vbCrLf
    s = s & " ,approvedby=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "where ponumber=" & DbQuote(Str, txtPONumber.Text) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    Call ShowPOStatus("Approved", HFApp.LoginID)

End Sub

Private Sub ShowPOStatus(Status As String, ApprovedBy As String)
    lblApproval.Caption = ""
    cmdApprovePO.Visible = False
    Select Case Status
        Case "", "not yet generated"
            lblApproval.Caption = ""
        Case "Pending"
            lblApproval.FontBold = True
            lblApproval.Caption = "APPROVAL REQUIRED"
            cmdApprovePO.Visible = True
        Case Else
            lblApproval.FontBold = False
            lblApproval.Caption = "Approved by " & ApprovedBy
    End Select

End Sub

Private Sub SetPOStatuses(Batch As Long)
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset

'MsgBox 101
    
    s = ""
    s = s & "update pomaster set status='Approved'" & vbCrLf
    s = s & " ,approvedby=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
    
    s = s & "where ponumber in(" & vbCrLf
    s = s & "  select p.ponumber" & vbCrLf
    s = s & "  from estimateitems e" & vbCrLf
    s = s & "  join pomaster p on e.ponumber=p.ponumber" & vbCrLf
    s = s & "  where e.pogenbatch=" & DbQuote(Num, Batch) & vbCrLf
    s = s & "  group by p.ponumber,p.status" & vbCrLf
    s = s & "  having sum(e.popretax) + sum(isnull(e.pojctax,0) + isnull(e.ponjctax,0)) <= " & DbQuote(Num, MaxPOAmount) & vbCrLf
    s = s & ")" & vbCrLf
    Call HFApp.SqlExec(s, , r)
    
    
'MsgBox 102
    s = ""
    s = s & "select count( distinct p.ponumber)" & vbCrLf
    s = s & "  from estimateitems e" & vbCrLf
    s = s & "  join pomaster p on e.ponumber=p.ponumber" & vbCrLf
    s = s & "  where e.pogenbatch=" & DbQuote(Num, Batch) & vbCrLf
    s = s & "  and p.status = 'Pending'" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    If rs(0) > 0 Then
        Call MsgBox("" & rs(0) & " of these POs exceeds your approval limit. Approval is required before printing or posting.", vbInformation, App.ProductName)
    End If
    
'MsgBox 103

Exit Sub
eh: Call errHandler(SRCFILE & "SetPOStatuses", s)
End Sub

Private Sub ClearItems()
    frmAssembly.Visible = False
    frmPOIndex.Visible = False
    gItems.Rows = 1
    Call Form_Resize
End Sub

Private Sub LoadItems()
On Error GoTo eh

    Dim s As String
    Dim EstAssemblyID As Long
    Dim POIndex       As String
    Dim PONumber      As String
    Dim POVendor      As String
    
    
    If Not SaveData(True) Then Exit Sub
    If gAssemblies.Row < 1 Then Exit Sub
    
    mChangeOrder = ""
    mCustomerNo = ""
    
    Select Case gAssemblies.Cell(flexcpData, gAssemblies.Row, 1)
        Case "EstAssemblyID":  EstAssemblyID = gAssemblies.Cell(flexcpText, gAssemblies.Row, 1)
        Case "POIndex":        POIndex = gAssemblies.Cell(flexcpText, gAssemblies.Row, 1)
        Case "PONumber":       PONumber = gAssemblies.Cell(flexcpText, gAssemblies.Row, 1)
        Case "POVendor":       POVendor = gAssemblies.Cell(flexcpText, gAssemblies.Row, 1)
        
        Case "CORSort", "ChangeOrder":
            mChangeOrder = gAssemblies.Cell(flexcpText, gAssemblies.Row, 1)
            If Left(mChangeOrder, 6) = "ZZZZZ" & Chr(1) Then mChangeOrder = ""                'Mid(mChangeOrder, 7)
            mHFLocation = Val(gAssemblies.Cell(flexcpText, gAssemblies.Row, gAssemblies.ColIndex("hflocation")))
            mCustomerNo = gAssemblies.Cell(flexcpText, gAssemblies.GetNodeRow(gAssemblies.Row, flexNTParent), 1)
        
        Case "PONumberVendor"
            s = gAssemblies.Cell(flexcpText, gAssemblies.Row, 1)
            PONumber = Parse(s, 1, Chr(2))
            POVendor = Parse(s, 2, Chr(2))
    End Select
    If EstAssemblyID = 0 And POIndex = "" Then POIndex = Parse(Parse(gAssemblies.RowData(gAssemblies.Row), 2, "POIndex = '"), 1, "' AND")
    
    
    frmAssembly.Visible = EstAssemblyID <> 0
    frmPOIndex.Visible = PONumber <> "" Or POIndex <> ""
    frmPOIndex.Height = IIf(PONumber = "", 2160, 2955)
    frmChangeOrder.Visible = mChangeOrder <> ""
    frmChangeOrder.Visible = False

    txtDeliveryRecipient.Visible = PONumber <> ""
    txtDeliveryAddress.Visible = PONumber <> ""
    
    Call Form_Resize
    
    Select Case True
        Case frmEstItems.Visible: Call LoadEstItems
        Case frmRFPs.Visible:     Call LoadRFQ
    End Select
    Dirty = False
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadItems", s)
End Sub

Private Sub LoadRFQ()
On Error GoTo eh
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset

    gBids.Rows = 1
    gBidItems.Rows = 1
    
    mRFP = Val(gAssemblies.TextMatrix(gAssemblies.Row, 1))
    If mRFP = 0 Then Exit Sub
    
    
    txtRFPDescription.Enabled = True
    txtRFPComments.Enabled = True
    gBids.Enabled = True
    gBidItems.Enabled = True
    
    'load rfp items
    s = "SELECT * FROM RFPs WHERE RFP=" & DbQuote(Num, mRFP)
    Set rs = HFApp.SqlExec(s)
    txtRFPDescription.Text = "" & rs("Description")
    txtRFPComments.Text = "" & rs("Comments")
    
    
        
    'load rfp items
    With gBidItems
        .Redraw = flexRDNone
        gBidItems.Cols = 64
        
        s = "SELECT * FROM EstimatedItems WHERE RFP=" & DbQuote(Num, mRFP) & " and DivisionID = " & HFApp.DivisionID
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            
            'ignore deleted items
            If (mMode <> fcePO And Not rs("BudgetDeleted")) Or _
               (mMode = fcePO And Not rs("PODeleted")) Then
                
                .AddItem ""
                r = .Rows - 1
                .Cell(flexcpText, r, .ColIndex("RFP")) = mRFP
                For c = 0 To .Cols - 1
                    Select Case .ColKey(c)
                        Case "TakeoffQty"
                            'takeoff qty in db always shows original takeoff amounts
                            'on screen we want it to show the budget or po qty in takeoff uom
                            If Val("" & rs("ConversionFactor")) = 0 Then
                                .Cell(flexcpText, r, c) = Val("" & rs(IIf(mMode <> fcePO, "BudgetQty", "POQty")))
                            Else
                                .Cell(flexcpText, r, c) = Val("" & rs(IIf(mMode <> fcePO, "BudgetQty", "POQty"))) / Val("" & rs("ConversionFactor"))
                            End If
                        Case "POPretax"
                            .TextMatrix(0, c) = IIf(mMode <> fcePO, "Budget Unit Rate", "PO Unit Rate")
                            .Cell(flexcpText, r, c) = Val("" & rs(IIf(mMode <> fcePO, "BudgetRate", "PORate")))
                        Case Else
                            On Error Resume Next
                            .Cell(flexcpText, r, c) = "" & rs(.ColKey(c))
                            On Error GoTo eh
                    End Select
                Next
                
            End If
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
    End With
    
    
    'load bidders
    With Me.gBids
        .Redraw = flexRDNone
        
        
        s = ""
        s = s & "select v.vendor_id vendor" & vbCrLf
        s = s & "      ,v.vendor_name company" & vbCrLf
        s = s & "      ,b.sentdate" & vbCrLf
        s = s & "      ,b.expirydate" & vbCrLf
        s = s & "      ,b.comments" & vbCrLf
        s = s & "  from vendorbids b" & vbCrLf
        s = s & "       join tblvendors v on (v.vendor_id=b.vendor and v.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
        s = s & " where b.RFP=" & DbQuote(Num, mRFP) & vbCrLf
        s = s & "order by v.vendor_name" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .Cell(flexcpText, r, .ColIndex("Vendor")) = "" & rs("Vendor")
            .Cell(flexcpText, r, .ColIndex("VendorName")) = "" & rs("Company")
            .Cell(flexcpPicture, r, .ColIndex("VendorName")) = FMain.SmallIcons.ListImages.Item("vendor").Picture
            .Cell(flexcpText, r, .ColIndex("SentDate")) = format("" & rs("SentDate"), "short date")
            .Cell(flexcpText, r, .ColIndex("ExpiryDate")) = format("" & rs("ExpiryDate"), "short date")
            
            c = gBidItems.Cols
            gBidItems.Cols = c + 1
            gBidItems.ColData(c) = "VendorRate"
            gBidItems.ColHidden(c) = False
            gBidItems.ColKey(c) = "vdR~" & .Cell(flexcpText, r, .ColIndex("Vendor"))
            gBidItems.TextMatrix(0, c) = .Cell(flexcpText, r, .ColIndex("VendorName"))
            
            
            c = gBidItems.Cols
            gBidItems.Cols = c + 1
            gBidItems.ColData(c) = "VendorComments"
            gBidItems.ColHidden(c) = True
            gBidItems.ColKey(c) = "vdC~" & .Cell(flexcpText, r, .ColIndex("Vendor"))
            
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
    End With
    
    
    'load bids
    s = "select * from vendorbiditems where RFP=" & DbQuote(Num, mRFP) & " order by estitemid"
    With gBidItems
        .Redraw = flexRDNone
        Set rs = HFApp.SqlExec(s, dbHomefront)
        While Not rs.EOF
            r = .FindRow(Val("" & rs("EstItemID")), , .ColIndex("EstItemID"), , True)
            If r <> -1 Then
            
                'comments
                c = .ColIndex("vdC~" & rs("Vendor"))
                If c <> -1 Then
                    .Cell(flexcpText, r, c) = "" & rs("Comment")
                End If
                
                'rate and status
                c = .ColIndex("vdR~" & rs("Vendor"))
                If c <> -1 Then
                
                    .Cell(flexcpText, r, c) = rs("Rate")
                    If .Cell(flexcpText, r, c) <> "" Then .Cell(flexcpText, r, c) = format(.Cell(flexcpText, r, c), "#,##0.00")
                    
                    
                    .Cell(flexcpData, r, c) = "" & rs("Status")
                     
                    If .Cell(flexcpData, r, c) = bsAccepted Then
                        .TextMatrix(r, .ColIndex("BidStatus")) = "" & rs("Vendor")
                    End If
                    If "" & rs("Comment") <> "" Then
                        .Cell(flexcpPicture, r, c) = FMain.SmallIcons.ListImages("cellcomments").Picture
                        .Cell(flexcpPictureAlignment, r, c) = flexPicAlignRightTop
                    End If
                End If
                
                
            End If
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
        On Error Resume Next
        .Row = 1
        .RowSel = .Rows - 1
    End With
    
    Call UpdateBidTotals


Exit Sub
eh: Call errHandler(SRCFILE & "LoadRFQ", s)
End Sub

Private Sub LoadEstItems()
On Error GoTo eh
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    Dim EstAssemblyID As Long
    Dim b As Boolean
    Dim POIndex       As String
    Dim PONumber      As String
    Dim POVendor      As String
    If gAssemblies.Row < 1 Then Exit Sub
    
    Select Case gAssemblies.Cell(flexcpData, gAssemblies.Row, 1)
        Case "EstAssemblyID":  EstAssemblyID = gAssemblies.TextMatrix(gAssemblies.Row, 1)
        Case "POIndex":        POIndex = gAssemblies.TextMatrix(gAssemblies.Row, 1)
        Case "PONumber":       PONumber = gAssemblies.TextMatrix(gAssemblies.Row, 1)
        Case "POVendor":       POVendor = gAssemblies.TextMatrix(gAssemblies.Row, 1)
        Case "PONumberVendor"
            s = gAssemblies.TextMatrix(gAssemblies.Row, 1)
            PONumber = Parse(s, 1, Chr(2))
            POVendor = Parse(s, 2, Chr(2))
    End Select
    If EstAssemblyID = 0 And POIndex = "" Then POIndex = Parse(Parse(gAssemblies.RowData(gAssemblies.Row), 2, "POIndex = '"), 1, "' AND")
    
    mEstAssemblyID = -1
    mAssemblyChanged = False
    
    
    If Me.frmChangeOrder.Visible Then
        s = ""
        s = s & "SELECT *" & vbCrLf
        s = s & "  FROM changeordermaster" & vbCrLf
        s = s & " WHERE customer_no=" & DbQuote(Str, mCustomerNo) & vbCrLf
        s = s & "   AND change_order_no=" & DbQuote(Str, mChangeOrder) & vbCrLf
        Set rs = HFApp.SqlExec(s)
        If rs.EOF Then
            frmChangeOrder.Visible = False
            Call Form_Resize
        Else
            txtCONumber.Text = "" & rs("change_order_no")
            txtCODate.Text = format("" & rs("change_date"), "medium date")
            txtCODateApproved.Text = format("" & rs("bldr_approved_date"), "medium date")
            txtCOApprovedBy.Text = "" & rs("bldr_approved_by")
            txtCOComments.Text = "" & rs("Comments")
            
            b = "" & rs("PurchasingCO") <> "True"
            txtCONumber.Locked = True
            txtCODate.Locked = b
            txtCODateApproved.Locked = b
            txtCOApprovedBy.Locked = b
            txtCOComments.Locked = b
            
        End If
    End If
    
    
    If frmAssembly.Visible Then
        mEstAssemblyID = EstAssemblyID
        s = ""
        s = s & "SELECT a.*" & vbCrLf
        s = s & "      ,isnull(co.Change_date,a.CreatedDate) ChangeOrderDate" & vbCrLf
        s = s & "      ,isnull(nullif(coa.user_name,''),coa.user_id) ChangeOrderApprovedBy" & vbCrLf
        s = s & "      ,co.bldr_approved_date ChangeOrderApprovedDate" & vbCrLf
        s = s & "      ,case when aa.TotalAmount is null then a.SalesRate else aa.TotalAmount end Price,c.Description HFCatDesc" & vbCrLf
        s = s & "  FROM EstimateAssemblies a" & vbCrLf
        s = s & "  LEFT OUTER JOIN tblcategories c on a.hfcategory=c.category" & vbCrLf
        s = s & "  LEFT OUTER JOIN changeordermaster co on (a.Customer_No=co.Customer_No and a.changeorder=co.change_order_no)" & vbCrLf
        s = s & "  LEFT OUTER JOIN user_manager coa on (co.bldr_approved_by=coa.user_id)" & vbCrLf
        s = s & "  LEFT OUTER JOIN addons aa on (a.EstAssemblyID=aa.AssemblyID and aa.basis='Total')" & vbCrLf
        s = s & " WHERE a.EstAssemblyID=" & DbQuote(Num, EstAssemblyID)
        
        
        
        Set rs = HFApp.SqlExec(s)
        
        If Not rs.EOF Then
            txtHFOption.Text = "" & rs("Assembly")
            txtHFCategory.Text = "" & rs("HFCatDesc")
            txtHFDescription.Text = "" & rs("HFDescription")
            
            lblQuantity.Caption = "" & rs("SalesQty")
            lblUOM.Caption = " " & rs("AssemblyUOM")
            
            txtQuantity.Locked = "" & HFApp.Options.ValueByName("LetPurchaserChangeSaleQty") <> "True"
            txtQuantity.Text = "" & rs("SalesQty") & " " & rs("AssemblyUOM")
            txtHFComments.Text = "" & rs("HFComments")
            txtEstimatorNotes.Text = "" & rs("EstimatorNotes")
            txtJCExtra.Text = "" & rs("JCExtra")
            
            txtColor.Text = "" & rs("Color")
            txtLocation.Text = "" & rs("Location")
            
            txtStyle.Text = "" & rs("Style")
            txtOther.Text = "" & rs("Other")
            txtFinish.Text = "" & rs("Finish")
            
            frmChangeRequest.Visible = True 'HFApp.Options.ValueByName("BuilderType") = "Commercial" '"" & rs("changeorder") = "" And "" & rs("ChangeRequestStatus") <> ""
            'Button to display addons
            cmdBrowse(9).Visible = "" & rs("changeorder") = "" And "" & rs("ChangeRequestStatus") <> ""
            'Status combobox
            cboChangeRequestStatus.Visible = "" & rs("changeorder") = "" And "" & rs("ChangeRequestStatus") <> ""
            'Status label
            Label18.Visible = "" & rs("changeorder") = "" And "" & rs("ChangeRequestStatus") <> ""
            txtChangeRequestedBy.Text = "" & rs("ChangeRequestedBy")
            txtChangeRequestedDate.Text = IIf("" & rs("ChangeRequestedDate") = "", "", format("" & rs("ChangeRequestedDate"), "medium date"))
            txtChangeApprovedBy.Text = "" & rs("ChangeApprovedBy")
            txtChangeApprovedDate.Text = IIf("" & rs("ChangeApprovedDate") = "", "", format("" & rs("ChangeApprovedDate"), "medium date"))
            Call SetListIndex(cboChangeRequestStatus, , "" & rs("ChangeRequestStatus"))
            
            txtPrice.Text = format("" & rs("Price"), "#,##0.00")
            
            Call BudgetsAreLocked("" & rs("BudgetsLocked") = "True", "" & rs("BudgetsLockedDate"), "" & rs("BudgetsLockedBy"))
            
        End If
    Else
        Call BudgetsAreLocked(False, "", "")
    End If
    If HFApp.UserPermission("IssuePOs") = False And mMode = fcePO Then
        Toolbar.Buttons("SaveAssembly").Enabled = False
    Else
        Toolbar.Buttons("SaveAssembly").Enabled = frmAssembly.Visible
    End If
    Toolbar.Buttons("Preview").Enabled = PONumber <> ""
    If frmPOIndex.Visible Then
        
        If PONumber <> "" Then
            s = ""
            s = s & "SELECT PONumber" & vbCrLf
            s = s & "      ,PostingBatch, Batches.TStmp PostingDate" & vbCrLf
            s = s & "      ,vendor_id" & vbCrLf
            s = s & "      ,ISNULL(Vendor_Name,'unknown vendor') Vendor_Name" & vbCrLf
            s = s & "      ,POIndex,Description,StandardText,ShipVia,FOB,Terms,OrderedBy,HideQty,HidePrice,TotalOnly" & vbCrLf
            s = s & "      ,DeliveryRecipient,DeliveryAddress,DeliveryDate,DueDate,DeliveryMethod,PODate" & vbCrLf
            s = s & "      ,RetainagePercent" & vbCrLf
            s = s & "      ,IncludeDocuments" & vbCrLf
            s = s & "      ,status,approvedby,approveddate" & vbCrLf
            s = s & "  FROM POMaster" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblVendors ON (Vendor=Vendor_ID and POmaster.DivisionID = tblVendors.DivisionID)" & vbCrLf
            s = s & "       LEFT OUTER JOIN Batches ON (PostingBatch=Batch)" & vbCrLf
            s = s & " WHERE POMaster.DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, Parse(PONumber, 1, Chr(2))) & vbCrLf
        Else
            s = ""
            s = s & "SELECT '' PONumber" & vbCrLf
            s = s & "      ,'' PostingBatch, NULL PostingDate" & vbCrLf
            s = s & "      ,v.Vendor_id" & vbCrLf
            s = s & "      ,ISNULL(v.Vendor_Name,'unknown vendor') Vendor_Name" & vbCrLf
            s = s & "      ,i.POIndex" & vbCrLf
            s = s & "      ,case when p.poindex is null then isnull(nullif(i.Description,''),i.poindex) else p.Description end Description" & vbCrLf
            s = s & "      ,case when p.poindex is null then i.StandardText else p.StandardText end StandardText " & vbCrLf
            s = s & "      ,case when p.poindex is null then i.ShipVia else p.ShipVia end ShipVia " & vbCrLf
            s = s & "      ,case when p.poindex is null then i.FOB else p.FOB end FOB " & vbCrLf
            s = s & "      ,case when p.poindex is null then i.Terms else p.Terms end Terms " & vbCrLf
            s = s & "      ,p.OrderedBy " & vbCrLf
            s = s & "      ,case when p.poindex is null then i.HideQty else p.HideQty end HideQty " & vbCrLf
            s = s & "      ,case when p.poindex is null then i.HidePrice else p.HidePrice end HidePrice" & vbCrLf
            s = s & "      ,case when p.poindex is null then i.TotalOnly else p.TotalOnly end TotalOnly " & vbCrLf
            s = s & "      ,'' DeliveryRecipient,'' DeliveryAddress,'' DeliveryMethod,NULL DeliveryDate,NULL DueDate,NULL PODate" & vbCrLf
            s = s & "      ,case when p.poindex is null then i.RetainagePercent else p.RetainagePercent end RetainagePercent" & vbCrLf
            s = s & "      ,'True' IncludeDocuments" & vbCrLf
            s = s & "      ,'not yet generated' status,'' approvedby,null approveddate" & vbCrLf
            s = s & "  FROM tblPOIndex i" & vbCrLf
            s = s & "       LEFT OUTER JOIN JobPOIndex p ON(p.DivisionID = i.DivisionID and i.poindex=p.poindex and p.Job_No=" & DbQuote(Str, mJob) & ")"
            If POVendor = "" Then
                s = s & "       LEFT OUTER JOIN tblVendors v ON(v.DivisionID = " & HFApp.DivisionID & " and v.Vendor_ID=dbo.Purch_GetCommunityVendor(p.Community, p.POIndex," & HFApp.DivisionID & "))" & vbCrLf
            Else
                s = s & "       LEFT OUTER JOIN tblVendors v ON(v.DivisionID = " & HFApp.DivisionID & " and v.Vendor_ID=" & DbQuote(Str, POVendor) & ")" & vbCrLf
            End If
            s = s & " WHERE i.DivisionID = " & HFApp.DivisionID & " and i.POIndex=" & DbQuote(Str, POIndex) & vbCrLf
        End If
        
        
        lblPOHistory.Caption = ""
        cmdApprovePO.Visible = False
        txtRetainagePercent.Enabled = True
        txtDeliveryRecipient.Enabled = True
        txtDeliveryAddress.Enabled = True
        
        Call ShowPOStatus("", "")
        
        Set rs = HFApp.SqlExec(s)
        If rs.EOF Then
            txtVendor.Text = ""
            txtVendor.Tag = ""
            txtPONumber.Text = ""
            txtPOIndex.Text = ""
            txtPODescription.Text = ""
            txtStandardText.Text = ""
            txtShipVia.Text = ""
            txtFOB.Text = ""
            txtTerms.Text = ""
            txtOrderedBy.Text = ""
            txtDeliveryRecipient.Text = ""
            txtDeliveryAddress.Text = ""
            txtDateIssued.Text = ""
            txtRetainagePercent.Text = ""
        Else
            
            s = ""
            If Val("" & rs("PostingBatch")) <> 0 Then
                txtRetainagePercent.Enabled = False
                s = s & "Posted " & format(rs("PostingDate"), "mmm d, yyyy") & "  ( batch " & rs("PostingBatch") & " )" & vbCrLf
            End If
            If "" & rs("DeliveryDate") <> "" Then
                txtDeliveryRecipient.Enabled = False
                txtDeliveryAddress.Enabled = True
                Select Case Val("" & rs("DeliveryMethod"))
                    Case dtPrint: s = s & "Printed " & format("" & rs("DeliveryDate"), "mmm d, yyyy") & vbCrLf
                    Case dtEmail: s = s & "Emailed " & format("" & rs("DeliveryDate"), "mmm d, yyyy") & vbCrLf
                    Case dtFax:   s = s & "Faxed " & format("" & rs("DeliveryDate"), "mmm d, yyyy") & vbCrLf
                End Select
            End If
            Call txtDeliveryAddress_Change
            lblPOHistory.Caption = s
            lblPOHistory.AutoSize = True
            
            txtVendor.Text = "" & rs("Vendor_Name")
            txtVendor.Tag = "" & rs("Vendor_id")
            txtVendor.Enabled = False
            
            Call ShowPOStatus("" & rs("Status"), "" & rs("ApprovedBy"))
            txtDateIssued.Text = "" & rs("PODate")
            
            txtPONumber.Text = "" & rs("PONumber")
            txtPOIndex.Text = "" & rs("POIndex")
            txtPODescription.Text = "" & rs("Description")
            txtStandardText.Text = "" & rs("StandardText")
            txtShipVia.Text = "" & rs("ShipVia")
            txtFOB.Text = "" & rs("FOB")
            txtTerms.Text = "" & rs("Terms")
            txtOrderedBy.Text = "" & rs("OrderedBy")
            chkHideQty.value = IIf("" & rs("HideQty") = "True", vbChecked, vbUnchecked)
            chkHidePrice.value = IIf("" & rs("HidePrice") = "True", vbChecked, vbUnchecked)
            chkTotalOnly.value = IIf("" & rs("TotalOnly") = "True", vbChecked, vbUnchecked)
            chkIncludeDocuments.value = IIf("" & rs("IncludeDocuments") = "True", vbChecked, vbUnchecked)
            txtDeliveryRecipient.Text = "" & rs("DeliveryRecipient")
            txtDeliveryAddress.Text = "" & rs("DeliveryAddress")
            txtRetainagePercent.Text = Val("" & rs("RetainagePercent")) & "%"
        End If
    End If
        
    With gItems
        .Redraw = flexRDNone
        .Rows = 1
        
        
        'EstItemID will be null for custom options or anything that has no assembly...
        s = ""
        s = s & "SELECT *" & vbCrLf
        s = s & "  FROM EstimatedItems i" & vbCrLf
        s = s & " WHERE (NOT i.EstItemID IS NULL)" & vbCrLf
        s = s & "   AND " & Mid(Replace(gAssemblies.RowData(gAssemblies.Row), "AssemblyJob_No", "Job_No"), 6) & vbCrLf
        s = s & "   AND i.DivisionID = " & HFApp.DivisionID
        If ProjectBased Then
            s = s & "   AND i.isquote=0" & vbCrLf
        End If
        
        
        
        Set rs = HFApp.SqlExec(s, dbHomefront)

        While Not rs.EOF
        
        
            'ignore deleted items
            If (mMode = fceQuote And Not rs("BudgetDeleted")) Or (mMode = fceBudget And Not rs("BudgetDeleted")) Or ((Not IsIn(mMode, fceBudget, fceQuote)) And Not rs("PODeleted")) Then
                
                .AddItem ""
                r = .Rows - 1
                
                On Error Resume Next
                .TextMatrix(r, .ColIndex("Invertible")) = "" & rs("Invertible")
                .TextMatrix(r, .ColIndex("Customer_No")) = "" & rs("Customer_No")
                .TextMatrix(r, .ColIndex("CustomerDesc")) = "" & rs("CustomerDesc")
                .TextMatrix(r, .ColIndex("Community")) = "" & rs("Community")
                .TextMatrix(r, .ColIndex("IsQuote")) = "" & rs("IsQuote")
                .TextMatrix(r, .ColIndex("CommunityDesc")) = "" & rs("CommunityDesc")
                .TextMatrix(r, .ColIndex("CommunityPhase")) = "" & rs("CommunityPhase")
                .TextMatrix(r, .ColIndex("AssemblyJob_No")) = "" & rs("AssemblyJob_No")
                .TextMatrix(r, .ColIndex("Job_No")) = "" & rs("Job_No")
                .TextMatrix(r, .ColIndex("Sequence")) = "" & rs("Sequence")
                .TextMatrix(r, .ColIndex("JobDesc")) = "" & rs("JobDesc")
                .TextMatrix(r, .ColIndex("EstimateIndex")) = "" & rs("EstimateIndex")
                .TextMatrix(r, .ColIndex("OptionType")) = "" & rs("OptionType")
                .TextMatrix(r, .ColIndex("Unit")) = "" & rs("Unit")
                .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
                .TextMatrix(r, .ColIndex("AssemblyDescription")) = "" & rs("AssemblyDescription")
                .TextMatrix(r, .ColIndex("OptionTypeDesc")) = "" & rs("OptionTypeDesc")
                .TextMatrix(r, .ColIndex("AssemblyType")) = "" & rs("AssemblyType")
                .TextMatrix(r, .ColIndex("AssemblyTypeDesc")) = "" & rs("AssemblyTypeDesc")
                .TextMatrix(r, .ColIndex("HFLocation")) = "" & rs("HFLocation")
                .TextMatrix(r, .ColIndex("HFLocationDesc")) = "" & rs("HFLocationDesc")
                .TextMatrix(r, .ColIndex("ChangeOrder")) = "" & rs("ChangeOrder")
                .TextMatrix(r, .ColIndex("Seq")) = "" & rs("Seq")
                .TextMatrix(r, .ColIndex("POGroup")) = "" & rs("POGroup")
                .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
                .TextMatrix(r, .ColIndex("POIndexDescription")) = "" & rs("POIndexDescription")
                .TextMatrix(r, .ColIndex("POFormat")) = "" & rs("POFormat")
                .TextMatrix(r, .ColIndex("PONumber")) = "" & rs("PONumber")
                .TextMatrix(r, .ColIndex("POStatus")) = "" & rs("POStatus")
                .TextMatrix(r, .ColIndex("PONumberVendor")) = "" & rs("PONumberVendor")
                .TextMatrix(r, .ColIndex("BudgetVendor")) = "" & rs("BudgetVendor")
                .TextMatrix(r, .ColIndex("BudgetVendorName")) = "" & rs("BudgetVendorName")
                .TextMatrix(r, .ColIndex("POVendor")) = "" & rs("POVendor")
                .TextMatrix(r, .ColIndex("POVendorName")) = "" & rs("POVendorName")
                .TextMatrix(r, .ColIndex("JCExtra")) = "" & rs("JCExtra")
                .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
                .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = "" & rs("JCCostCodeDesc")
                .TextMatrix(r, .ColIndex("JCCostCodeGroup")) = "" & rs("JCCostCodeGroup")
                .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
                .TextMatrix(r, .ColIndex("JCCategoryDesc")) = "" & rs("JCCategoryDesc")
                .TextMatrix(r, .ColIndex("VarianceJCCategory")) = "" & rs("VarianceJCCategory")
                .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = "" & rs("VarianceJCCategoryDesc")
                .TextMatrix(r, .ColIndex("VariancePretax")) = "" & rs("VariancePretax")
                .TextMatrix(r, .ColIndex("VarianceNJCTax")) = "" & rs("VarianceNJCTax")
                .TextMatrix(r, .ColIndex("VarianceJCTax")) = "" & rs("VarianceJCTax")
                .TextMatrix(r, .ColIndex("HFCategory")) = "" & rs("HFCategory")
                .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
                .TextMatrix(r, .ColIndex("OptionID")) = "" & rs("OptionID")
                .TextMatrix(r, .ColIndex("HFDescription")) = "" & rs("HFDescription")
                .TextMatrix(r, .ColIndex("HFComments")) = "" & rs("HFComments")
                .TextMatrix(r, .ColIndex("SalesWorksheet")) = "" & rs("SalesWorksheet")
                .TextMatrix(r, .ColIndex("EstItemID")) = "" & rs("EstItemID")
                .TextMatrix(r, .ColIndex("EstAssemblyID")) = "" & rs("EstAssemblyID")
                .TextMatrix(r, .ColIndex("EstPhase")) = "" & rs("EstPhase")
                .TextMatrix(r, .ColIndex("EstItem")) = "" & rs("EstItem")
                .TextMatrix(r, .ColIndex("SortOrder")) = "" & rs("SortOrder")
                .TextMatrix(r, .ColIndex("ItemDesc")) = "" & rs("ItemDesc")
                .TextMatrix(r, .ColIndex("ItemComments")) = "" & rs("ItemComments")
                .TextMatrix(r, .ColIndex("TakeoffQty")) = "" & rs("TakeoffQty")
                .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
                .TextMatrix(r, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
                .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
                .TextMatrix(r, .ColIndex("BudgetQty")) = "" & rs("BudgetQty")
                .TextMatrix(r, .ColIndex("BudgetRate")) = "" & rs("BudgetRate")
                .TextMatrix(r, .ColIndex("BudgetPretax")) = "" & rs("BudgetPretax")
                .TextMatrix(r, .ColIndex("BudgetTaxGroup")) = "" & rs("BudgetTaxGroup")
                .TextMatrix(r, .ColIndex("BudgetJCTAX")) = "" & rs("BudgetJCTAX")
                .TextMatrix(r, .ColIndex("BudgetJCTaxRate")) = "" & rs("BudgetJCTaxRate")
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = "" & rs("BudgetNJCTax")
                .TextMatrix(r, .ColIndex("BudgetNJCTaxRate")) = "" & rs("BudgetNJCTaxRate")
                .TextMatrix(r, .ColIndex("BudgetPostingBatch")) = "" & rs("BudgetPostingBatch")
                .TextMatrix(r, .ColIndex("BudgetGenerated")) = "" & rs("BudgetGenerated")
                .TextMatrix(r, .ColIndex("POGenBatch")) = "" & rs("POGenBatch")
                                
                .TextMatrix(r, .ColIndex("Budget1Rate")) = "" & rs("BudgetRate1")
                .TextMatrix(r, .ColIndex("Budget1Qty")) = "" & rs("BudgetQty1")
                .TextMatrix(r, .ColIndex("Budget2Rate")) = "" & rs("BudgetRate2")
                .TextMatrix(r, .ColIndex("Budget2Qty")) = "" & rs("BudgetQty2")
                .TextMatrix(r, .ColIndex("Budget3Rate")) = "" & rs("BudgetRate3")
                .TextMatrix(r, .ColIndex("Budget3Qty")) = "" & rs("BudgetQty3")
                
                .TextMatrix(r, .ColIndex("POQty")) = "" & rs("POQty")
                .TextMatrix(r, .ColIndex("PORate")) = "" & rs("PORate")
                .TextMatrix(r, .ColIndex("POPretax")) = "" & rs("POPretax")
                .TextMatrix(r, .ColIndex("POTaxGroup")) = "" & rs("POTaxGroup")
                .TextMatrix(r, .ColIndex("POJCTAX")) = "" & rs("POJCTAX")
                .TextMatrix(r, .ColIndex("POJCTaxRate")) = "" & rs("POJCTaxRate")
                .TextMatrix(r, .ColIndex("PONJCTax")) = "" & rs("PONJCTax")
                .TextMatrix(r, .ColIndex("PONJCTaxRate")) = "" & rs("PONJCTaxRate")
                .TextMatrix(r, .ColIndex("POOverridden")) = "" & rs("POOverridden")
                .TextMatrix(r, .ColIndex("BudgetOverridden")) = "" & rs("BudgetOverridden")
                .Cell(flexcpChecked, r, .ColIndex("ExcludeFromPO")) = IIf(rs("ExcludeFromPO"), flexChecked, flexUnchecked)
                .TextMatrix(r, .ColIndex("BudgetDeleted")) = "" & rs("BudgetDeleted")
                .TextMatrix(r, .ColIndex("PODeleted")) = "" & rs("PODeleted")
                .TextMatrix(r, .ColIndex("PartNumber")) = "" & rs("PartNumber")
                .TextMatrix(r, .ColIndex("RoundDir")) = "" & rs("RoundDir")
                .TextMatrix(r, .ColIndex("RoundTo")) = "" & rs("RoundTo")
                .TextMatrix(r, .ColIndex("WastePercent")) = "" & rs("WastePercent")
                .TextMatrix(r, .ColIndex("EstPhaseDesc")) = "" & rs("EstPhaseDesc")
                .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
                .TextMatrix(r, .ColIndex("SalesQty")) = "" & rs("SalesQty")
                
                For c = 1 To 40
                    If "" & rs("wbsdesc" & format(c, "00")) <> "" Then
                        .TextMatrix(0, .ColIndex("wbsdesc" & format(c, "00"))) = "" & rs("wbsdesc" & format(c, "00"))
                        .TextMatrix(r, .ColIndex("wbs" & format(c, "00"))) = "" & rs("wbs" & format(c, "00"))
                    End If
                Next
                
                .TextMatrix(r, .ColIndex("RFP")) = "" & rs("RFP")
                .TextMatrix(r, .ColIndex("Formula")) = "" & rs("Formula")
                .TextMatrix(r, .ColIndex("ItemNumber")) = "" & rs("ItemNumber")
                .TextMatrix(r, .ColIndex("RFPGuid")) = "" & rs("RFPGuid")
                .TextMatrix(r, .ColIndex("RFPTitle")) = "" & rs("RFPTitle")
                .TextMatrix(r, .ColIndex("RFPDesc")) = "" & rs("RFPDesc")
                .TextMatrix(r, .ColIndex("RFPComments")) = "" & rs("RFPComments")
                .TextMatrix(r, .ColIndex("BudgetsLocked")) = "" & rs("BudgetsLocked")
                .TextMatrix(r, .ColIndex("BudgetsLockedBy")) = "" & rs("BudgetsLockedBy")
                .TextMatrix(r, .ColIndex("BudgetsLockedDate")) = "" & rs("BudgetsLockedDate")
                .TextMatrix(r, .ColIndex("ChangeRequestStatus")) = "" & rs("ChangeRequestStatus")
                .TextMatrix(r, .ColIndex("ChangeRequestedBy")) = "" & rs("ChangeRequestedBy")
                .TextMatrix(r, .ColIndex("ChangeRequestedDate")) = "" & rs("ChangeRequestedDate")
                .TextMatrix(r, .ColIndex("ChangeApprovedBy")) = "" & rs("ChangeApprovedBy")
                .TextMatrix(r, .ColIndex("ChangeApprovedDate")) = "" & rs("ChangeApprovedDate")
                .TextMatrix(r, .ColIndex("CORSort")) = "" & rs("CORSort")
                .TextMatrix(r, .ColIndex("CORNumber")) = "" & rs("CORNumber")
                .TextMatrix(r, .ColIndex("IsChangeRequest")) = "" & rs("IsChangeRequest")
                .TextMatrix(r, .ColIndex("IsChange")) = "" & rs("IsChange")
                .TextMatrix(r, .ColIndex("BillingItemID")) = "" & rs("BillingItemID")
                .TextMatrix(r, .ColIndex("HoldbackRate")) = "" & rs("HoldbackRate")
                .TextMatrix(r, .ColIndex("ARTaxGroup")) = "" & rs("ARTaxGroup")
                .TextMatrix(r, .ColIndex("RevenueAccount")) = "" & rs("RevenueAccount")
                .TextMatrix(r, .ColIndex("BudgetPostingDate")) = "" & rs("BudgetPostingDate")
                .TextMatrix(r, .ColIndex("DivisionID")) = "" & rs("DivisionID")
                .TextMatrix(r, .ColIndex("ExternalCostCode")) = "" & rs("ExternalCostCode")
                .TextMatrix(r, .ColIndex("ExternalCategory")) = "" & rs("ExternalCategory")
                .TextMatrix(r, .ColIndex("QuoteAttachmentID")) = "" & rs("QuoteAttachmentID")
                .TextMatrix(r, .ColIndex("ReversingItemID")) = "" & rs("ReversingItemID")
                .TextMatrix(r, .ColIndex("CorrectingItemID")) = "" & rs("CorrectingItemID")
                .TextMatrix(r, .ColIndex("IsReversingItem")) = "" & rs("IsReversingItem")
                .TextMatrix(r, .ColIndex("IsCorrectingItem")) = "" & rs("IsCorrectingItem")
                
                On Error GoTo eh
                
                
                Dim t As String
                Select Case True
                    
                    'original - cancelled
                    Case .ValueMatrix(r, .ColIndex("ReversingItemID")) <> 0
                         t = "Original - Cancelled"
                    
                    'original - reversing
                    Case .TextMatrix(r, .ColIndex("IsReversingItem")) = "True"
                         t = "Original - Reversal"
                    
                    'original - correcting
                    Case .TextMatrix(r, .ColIndex("IsCorrectingItem")) = "True"
                         t = "Original - Correction"
                    
                    'original
                    Case .TextMatrix(r, .ColIndex("BudgetDeleted")) = "False" And .TextMatrix(r, .ColIndex("PODeleted")) = "False"
                         t = "Original"

                    'budget only
                    Case .TextMatrix(r, .ColIndex("PODeleted")) = "True"
                         t = "Budget Only"
                    
                    'po only
                    Case .TextMatrix(r, .ColIndex("BudgetDeleted")) = "True"
                         t = "PO Only"
                    
                                        
                    Case Else
                         t = "???"
                         
                End Select
                .TextMatrix(r, .ColIndex("RowType")) = t
                
                
                .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                
                
                'Why are items that are part of a change request not available to be put on a po?
                If .TextMatrix(r, .ColIndex("IsChangeRequest")) = "1" Then
                    .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexNoCheckbox
                Else
                    .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexUnchecked
                End If
            End If
            rs.MoveNext
        Wend
        
        Call ShowCancelledItems
        Call ColorizeItems(-1)
        Call GroupGrid


        .Redraw = flexRDBuffered
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadEstItems", s)
End Sub

Private Sub ShowCancelledItems()
    Dim r As Long
    Dim i As Long
    Dim hidden As Boolean
    
    hidden = chkCancelledItems.value = vbChecked
    With gItems
        For r = 1 To .Rows - 1
            'only apply to none deleted items
            If .TextMatrix(r, .ColIndex("PODeleted")) <> "True" Then
                If IsIn(.TextMatrix(r, .ColIndex("RowType")), "Original - Cancelled", "Original - Reversal") Then
'                    If r = 8 Then Stop
                    
                    i = .GetNodeRow(r, flexNTParent)
                    If i > 0 Then
                        If .GetNode(i).Expanded Then
                            .RowHidden(r) = hidden
                        Else
                            .RowHidden(r) = True
                        End If
                    Else
                        .RowHidden(r) = hidden
                    End If
                    
                End If
            End If
        Next
    End With
End Sub
Public Sub ColorizeItems(ByVal r As Long)
    Const RateZeroColor = 16646111 'light cyan &HC0C0FF 'pink
    Dim s           As String
    Dim rowBudgeted As Boolean
    Dim rowPOed     As Boolean
    Dim rowLocked   As Boolean
    Dim startRow    As Long
    Dim EndRow      As Long
        
    Dim CanIssuePOs As Boolean
        
    With gItems
        .Redraw = flexRDNone
        CanIssuePOs = HFApp.UserPermission("IssuePOs")
        
        If r > 0 Then
            startRow = r
            EndRow = r
        Else
            startRow = 1
            EndRow = .Rows - 1
        End If
        For r = startRow To EndRow
        If .IsSubtotal(r) Then
'            .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbHighlight
'            .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbWindowBackground
        Else
            s = ""
            rowBudgeted = .ValueMatrix(r, .ColIndex("BudgetGenerated")) <> 0
            rowPOed = Trim(.TextMatrix(r, .ColIndex("PONumber"))) <> ""
            rowLocked = (rowBudgeted And mMode <> fcePO) Or (rowPOed And mMode = fcePO) Or (CanIssuePOs = False And mMode = fcePO)
            
            'clear all
            .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbWindowText
            .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbWindowBackground
            .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = False
            .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = False
            .Cell(flexcpFontStrikethru, r, 0, r, .Cols - 1) = False
            
            
            ' cancelled and reversals
            If IsIn(.TextMatrix(r, .ColIndex("RowType")), "Original - Cancelled", "Original - Reversal") Then
                .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbGrayText
                .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbWindowBackground
                .Cell(flexcpFontStrikethru, r, 0, r, .Cols - 1) = True
                .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = False
                .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = True
            End If
                

            
            If rowLocked Then
                .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbGrayText
            Else
                                
                ' less than zero
                If .ValueMatrix(r, .ColIndex(IIf(mMode <> fcePO, "BudgetRate", "PORate"))) < 0 Or .ValueMatrix(r, .ColIndex(IIf(mMode <> fcePO, "BudgetQty", "POQty"))) < 0 Then
                    .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_QtyRateLTZero_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_QtyRateLTZero_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_QtyRateLTZero_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_QtyRateLTZero_FontStyle), "Bold")
                End If
                
                ' equal to zero
                If .ValueMatrix(r, .ColIndex(IIf(mMode <> fcePO, "BudgetPretax", "POPretax"))) = 0 Then
                    .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_QtyRateEQZero_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_QtyRateEQZero_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_QtyRateEQZero_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_QtyRateEQZero_FontStyle), "Bold")
                End If
                
                
                ' if overridden value
                If .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True" Then
                    .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_NonSysRate_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_NonSysRate_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Bold")
                End If
                If .TextMatrix(r, .ColIndex("POOverridden")) = "True" Then
                    .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_NonSysRate_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_NonSysRate_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Bold")
                End If
                
                
                
                
                ' figure out if there are errors
                If mMode <> fcePO Then
                    If Trim(.TextMatrix(r, .ColIndex("BudgetVendor"))) = "" Then s = s & vbCrLf & "Vendor not specified"
                    If HFApp.Options(TaxGroupRequired) Then If Trim(.TextMatrix(r, .ColIndex("BudgetTaxGroup"))) = "" Then s = s & vbCrLf & "Tax Group not specified"
                Else
                    If Trim(.TextMatrix(r, .ColIndex("POVendor"))) = "" Then s = s & vbCrLf & "Vendor not specified"
                    If HFApp.Options(TaxGroupRequired) Then If Trim(.TextMatrix(r, .ColIndex("POTaxGroup"))) = "" Then s = s & vbCrLf & "Tax Group not specified"
                End If
                    
                If Trim(.TextMatrix(r, .ColIndex("POIndex"))) = "" Then s = s & vbCrLf & "PO Index not specified"
                If Trim(.TextMatrix(r, .ColIndex("JCCostCode"))) = "" Then s = s & vbCrLf & "Cost Code not specified"
                If Trim(.TextMatrix(r, .ColIndex("JCCategory"))) = "" Then s = s & vbCrLf & "Category not specified"
                
                gItems.Cell(flexcpData, r, 0, r, .Cols - 1) = Mid(s, 3)
                If s = "" Then
                    .Cell(flexcpPicture, r, .ColIndex("WarningMessages")) = Nothing
                Else
                    .Cell(flexcpPicture, r, .ColIndex("WarningMessages")) = imgWarning.Picture
                    .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_InvalidData_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_InvalidData_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_InvalidData_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_InvalidData_FontStyle), "Bold")
                End If
                
                
                
            End If
        End If
        Next
        
        
        
        .Redraw = flexRDBuffered
    End With


End Sub

Public Function SaveItems() As Boolean
On Error GoTo eh
    Dim w As Long
    Dim i As Long
    Dim s As String
    Dim r As Long
    Dim AssmID As Integer
    Dim sJobNumber As String
    
    Dim IsReversingItem As Boolean
    Dim IsCorrectingItem As Boolean
    
    
    If frmChangeOrder.Visible Then
        s = ""
        s = s & "UPDATE changeordermaster" & vbCrLf
        s = s & "   SET comments=" & DbQuote(Str, txtCOComments.Text) & vbCrLf
        s = s & "      ,bldr_approved_by=" & DbQuote(Str, txtCOApprovedBy.Text) & vbCrLf
        s = s & "      ,bldr_approved_date=" & DbQuote(Date, txtCODateApproved.Text) & vbCrLf
        s = s & "      ,change_date=" & DbQuote(Date, txtCODate.Text) & vbCrLf
        s = s & "WHERE Customer_no=" & DbQuote(Str, mCustomerNo) & vbCrLf
        s = s & "  AND Change_Order_No=" & DbQuote(Str, mChangeOrder) & vbCrLf
        Call HFApp.SqlExec(s)
    End If
    
    
    
    If frmPOIndex.Visible Then
        If Trim(txtPONumber.Text) <> "" Then
            'save to pomaster
            s = ""
            s = s & "UPDATE POMaster" & vbCrLf
            s = s & "   SET Description=" & DbQuote(Str, txtPODescription.Text) & vbCrLf
            s = s & "      ,StandardText=" & DbQuote(Str, txtStandardText.Text) & vbCrLf
            s = s & "      ,ShipVia=" & DbQuote(Str, txtShipVia.Text) & vbCrLf
            s = s & "      ,FOB=" & DbQuote(Str, txtFOB.Text) & vbCrLf
            s = s & "      ,Terms=" & DbQuote(Str, txtTerms.Text) & vbCrLf
            s = s & "      ,OrderedBy=" & DbQuote(Str, txtOrderedBy.Text) & vbCrLf
            s = s & "      ,HideQty=" & DbQuote(Bit, chkHideQty.value = vbChecked) & vbCrLf
            s = s & "      ,HidePrice=" & DbQuote(Bit, chkHidePrice.value = vbChecked) & vbCrLf
            s = s & "      ,TotalOnly=" & DbQuote(Bit, chkTotalOnly.value = vbChecked) & vbCrLf
            s = s & "      ,IncludeDocuments=" & DbQuote(Bit, chkIncludeDocuments.value = vbChecked) & vbCrLf
            s = s & "      ,PODate=" & DbQuote(Date, txtDateIssued.Text) & vbCrLf
            s = s & "      ,DeliveryRecipient=" & DbQuote(Str, txtDeliveryRecipient.Text) & vbCrLf
            s = s & "      ,DeliveryAddress=" & DbQuote(Str, txtDeliveryAddress.Text) & vbCrLf
            s = s & "      ,RetainagePercent=" & DbQuote(Num, Val(Replace(txtRetainagePercent.Text, "%", ""))) & vbCrLf
            Select Case True
                Case InStr(1, txtDeliveryAddress.Text, "@") > 0:  s = s & "      ,DeliveryMethod=" & DbQuote(Num, dtEmail) & vbCrLf
                Case Trim(txtDeliveryAddress.Text) = "":          s = s & "      ,DeliveryMethod=" & DbQuote(Num, dtPrint) & vbCrLf
                Case Else:                                        s = s & "      ,DeliveryMethod=" & DbQuote(Num, dtFax) & vbCrLf
            End Select
            s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, txtPONumber) & vbCrLf
            
        Else
            'save to jobpoindex
            On Error Resume Next
            Call HFApp.SqlExec("INSERT INTO JobPOIndex(DivisionID,POIndex,Job_No) VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, txtPOIndex) & "," & DbQuote(Str, mJob) & ")")
            On Error GoTo eh
            
            s = ""
            s = s & "UPDATE JobPOIndex" & vbCrLf
            s = s & "   SET Description=" & DbQuote(Str, txtPODescription.Text) & vbCrLf
            s = s & "      ,StandardText=" & DbQuote(Str, txtStandardText.Text) & vbCrLf
            s = s & "      ,ShipVia=" & DbQuote(Str, txtShipVia.Text) & vbCrLf
            s = s & "      ,FOB=" & DbQuote(Str, txtFOB.Text) & vbCrLf
            s = s & "      ,Terms=" & DbQuote(Str, txtTerms.Text) & vbCrLf
            s = s & "      ,OrderedBy=" & DbQuote(Str, txtOrderedBy.Text) & vbCrLf
            s = s & "      ,HideQty=" & DbQuote(Bit, chkHideQty.value = vbChecked) & vbCrLf
            s = s & "      ,HidePrice=" & DbQuote(Bit, chkHidePrice.value = vbChecked) & vbCrLf
            s = s & "      ,TotalOnly=" & DbQuote(Bit, chkTotalOnly.value = vbChecked) & vbCrLf
            s = s & "      ,RetainagePercent=" & DbQuote(Num, txtRetainagePercent.Text) & vbCrLf
            s = s & "WHERE POIndex=" & DbQuote(Str, txtPOIndex) & vbCrLf
            s = s & "   AND Job_No=" & DbQuote(Str, mJob) & vbCrLf
            s = s & "   and DivisionID = " & HFApp.DivisionID
        End If
        Call HFApp.SqlExec(s)
    End If
    
    
    If mAssemblyChanged Then
        'change EstimateAssembly also
        s = ""
        s = s & "UPDATE EstimateAssemblies" & vbCrLf
        s = s & "SET JCExtra=" & DbQuote(Str, txtJCExtra.Text) & vbCrLf
        s = s & "   ,HFDescription=" & DbQuote(Str, txtHFDescription.Text) & vbCrLf
        s = s & "   ,SalesQty=" & DbQuote(Num, txtQuantity.Text) & vbCrLf
        s = s & "   ,HFComments=" & DbQuote(Str, txtHFComments.Text) & vbCrLf
        s = s & "   ,EstimatorNotes=" & DbQuote(Str, txtEstimatorNotes.Text) & vbCrLf
        s = s & "   ,color=" & DbQuote(Str, txtColor.Text) & vbCrLf
        s = s & "   ,location=" & DbQuote(Str, txtLocation.Text) & vbCrLf
        s = s & "   ,Style=" & DbQuote(Str, txtStyle.Text) & vbCrLf
        s = s & "   ,Other=" & DbQuote(Str, txtOther.Text) & vbCrLf
        s = s & "   ,Finish=" & DbQuote(Str, txtFinish.Text) & vbCrLf
        s = s & "   ,SalesRate =" & DbQuote(Num, txtPrice.Text) & vbCrLf
        s = s & "   ,ChangeRequestedBy=" & DbQuote(Str, txtChangeRequestedBy.Text) & vbCrLf
        s = s & "   ,ChangeRequestedDate=" & DbQuote(Date, txtChangeRequestedDate.Text) & vbCrLf
        s = s & "   ,ChangeApprovedBy=" & DbQuote(Str, txtChangeApprovedBy.Text) & vbCrLf
        s = s & "   ,ChangeApprovedDate=" & DbQuote(Date, txtChangeApprovedDate.Text) & vbCrLf
        s = s & "   ,ChangeRequestStatus=" & DbQuote(Str, cboChangeRequestStatus.Text) & vbCrLf
        
        s = s & "WHERE EstAssemblyID=" & DbQuote(Num, mEstAssemblyID) & vbCrLf
        Call HFApp.SqlExec(s)
    End If
    
    With gItems
        For i = .Rows - 1 To 1 Step -1
            If .RowData(i) = "NEW" Then
                If "" & .TextMatrix(i, .ColIndex("Job_No")) = "" Then
                     If Val(.TextMatrix(i, .ColIndex("EstAssemblyID"))) <> AssmID Then
                          AssmID = Val(.TextMatrix(i, .ColIndex("EstAssemblyID")))
                          sJobNumber = "" & HFApp.SqlExec("select Job from estimateassemblies where EstAssemblyid = " & DbQuote(Num, .TextMatrix(i, .ColIndex("EstAssemblyID"))))(0)
                     End If
                Else
                     sJobNumber = .TextMatrix(i, .ColIndex("Job_No"))
                End If
                
                
                s = ""
                s = s & "INSERT INTO EstimateItems(EstAssemblyID,DivisionID,Assembly,AssemblyDescription,POIndex,Model,Sequence,Phase,Item,Job,JCExtra,JCCostCode,JCCategory,VarianceJCCategory,SortOrder,Description,Comments,formula,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,BudgetDeleted,PODeleted,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate,BudgetGenerated,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,POOverridden,BudgetOverridden,ExcludeFromPO,OriginalJCCategory,SalesQty,Location,Unit,IsReversingItem,IsCorrectingItem,ReversingItemID,CorrectingItemID"
                For w = 1 To 40
                    s = s & ",WBS" & format(w, "00")
                Next
                s = s & ")" & vbCrLf
                s = s & "VALUES(" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstAssemblyID"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("AssemblyDescription"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("POIndex"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Model"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Sequence"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstPhase"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstItem"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, sJobNumber) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCExtra"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("VarianceJCCategory"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, i) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemDesc"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemComments"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Formula"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TakeoffUOM"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("ConversionFactor"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM"))) & vbCrLf
                s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("BudgetDeleted"))) & vbCrLf
                s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("PODeleted"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetVendor"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetQty"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetPretax"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetTaxGroup"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTax"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTax"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetGenerated"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("POVendor"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("POQty"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("PORate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("POPretax"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("POTaxGroup"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("POJCTax"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("POJCTaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("PONJCTax"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("PONJCTaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("POOverridden")) = "True") & vbCrLf
                s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("BudgetOverridden")) = "True") & vbCrLf
                s = s & "      ," & IIf(.Cell(flexcpChecked, i, .ColIndex("ExcludeFromPO")) = flexChecked, 1, 0) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("OriginalJCCategory"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("SalesQty"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Location"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Unit"))) & vbCrLf
                s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("IsReversingItem"))) & vbCrLf
                s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("IsCorrectingItem"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("ReversingItemID")), True) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("CorrectingItemID")), True) & vbCrLf
                IsReversingItem = .TextMatrix(i, .ColIndex("IsReversingItem")) = "true"
                IsCorrectingItem = .TextMatrix(i, .ColIndex("IsCorrectingItem")) = "true"
                For w = 1 To 40
                    s = s & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("WBS" & format(w, "00")))) & vbCrLf
                Next
                s = s & ")" & vbCrLf
                HFApp.SqlExec s
                .TextMatrix(i, .ColIndex("EstItemID")) = HFApp.SqlIdentity("EstimateItems")
                
                If IsReversingItem Then
                    'find row where reversingitemid=this row num, set reversingitemid on that row to id just created
                    r = .FindRow(-1 * i, , .ColIndex("ReversingItemID"))
                    If r > 0 Then
                        .TextMatrix(r, .ColIndex("ReversingItemID")) = .TextMatrix(i, .ColIndex("EstItemID"))
                        If .RowData(r) <> "NEW" Then .RowData(r) = "DIRTY"
                    End If
                End If
                If IsCorrectingItem Then
                    'find row where correctingitemid=this row num, set correctingitemid on that row to id just created
                    r = .FindRow(-1 * i, , .ColIndex("CorrectingItemID"))
                    If r > 0 Then
                        .TextMatrix(r, .ColIndex("CorrectingItemID")) = .TextMatrix(i, .ColIndex("EstItemID"))
                        If .RowData(r) <> "NEW" Then .RowData(r) = "DIRTY"
                    End If
                
                    'repeat the process to assign correcting item id to the reversing row
                    r = .FindRow(-1 * i, , .ColIndex("CorrectingItemID"))
                    If r > 0 Then
                        .TextMatrix(r, .ColIndex("CorrectingItemID")) = .TextMatrix(i, .ColIndex("EstItemID"))
                        If .RowData(r) <> "NEW" Then .RowData(r) = "DIRTY"
                    End If
                End If
                    
                
                
                .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexUnchecked
                .RowData(i) = ""
            End If
            
            If .RowData(i) = "DIRTY" Then
            
                If "" & .TextMatrix(i, .ColIndex("Job_No")) = "" Then
                     If Val(.TextMatrix(i, .ColIndex("EstAssemblyID"))) <> AssmID Then
                          AssmID = Val(.TextMatrix(i, .ColIndex("EstAssemblyID")))
                          sJobNumber = "" & HFApp.SqlExec("select Job from estimateassemblies where EstAssemblyid = " & DbQuote(Num, .TextMatrix(i, .ColIndex("EstAssemblyID"))))(0)
                          .TextMatrix(i, .ColIndex("Job_No")) = sJobNumber
                     End If
                End If
                
                s = ""
                s = s & "UPDATE EstimateItems" & vbCrLf
                s = s & "SET Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("EstPhase"))) & vbCrLf
                s = s & "   ,Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("EstItem"))) & vbCrLf
                s = s & "   ,TakeoffQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty"))) & vbCrLf
                s = s & "   ,BudgetQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetQty"))) & vbCrLf
                s = s & "   ,POQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("POQty"))) & vbCrLf
                s = s & "   ,Job=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Job_No"))) & vbCrLf
                s = s & "   ,EstAssemblyID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstAssemblyID"))) & vbCrLf
                s = s & "   ,Assembly=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf
                s = s & "   ,AssemblyDescription=" & DbQuote(Str, .TextMatrix(i, .ColIndex("AssemblyDescription"))) & vbCrLf
                s = s & "   ,POIndex=" & DbQuote(Str, .TextMatrix(i, .ColIndex("POIndex"))) & vbCrLf
                s = s & "   ,JCExtra=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCExtra")), , True) & vbCrLf
                s = s & "   ,JCCostCode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
                s = s & "   ,JCCategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
                s = s & "   ,VarianceJCCategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("VarianceJCCategory"))) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemDesc"))) & vbCrLf
                s = s & "   ,Comments=" & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemComments"))) & vbCrLf
                s = s & "   ,OrderUOM=" & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM"))) & vbCrLf
                s = s & "   ,Formula=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Formula"))) & vbCrLf
                s = s & "   ,ConversionFactor=" & DbQuote(Num, .TextMatrix(i, .ColIndex("ConversionFactor"))) & vbCrLf
                s = s & "   ,BudgetDeleted=" & DbQuote(Bit, .TextMatrix(i, .ColIndex("BudgetDeleted"))) & vbCrLf
                s = s & "   ,PODeleted=" & DbQuote(Bit, .TextMatrix(i, .ColIndex("PODeleted"))) & vbCrLf
                s = s & "   ,BudgetVendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetVendor"))) & vbCrLf
                s = s & "   ,BudgetRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetRate"))) & vbCrLf
                s = s & "   ,BudgetPretax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetPretax"))) & vbCrLf
                s = s & "   ,BudgetTaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetTaxGroup"))) & vbCrLf
                s = s & "   ,BudgetJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTax"))) & vbCrLf
                s = s & "   ,BudgetJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTaxRate"))) & vbCrLf
                s = s & "   ,BudgetNJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTax"))) & vbCrLf
                s = s & "   ,BudgetNJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTaxRate"))) & vbCrLf
                s = s & "   ,POVendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("POVendor"))) & vbCrLf
                s = s & "   ,PORate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("PORate"))) & vbCrLf
                s = s & "   ,POPretax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("POPretax"))) & vbCrLf
                s = s & "   ,POTaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("POTaxGroup"))) & vbCrLf
                s = s & "   ,POJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("POJCTax"))) & vbCrLf
                s = s & "   ,POJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("POJCTaxRate"))) & vbCrLf
                s = s & "   ,PONJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("PONJCTax"))) & vbCrLf
                s = s & "   ,PONJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("PONJCTaxRate"))) & vbCrLf
                s = s & "   ,ExcludeFromPO=" & IIf(.Cell(flexcpChecked, i, .ColIndex("ExcludeFromPO")) = flexChecked, 1, 0) & vbCrLf
                s = s & "   ,BudgetOverridden=" & DbQuote(Bit, .TextMatrix(i, .ColIndex("BudgetOverridden")) = "True") & vbCrLf
                s = s & "   ,POOverridden=" & DbQuote(Bit, .TextMatrix(i, .ColIndex("POOverridden")) = "True") & vbCrLf
                s = s & "   ,SalesQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("SalesQty"))) & vbCrLf
                s = s & "   ,Location=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Location"))) & vbCrLf
                s = s & "   ,Unit=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Unit"))) & vbCrLf
                s = s & "   ,ReversingItemID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("ReversingItemID"))) & vbCrLf
                s = s & "   ,CorrectingItemID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("CorrectingItemID"))) & vbCrLf
                For w = 1 To 40
                    s = s & "   ,WBS" & format(w, "00") & "=" & DbQuote(Str, .TextMatrix(i, .ColIndex("WBS" & format(w, "00")))) & vbCrLf
                Next
                s = s & "WHERE EstItemID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstItemID"))) & vbCrLf
                Call HFApp.SqlExec(s)
                .RowData(i) = ""
            End If
            
            
            
            If ((IsIn(mMode, fceQuote, fceBudget) And .TextMatrix(i, .ColIndex("BudgetDeleted")) = "True") Or ((Not IsIn(mMode, fceQuote, fceBudget)) And .TextMatrix(i, .ColIndex("PODeleted")) = "True")) Then
                'remove row if its been deleted
                Call .RemoveItem(i)
            End If
            
            
        Next
        
        
        'remove records that are both budgetdeleted and podeleted
        'if you delete a reversing item or a correcting item, delete it's pair and remove the references from the original item leaving it in an uncorrected state
        'also ensure that the category of reversal matches that of the correcting entry
        s = ""
        s = s & "--delete reversing entry" & vbCrLf
        s = s & "delete r" & vbCrLf
        s = s & "from estimateitems i" & vbCrLf
        s = s & "join estimateitems o on (o.reversingitemid=i.estitemid or o.correctingitemid=i.estitemid) and isnull(o.isreversingitem,0)=0 " & vbCrLf
        s = s & "join estimateitems r on r.estitemid=o.reversingitemid" & vbCrLf
        s = s & "where i.budgetdeleted=1" & vbCrLf
        s = s & "and i.podeleted=1" & vbCrLf
        s = s & "and i.divisionid = " & HFApp.DivisionID & vbCrLf
        s = s & "and i.job=" & DbQuote(Str, mJob) & vbCrLf
        s = s & "" & vbCrLf
        
        s = s & "--delete correcting entry" & vbCrLf
        s = s & "delete c" & vbCrLf
        s = s & "from estimateitems i" & vbCrLf
        s = s & "join estimateitems o on (o.reversingitemid=i.estitemid or o.correctingitemid=i.estitemid) and isnull(o.isreversingitem,0)=0 " & vbCrLf
        s = s & "join estimateitems c on c.estitemid=o.correctingitemid" & vbCrLf
        s = s & "where i.budgetdeleted=1" & vbCrLf
        s = s & "and i.podeleted=1" & vbCrLf
        s = s & "and i.divisionid = " & HFApp.DivisionID & vbCrLf
        s = s & "and i.job=" & DbQuote(Str, mJob) & vbCrLf
        s = s & "" & vbCrLf
        
        s = s & "--remove references from original if the reversing item is gone" & vbCrLf
        s = s & "update i set reversingitemid=null, correctingitemid=null" & vbCrLf
        s = s & "from estimateitems i" & vbCrLf
        s = s & "left outer join estimateitems x on (i.reversingitemid=x.estitemid)" & vbCrLf
        s = s & "where i.divisionid = " & HFApp.DivisionID & vbCrLf
        s = s & "and i.job=" & DbQuote(Str, mJob) & vbCrLf
        s = s & "and i.reversingitemid is not null " & vbCrLf
        s = s & "and x.estitemid is null" & vbCrLf
        s = s & "" & vbCrLf
        
        s = s & "--remove the thing you removed" & vbCrLf
        s = s & "delete estimateitems " & vbCrLf
        s = s & "where budgetdeleted=1" & vbCrLf
        s = s & "and podeleted=1" & vbCrLf
        s = s & "and divisionid = " & HFApp.DivisionID & vbCrLf
        s = s & "and job=" & DbQuote(Str, mJob) & vbCrLf
        s = s & "" & vbCrLf
        
        s = s & "--update category of reversals" & vbCrLf
        s = s & "update r set jccategory=c.jccategory" & vbCrLf
        s = s & "from estimateitems c" & vbCrLf
        s = s & "left join estimateitems r on c.estitemid=r.correctingitemid and isnull(r.isreversingitem,0)=1" & vbCrLf
        s = s & "where c.iscorrectingitem=1 " & vbCrLf
        s = s & "and c.divisionid = " & HFApp.DivisionID & vbCrLf
        s = s & "and c.job=" & DbQuote(Str, mJob) & vbCrLf

        Call HFApp.SqlExec(s)

    End With
    SaveItems = True
    Call ReloadAssemblyTotals(-1)
Exit Function
eh: Call errHandler(SRCFILE & "SaveItems", s)
End Function

Private Sub txtHFComments_Change()
    Dirty = True
    mAssemblyChanged = True
End Sub

Private Sub txtHFDescription_Change()
    Dirty = True
    mAssemblyChanged = True
End Sub

Private Sub txtHFOption_GotFocus()
    SelectAll txtHFOption
End Sub

Private Sub txtJCExtra_Change()
    Dirty = True
    mAssemblyChanged = True
End Sub

Private Sub txtJCExtra_GotFocus()
    SelectAll txtJCExtra
End Sub

Private Sub txtJob_Change()
        
    Dirty = True
End Sub

Private Sub txtJob_GotFocus()
    SelectAll txtJob
End Sub

Private Sub txtLot_Change()
    Dirty = True
End Sub

Private Sub txtLot_GotFocus()
    SelectAll txtLot
End Sub

Private Sub txtLotPlan_Change()
    Dirty = True
End Sub

Private Sub txtLotPlan_GotFocus()
    SelectAll txtLotPlan
End Sub

Private Sub txtNotes_Change()
    Dirty = True
End Sub

Private Sub txtNotes_GotFocus()
    SelectAll txtNotes
End Sub

Private Sub txtOrderedBy_GotFocus()
    SelectAll txtOrderedBy
End Sub

Private Sub txtPODescription_Change()
    Dirty = True
End Sub

Private Sub txtPODescription_GotFocus()
    SelectAll txtPODescription
End Sub


Private Sub txtPOIndex_GotFocus()
    SelectAll txtPOIndex
End Sub

Private Sub txtPONumber_GotFocus()
    SelectAll txtPONumber
End Sub

Private Sub txtPONumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(5)
End Sub

Private Sub txtPrice_Change()
    Dirty = True
    mAssemblyChanged = True
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
txtPrice.Text = format("" & txtPrice.Text, "#,##0.00")
End Sub

Private Sub txtQuantity_Change()
    Dirty = True
    mAssemblyChanged = True
End Sub

Private Sub txtQuantity_GotFocus()
    SelectAll txtQuantity
End Sub

Private Sub txtRetainagePercent_Change()
    Dirty = True
End Sub
Private Sub txtRetainagePercent_Validate(Cancel As Boolean)
    txtRetainagePercent.Text = Min(Max(Val(txtRetainagePercent.Text), 0), 100) & "%"
End Sub
Private Sub txtRetainagePercent_GotFocus()
    SelectAll txtRetainagePercent
End Sub

Private Sub txtRFPComments_Change()
    Dirty = True
End Sub

Private Sub txtRFPComments_GotFocus()
    SelectAll txtRFPComments
End Sub

Private Sub txtRFPDescription_Change()
    Dirty = True
End Sub

Private Sub txtRFPDescription_GotFocus()
    SelectAll txtRFPDescription
End Sub

Private Sub txtShipVia_GotFocus()
    SelectAll txtShipVia
End Sub

Private Sub txtStandardText_Change()
    Dirty = True
End Sub
Private Sub txtShipVia_Change()
    Dirty = True
End Sub
Private Sub txtFOB_Change()
    Dirty = True
End Sub

Private Sub txtStandardText_GotFocus()
    SelectAll txtStandardText
End Sub


Private Sub txtStandardText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(4)
End Sub

Private Sub txtStyle_Change()
    Dirty = True
End Sub
Private Sub txtOther_Change()
    Dirty = True
End Sub
Private Sub txtStyle_GotFocus()
    SelectAll txtStyle
End Sub
Private Sub txtOther_GotFocus()
    SelectAll txtOther
End Sub

Private Sub txtTerms_Change()
    Dirty = True
End Sub
Private Sub txtOrderedBy_Change()
    Dirty = True
End Sub
Private Sub txtDeliveryRecipient_Change()
    Dirty = True
End Sub
Private Sub txtDeliveryAddress_Change()
    Dirty = True

    Select Case True
        Case InStr(1, txtDeliveryAddress.Text, "@") > 0:  chkIncludeDocuments.Visible = True
        Case Trim(txtDeliveryAddress.Text) = "":          chkIncludeDocuments.Visible = False
        Case Else:                                        chkIncludeDocuments.Visible = False
    End Select

End Sub

Private Function PrettyName(fieldname As String) As String
    Dim s As String
    s = fieldname
    Select Case s
        Case "JCCostCode+' - '+JCCostCodeDesc": s = "Cost Code"
        Case "POIndex + ' ' + POIndexDescription": s = "PO Index"
        Case "case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end": s = "PO Index"
        Case "CASE ChangeOrder WHEN '' THEN 'Initial Contract' ELSE ChangeOrder END": s = "Change Order"
        Case "CASE ChangeOrder WHEN '' THEN 'Initial Contract' ELSE ChangeOrder END ChangeOrder": s = "Change Order"
        Case "PONumber + ' - ' + POVendorName":                                                                                                s = "PO Number"
        Case "PONumber + ' - ' + POIndex":                                                                                                     s = "PO Number"
        Case "PONumber + CASE PONumber WHEN '' THEN '' ELSE ' -- ' + POVendorName END":                                                        s = "PO Number"
        Case "PONUmber + ' - ' + case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end":               s = "PO Number"
        Case "POGroup":           s = "PO Group"
        Case "POIndex":           s = "PO Index"
        Case "BudgetVendorName":  s = "Vendor"
        Case "POVendorName":      s = "Vendor"
        Case "CustomerDesc":      s = "Customer"
        Case "ChangeOrder":       s = "Change Order"
        Case "HFDescription":     s = "Assembly"
        Case "HFCategory":        s = "Option Category"
        Case "JCExtra":           s = "Extra"
        Case "JCCostCodeDesc":    s = "Cost Code"
        Case "JCCategoryDesc":    s = "Category"
        Case "EstimateIndex":     s = "Estimate Number"
        Case "AssemblyTypeDesc":  s = "Assembly Type"
        Case "POGroup":           s = "PO Group"
        Case "PONumber":          s = "PO Number"
        Case "JobDesc":           s = "Job"
        Case "RFPTitle":          s = "RFP"
        Case "CommunityDesc":     s = "Project"
        Case "Job_No":            s = "Job"
        Case "AssemblyJob_No":        s = "Job"
        Case "Job_no+' '+CustomerDesc": s = "Job"
        Case "AssemblyJob_No+' '+JobDesc": s = "Job"
        Case "CORNumber":         s = "Change Order"
        Case "CommunityPhase":     s = "Phase"
        Case "":                  s = ""
        Case Else
            MsgBox "STOP!!!" & vbCrLf & vbCrLf & "YOU NEED TO TRANSLATE THIS", vbExclamation
            Stop
    End Select
    PrettyName = s
End Function


Private Function GetVendorCost(Row As Long, Vendor As String, rate As Double) As Boolean
    Dim s As String
    Dim r As Double
    With gItems
    
        s = ""
        s = s & "SELECT dbo.Purch_GetItemRate(" & vbCrLf
        s = s & "       0" & vbCrLf
        s = s & "      ,0" & vbCrLf
        s = s & "      ," & DbQuote(Str, mCommunity) & vbCrLf
        s = s & "      ," & DbQuote(Str, mCommunityPhase) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Model"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("OptionID"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & vbCrLf
        s = s & "      ," & DbQuote(Num, .TextMatrix(Row, .ColIndex("Sequence"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, Vendor) & vbCrLf
        s = s & "      ,GETDATE()," & HFApp.DivisionID & ")"
        On Error Resume Next
        r = HFApp.SqlExec(s)(0)
        
        If r <> 0 Or HFApp.Options(ZeroRateOnChangeVendor) Then
            GetVendorCost = True
            rate = r
        End If
    End With
End Function

Private Sub txtTerms_GotFocus()
    SelectAll txtTerms
End Sub

Public Sub mnuEstimateItemsCustomerSub_Click(Index As Integer)
End Sub

Public Sub mnuEstimateItemsGridSub_Click(Index As Integer)
    Dim rs As Recordset
    Dim rc As Double
    Dim rc2 As Long
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim NewRow As Long
    Dim RevRow As Long
    
    Dim Phase As String
    Dim Item As String
    Dim Description As String
    
    With gItems
        If .Row <= 0 Then Exit Sub
        
        Select Case Index
        
            Case mcITEM_FORMATTING
                Call FEstimateItemsFormatting.ShowPurchasingRules
                Call ColorizeItems(0)
                
            Case mcITEM_VIEWFILES
                Call SaveData(False)
                
                Timer1.Enabled = True
                Timer1.Interval = 10
                
                s = ""
                s = s & "EditAttachments|"
                s = s & "JEI~" & .TextMatrix(.Row, .ColIndex("EstItemID")) & "|Job|"
                
                If .TextMatrix(.Row, .ColIndex("QuoteAttachmentID")) <> "" Then
                    s = s & .TextMatrix(.Row, .ColIndex("QuoteAttachmentID")) & "|Quote|"
                End If
                
                s = s & "ASM~" & mCommunity & "~" & .TextMatrix(.Row, .ColIndex("Model")) & "~" & .TextMatrix(.Row, .ColIndex("optionid")) & "~" & .TextMatrix(.Row, .ColIndex("assembly")) & "~" & .TextMatrix(.Row, .ColIndex("EstPhase")) & "~" & .TextMatrix(.Row, .ColIndex("EstItem")) & "|Model Option Library|"
                s = s & "ITM~" & .TextMatrix(.Row, .ColIndex("EstPhase")) & "~" & .TextMatrix(.Row, .ColIndex("EstItem")) & "|Item Database"
                mTimerTask = s
                
            Case mcITEM_REMOVEITEMS
                Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)
        
            Case mcITEM_MODIFYQTY
                If Not PostPOQtyToAccounting Then Exit Sub
                
                Dirty = True
                
                RevRow = .Row + 1
                NewRow = .Row + 2
                
                
                'ADD REVERSING ENTRY
                .AddItem "", RevRow
                .RowData(RevRow) = "NEW"
                For c = 0 To .Cols - 1
                    .TextMatrix(RevRow, c) = .TextMatrix(.Row, c)
                Next
                .TextMatrix(RevRow, .ColIndex("RowType")) = "Original - Reversal"
                Call ColorizeItems(RevRow)
                .TextMatrix(RevRow, .ColIndex("ReversingItemID")) = ""
                .TextMatrix(RevRow, .ColIndex("CorrectingItemID")) = ""
                .TextMatrix(RevRow, .ColIndex("IsReversingItem")) = True
                .TextMatrix(RevRow, .ColIndex("POQty")) = -1 * .TextMatrix(.Row, .ColIndex("POQty"))
                .TextMatrix(RevRow, .ColIndex("POPretax")) = -1 * .TextMatrix(.Row, .ColIndex("POPretax"))
                .TextMatrix(RevRow, .ColIndex("POJCTax")) = -1 * .TextMatrix(.Row, .ColIndex("POJCTax"))
                .TextMatrix(RevRow, .ColIndex("PONJCTax")) = -1 * .TextMatrix(.Row, .ColIndex("PONJCTax"))
                .TextMatrix(RevRow, .ColIndex("POTax")) = -1 * .TextMatrix(.Row, .ColIndex("POTax"))
                .TextMatrix(RevRow, .ColIndex("POTotal")) = -1 * .TextMatrix(.Row, .ColIndex("POTotal"))
                .TextMatrix(RevRow, .ColIndex("EstItemID")) = "0"
                .TextMatrix(RevRow, .ColIndex("BudgetQty")) = 0
                .TextMatrix(RevRow, .ColIndex("BudgetPretax")) = 0
                .TextMatrix(RevRow, .ColIndex("BudgetGenerated")) = "False"
                .TextMatrix(RevRow, .ColIndex("BudgetPostingBatch")) = "0"
                .TextMatrix(RevRow, .ColIndex("BudgetDeleted")) = True
                
                'ADD NEW ENTRY
                .AddItem "", NewRow
                .RowData(NewRow) = "NEW"
                For c = 0 To .Cols - 1
                    .TextMatrix(NewRow, c) = .TextMatrix(.Row, c)
                Next
                .TextMatrix(NewRow, .ColIndex("RowType")) = "Original - Correction"
                Call ColorizeItems(NewRow)
                
                .TextMatrix(NewRow, .ColIndex("ReversingItemID")) = ""
                .TextMatrix(NewRow, .ColIndex("CorrectingItemID")) = ""
                .TextMatrix(NewRow, .ColIndex("IsCorrectingItem")) = True
                .TextMatrix(NewRow, .ColIndex("EstItemID")) = "0"
                .TextMatrix(NewRow, .ColIndex("BudgetQty")) = 0
                .TextMatrix(NewRow, .ColIndex("BudgetPretax")) = 0
                .TextMatrix(NewRow, .ColIndex("BudgetGenerated")) = "False"
                .TextMatrix(NewRow, .ColIndex("BudgetPostingBatch")) = "0"
                .TextMatrix(NewRow, .ColIndex("BudgetDeleted")) = True
                                
                
                'for new/rev rows, temporarily put row num in the orginal row's id columns. When saving make sure to replace these with the
                'new/rev itemids as they are created. use negative numbers so they dont collide with actual itemids. chances of that are
                'minimal but this should make the chance zero.
                .TextMatrix(.Row, .ColIndex("RowType")) = "Original - Cancelled"
                Call ColorizeItems(.Row)
                .TextMatrix(.Row, .ColIndex("ReversingItemID")) = -1 * RevRow
                .TextMatrix(.Row, .ColIndex("CorrectingItemID")) = -1 * NewRow
                'rev row also gets reference to correcting row
                .TextMatrix(RevRow, .ColIndex("CorrectingItemID")) = -1 * NewRow
                        
                
                'choose new category. do not use varcats when posting qty to accounting
                s = "SELECT Category,Description FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, .TextMatrix(.Row, .ColIndex("JCCategory")), , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Category", "")) Then
                    .TextMatrix(NewRow, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .TextMatrix(NewRow, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                    .TextMatrix(RevRow, .ColIndex("JCCategory")) = .TextMatrix(NewRow, .ColIndex("JCCategory"))
                    .TextMatrix(RevRow, .ColIndex("JCCategoryDesc")) = .TextMatrix(NewRow, .ColIndex("JCCategoryDesc"))
                End If
                
                'you MUST do this right away or the row numbers will get screwed up by sorting or grouping
                Call SaveData(False)
                Call LoadEstItems
                
        
            Case mcITEM_SAVEONETIMETODB
                Description = .TextMatrix(.Row, .ColIndex("itemdesc"))
                If FItem.Add(Phase, Item, Description) Then
                    .TextMatrix(.Row, .ColIndex("estphase")) = Phase
                    .TextMatrix(.Row, .ColIndex("estitem")) = Item
                    .TextMatrix(.Row, .ColIndex("itemdesc")) = Description
                    If .RowData(.Row) <> "NEW" Then .RowData(.Row) = "DIRTY"
                    
                    s = ""
                    s = s & "update tblphaseitem" & vbCrLf
                    s = s & "set takeoffuom=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("takeoffuom"))) & vbCrLf
                    s = s & "   ,orderuom=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("orderuom"))) & vbCrLf
                    s = s & "   ,price=" & DbQuote(Num, .TextMatrix(.Row, .ColIndex(IIf(mMode <> fcePO, "budget", "po") & "rate"))) & vbCrLf
                    s = s & "   ,poindex=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("poindex"))) & vbCrLf
                    s = s & "   ,jccostcode=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("jccostcode"))) & vbCrLf
                    s = s & "   ,jccategory=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("jccategory"))) & vbCrLf
                    s = s & "   ,ustmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
                    s = s & "   ,tstmp=getdate()" & vbCrLf
                    s = s & "   ,taxgroup=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex(IIf(mMode <> fcePO, "budget", "po") & "taxgroup"))) & vbCrLf
                    s = s & "where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, Phase) & vbCrLf
                    s = s & "and item=" & DbQuote(Str, Item) & vbCrLf
                    
                    Call HFApp.SqlExec(s, dbHomefront)
                    
                    Call SaveData(False)
                    
                End If
            
            
            Case mcITEM_SUBSTITUEITEM
            
                s = ""
                s = s & "SELECT isnull(phase,'')+char(1)+isnull(item,'') phaseitem " & vbCrLf
                s = s & "      ,i.Phase" & vbCrLf
                s = s & "      ,i.Item" & vbCrLf
                s = s & "      ,i.Description" & vbCrLf
                s = s & "      ,i.TakeoffUOM" & vbCrLf
                s = s & "      ,i.ConversionFactor,i.RoundDir,i.RoundTo" & vbCrLf
                s = s & "      ,i.OrderUOM" & vbCrLf
                s = s & "      ,p.POIndex,i.PartNumber" & vbCrLf
                s = s & "      ,p.FullDescription POIndexDescription" & vbCrLf
                s = s & "      ,i.Notes" & vbCrLf
                s = s & "      ,cod.CostCode" & vbCrLf
                s = s & "      ,cod.description CostCodeDesc" & vbCrLf
                s = s & "      ,cat.Category" & vbCrLf
                s = s & "      ,cat.Description CategoryDesc,ConversionFactor,WastePercent" & vbCrLf
                s = s & "  FROM tblPhaseItem i" & vbCrLf
                s = s & "LEFT OUTER JOIN tblPOIndex p ON(i.DivisionID = p.DivisionID and i.poindex=p.poindex)" & vbCrLf
                s = s & "LEFT OUTER JOIN StandardCostCodes cod ON(i.DivisionID = cod.DivisionID and isnull(nullif(i.JCCostCode,''),p.JCCostCode)=cod.costcode)" & vbCrLf
                s = s & "LEFT OUTER JOIN StandardCategories cat ON(i.DivisionID = cat.DivisionID and isnull(nullif(i.JCCategory,''),p.JCCategory)=cat.category)" & vbCrLf
                s = s & "WHERE i.DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s, .TextMatrix(.Row, .ColIndex("EstPhase")) & Chr(1) & .TextMatrix(.Row, .ColIndex("EstItem")), , , , "phaseitem,conversionfactor,wastepercent,TakeoffUOM,EffectiveConversionFactor,OrderUOM,POIndex,Notes,CategoryDesc") Then
                    'this can only happen when neither pos or budgets have been generated so we only need to set budget values. the afteredit event will set the po values
                    
                    
                    s = ""
                    s = s & "select vendor_id,vendor_name" & vbCrLf
                    s = s & "from tblvendors" & vbCrLf
                    s = s & "where divisionid = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                    s = s & "and vendor_id=dbo.Purch_GetCommunityVendor(" & DbQuote(Str, mCommunity) & ", " & DbQuote(Str, FPickList.SelectedItem("POIndex")) & ", " & DbQuote(Num, HFApp.DivisionID) & ")" & vbCrLf
                    Set rs = HFApp.SqlExec(s)
                    
                    For r = .Row To .RowSel
                    
                        .TextMatrix(r, .ColIndex("EstPhase")) = FPickList.SelectedItem("Phase")
                        .TextMatrix(r, .ColIndex("EstItem")) = FPickList.SelectedItem("Item")
                        .TextMatrix(r, .ColIndex("ItemDesc")) = FPickList.SelectedItem("Description")
                        .TextMatrix(r, .ColIndex("ItemComments")) = FPickList.SelectedItem("Notes")
                        .TextMatrix(r, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                        .TextMatrix(r, .ColIndex("ConversionFactor")) = FPickList.SelectedItem("ConversionFactor")
                        .TextMatrix(r, .ColIndex("WastePercent")) = FPickList.SelectedItem("WastePercent")
                        .TextMatrix(r, .ColIndex("RoundDir")) = FPickList.SelectedItem("RoundDir")
                        .TextMatrix(r, .ColIndex("RoundTo")) = FPickList.SelectedItem("RoundTo")
                        
                        .TextMatrix(r, .ColIndex("TakeoffUOM")) = FPickList.SelectedItem("TakeoffUOM")
                        .TextMatrix(r, .ColIndex("OrderUOM")) = FPickList.SelectedItem("OrderUOM")
                        .TextMatrix(r, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                        .TextMatrix(r, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                        .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("CostCodeDesc")
                        .TextMatrix(r, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("CategoryDesc")
                        
                        If Not rs.EOF Then
                            .TextMatrix(r, .ColIndex("BudgetVendor")) = "" & rs("vendor_id")
                            .TextMatrix(r, .ColIndex("BudgetVendorName")) = "" & rs("vendor_name")
                            .TextMatrix(r, .ColIndex("POVendor")) = "" & rs("vendor_id")
                            .TextMatrix(r, .ColIndex("POVendorName")) = "" & rs("vendor_name")
                        End If
                        
                        
                        On Error Resume Next
                        rc = 0
                        rc = HFApp.SqlExec("select dbo.Purch_GetItemRate(0,0," & _
                                                          DbQuote(Str, mCurrentCommunity) & "," & _
                                                          DbQuote(Str, mCurrentCommunityPhase) & "," & _
                                                          DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & "," & _
                                                          DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & "," & _
                                                          DbQuote(Str, .TextMatrix(r, .ColIndex("OptionID"))) & "," & _
                                                          DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & "," & _
                                                          DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & "," & _
                                                          DbQuote(Num, .TextMatrix(r, .ColIndex("Sequence"))) & "," & _
                                                          DbQuote(Str, .TextMatrix(r, .ColIndex("BudgetVendor"))) & "," & _
                                                          "getdate()," & HFApp.DivisionID & ")", dbHomefront)(0)
                        
                        
                        
                        If rc <> 0 Or HFApp.Options(ZeroRateOnRefreshCosts) = "True" Then
                            .TextMatrix(r, .ColIndex("BudgetRate")) = rc
                        End If
                        
                        Call gItems_AfterEdit(r, .ColIndex("BudgetRate"))
                        .TextMatrix(r, .ColIndex("POOverridden")) = "False"
                        .TextMatrix(r, .ColIndex("BudgetOverridden")) = "False"
                        Call ColorizeItems(r)
                        
                        If .RowData(r) <> "NEW" Then .RowData(r) = "DIRTY"
                    Next
                    Dirty = True

                End If
                
            
            Case mcITEM_COPYITEMS
            
                
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    
                  
                    
                    Dirty = True
                    NewRow = Max(.Row, .RowSel) + 1
                    .AddItem "", NewRow
                    .RowData(NewRow) = "NEW"
                    For c = 0 To .Cols - 1
                        .TextMatrix(NewRow, c) = .TextMatrix(r, c)
                    Next
                    
                    If .TextMatrix(r, .ColIndex("BudgetsLocked")) = "True" And HFApp.Options.ValueByName("UseVarianceReporting") = "True" Then
                        If FPickList.Choose(HFApp.Databases(dbHomefront), "Variance Category", "Select [Category], [Description] from StandardCategories where isvariance=1 and DivisionID =" & HFApp.DivisionID, VarianceCatDef) Then
                            .TextMatrix(NewRow, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                            .TextMatrix(NewRow, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                        
                            .TextMatrix(NewRow, .ColIndex("VarianceJCCategory")) = FPickList.SelectedItem("Category")
                            .TextMatrix(NewRow, .ColIndex("VarianceJCCategoryDesc")) = FPickList.SelectedItem("Description")
                        
                        End If
                    End If
                    
                    .TextMatrix(NewRow, .ColIndex("ReversingItemID")) = ""
                    .TextMatrix(NewRow, .ColIndex("CorrectingItemID")) = ""
                    .TextMatrix(NewRow, .ColIndex("IsReversingItem")) = ""
                    .TextMatrix(NewRow, .ColIndex("IsCorrectingItem")) = ""
                    
                    .TextMatrix(NewRow, .ColIndex("BudgetQty")) = 0
                    .TextMatrix(NewRow, .ColIndex("POQty")) = 0
                    .TextMatrix(NewRow, .ColIndex("TakeoffQty")) = 0
                    .TextMatrix(NewRow, .ColIndex("POPretax")) = 0
                    .TextMatrix(NewRow, .ColIndex("BudgetPretax")) = 0
                    .TextMatrix(NewRow, .ColIndex("PONumber")) = ""
                    .TextMatrix(NewRow, .ColIndex("EstItemID")) = "0"
                    .TextMatrix(NewRow, .ColIndex("BudgetGenerated")) = "False"
                    .TextMatrix(NewRow, .ColIndex("BudgetPostingBatch")) = "0"
                    .TextMatrix(NewRow, .ColIndex("BudgetsLocked")) = .TextMatrix(r, .ColIndex("BudgetsLocked"))
                    .TextMatrix(NewRow, .ColIndex("BudgetDeleted")) = .TextMatrix(r, .ColIndex("BudgetsLocked"))
                    Call ColorizeItems(NewRow)
                Next
                
            
            Case mcITEM_SPLITITEMS
                Call SplitItems
            
            
            
            Case mcITEM_CANCELBUDGETS
                Call CancelBudgets
                
            Case mcITEM_COMPAREPRICES
                r = .Row
                If r < 1 Then Exit Sub
                Call FPriceComparison.ShowForm(.TextMatrix(r, .ColIndex(IIf(mMode <> fcePO, "BudgetVendor", "POVendor"))), _
                                               mCommunity, _
                                               mCommunityPhase, _
                                               .TextMatrix(r, .ColIndex("Assembly")), _
                                               .TextMatrix(r, .ColIndex("Model")), _
                                               .TextMatrix(r, .ColIndex("EstPhase")), _
                                               .TextMatrix(r, .ColIndex("EstItem")), _
                                               .TextMatrix(r, .ColIndex("ItemDesc")))
              
            Case mcITEM_UPDATEPRICELIST
                r = .Row
                If r < 1 Then Exit Sub
                Call FPriceListUpdate.ShowForm(mMode <> fcePO, gItems, mCommunity, mCommunityPhase, FMain.CD_Community, .TextMatrix(r, .ColIndex("AssemblyType")))
            
        End Select
        Call GroupGrid
    End With
End Sub
Private Sub txtTerms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(3)
End Sub

Private Sub txtVendor_GotFocus()
    SelectAll txtVendor
End Sub


Private Sub SaveDBAssembly()
    Dim f As FAssembly
    Dim r As Long
    Dim rs As Recordset
    Dim AssemblyType As AssemblyTypes
    Dim WBS(40) As String
    If mEstAssemblyID <> -1 Then
        
        
        mCurrentEstAssemblyID = mEstAssemblyID
        mCurrentAssemblyDesc = txtHFDescription.Text
        
        
        Set rs = HFApp.SqlExec("SELECT DISTINCT Community,AssemblyType,Assembly,Model,OptionID FROM EstimatedItems WHERE EstAssemblyID=" & DbQuote(Num, mCurrentEstAssemblyID))
        mCurrentAssemblyType = Val("" & rs("AssemblyType"))
        mCurrentModel = "" & rs("Model")
        mCurrentOptionID = "" & rs("OptionID")
        mCurrentCommunity = "" & rs("Community")
        mCurrentAssembly = "" & rs("Assembly")
        
        
        AssemblyType = mCurrentAssemblyType
        
        
        If AssemblyType = -1 Then
            If mCurrentOptionID = "CUSTOM" Then
                AssemblyType = atoption
                If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Assembly Type", "select 2 TypeID,'Model Option' [Assembly Type] union select 3,'Global Option' union select 4,'Design Center Option'", "" & AssemblyType, , , , "TypeID") Then Exit Sub
                AssemblyType = FPickList.SelectedItem("TypeID")
            Else
                AssemblyType = atModel
                If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Assembly Type", "select 0 TypeID,'Model' [Assembly Type] union select 2,'Model Option' union select 3,'Global Option' union select 4,'Design Center Option'", "" & AssemblyType, , , , "TypeID") Then Exit Sub
                AssemblyType = FPickList.SelectedItem("TypeID")
            End If
        End If
        
        Set f = New FAssembly
        If f.ShowForm(True, mCurrentCommunity, mCurrentModel, mCurrentOptionID, txtHFOption.Text, mCurrentAssemblyDesc, AssemblyType) Then
            With gItems
            For r = 1 To .Rows - 1
            If Not .IsSubtotal(r) Then
                Call f.AddItem("", "", .TextMatrix(r, .ColIndex("Model")), _
                               .TextMatrix(r, .ColIndex("EstPhase")), .TextMatrix(r, .ColIndex("EstItem")), .TextMatrix(r, .ColIndex("Sequence")), _
                               .TextMatrix(r, .ColIndex("ItemDesc")), .ValueMatrix(r, .ColIndex("BudgetQty")), _
                               .TextMatrix(r, .ColIndex("OrderUOM")), .ValueMatrix(r, .ColIndex("TakeoffQty")), _
                               .TextMatrix(r, .ColIndex("TakeoffUOM")), .ValueMatrix(r, .ColIndex("ConversionFactor")), _
                               .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")), _
                               .ValueMatrix(r, .ColIndex("WastePercent")), .TextMatrix(r, .ColIndex("JCExtra")), _
                               .TextMatrix(r, .ColIndex("JCCostCode")), .TextMatrix(r, .ColIndex("JCCostCodeDesc")), _
                               .TextMatrix(r, .ColIndex("JCCategory")), .TextMatrix(r, .ColIndex("JCCategoryDesc")), _
                               .TextMatrix(r, .ColIndex("BudgetVendor")), .TextMatrix(r, .ColIndex("BudgetVendorName")), _
                               .ValueMatrix(r, .ColIndex("BudgetRate")), .TextMatrix(r, .ColIndex("BudgetTaxGroup")), _
                               "", _
                               .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")), _
                               .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")), _
                               .TextMatrix(r, .ColIndex("POIndex")), _
                               .TextMatrix(r, .ColIndex("ItemComments")), .TextMatrix(r, .ColIndex("Formula")), 0, .TextMatrix(r, .ColIndex("Location")), WBS)
            End If
            Next
            End With
                        
        End If
        
    End If
   
End Sub


Public Sub mnuEstimateItemsVendorsSub_Click(Index As Integer)
On Error GoTo eh
    
    Dim Vendor As String
    
    Select Case Index
    
        Case mcVENDOR_EDIT
            Vendor = gAssemblies.Cell(flexcpText, gAssemblies.Row, 1)
            Call HFApp.EditVendor(Vendor)
            
            
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "mnuEstimateItemsVendorsSub_Click")
End Sub


Public Sub mnuEstimateItemsPOSub_Click(Index As Integer)
On Error GoTo eh
    Dim s As String
    Dim Vendor As String
    Dim PONumber As String
    
    With gAssemblies
    Select Case Index
    
            
        Case mcPO_SEND
            Call SendingWizard("PO", , SelectedPOs)
            
        Case mcPO_PREVIEW
            Call ShowSelectedPOs(True)
            
        Case mcPO_PRINT
            Call ShowSelectedPOs(False)
            Screen.MousePointer = vbDefault
            
        Case mcPO_EDITVENDOR
            Vendor = ""
            PONumber = Parse(.Cell(flexcpText, .Row, 1), 1, Chr(2))
            On Error Resume Next
            Vendor = HFApp.SqlExec("select vendor from pomaster where DivisionID = " & HFApp.DivisionID & " and ponumber=" & DbQuote(Str, PONumber), dbHomefront)(0)
            On Error GoTo eh
            Call HFApp.EditVendor(Vendor)
            
        Case mcPO_CHANGE
            PONumber = Parse(.Cell(flexcpText, .Row, 1), 1, Chr(2))
            If SaveData(True) Then If ChangePOVendor(PONumber) Then Call RefreshTree
        
        Case mcPO_CANCEL
            PONumber = Parse(.Cell(flexcpText, .Row, 1), 1, Chr(2))
            If SaveData(True) Then If CancelPO(PONumber) Then Call RefreshTree
            
        Case mcPO_REPRICE
            PONumber = Parse(.Cell(flexcpText, .Row, 1), 1, Chr(2))
            Call FPOPriceChangeWiz.ShowForm(PONumber)
            
    End Select
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "mnuEstimateItemsPOSub_Click")
End Sub

Private Sub RefreshTree()
On Error GoTo eh
    Dim i As Long
    Dim n As VSFlexNode
    Dim s As String
    Dim origRow As Long
    Dim done As Boolean
    Dim Key As String
    
    With gAssemblies
        .Redraw = flexRDNone
        
        'get path to this node
        s = ""
        origRow = .Row
        Set n = .GetNode
        While Not n Is Nothing
            s = .Cell(flexcpText, n.Row, 1) & Chr(1) & s
            Set n = n.GetNode(flexNTParent)
        Wend
        s = Left(s, Len(s) - 1)
        
        'collapse tree
        Call LoadAssemblies(True)
        Call ClearItems
        
        'try to navigate back to that node
        done = False
        i = 1
        Key = Parse(s, i, Chr(1))
        .Row = 1
        While Not done
            If Key = .Cell(flexcpText, .Row, 1) Then
                
                Call LoadAssemblies(False)
                
                i = i + 1
                done = i > Parse(s, , Chr(1))
                If Not done Then
                    Key = Parse(s, i, Chr(1))
                End If
            
            End If
            If Not done Then
                If .Row = .Rows - 1 Then
                    'cant find it so assume it's the same row
                    .Row = origRow
                    done = True
                Else
                    .Row = .Row + 1
                End If
            End If
        Wend
        
eh: 'ignore errors just get out
        .Redraw = flexRDBuffered
    End With
End Sub


Public Sub mnuEstimateItemsAssemblySub_Click(Index As Integer)
On Error GoTo eh
    Dim SrcJob As String
    Dim Customer As String
    Dim ChangeOrder As String
    Dim AssemblyID As Long
    Dim ChangeStatus As String
    Dim j As String
    Dim coNumber As Integer
    Dim b As Boolean
    Dim s As String
    Dim r As Long
    Dim X As Long
    Dim ID As Long
    Dim Quote As String
    Dim Description  As String
    Dim rs As Recordset
    
    
    Dim rpt As String
    'Dim f As FRptViewer
    
    With gAssemblies
    Select Case Index
    
        Case mcASMBlY_PRINT
            Customer = ""
            ChangeOrder = ""
            AssemblyID = 0
            Select Case .Cell(flexcpData, .Row, 1)
                Case "Customer_no"
                    rpt = "JobQuote.rpt"
                    Customer = .Cell(flexcpText, .Row, 1)
                Case "CORSort"
                    rpt = "JobChangeOrder.rpt"
                    Customer = .Cell(flexcpText, .GetNodeRow(.Row, flexNTParent), 1)
                    ChangeOrder = .Cell(flexcpText, .Row, 1)
                    If Left(ChangeOrder, 6) = "ZZZZZ" & Chr(1) Then
                        ChangeStatus = Mid(ChangeOrder, 7)
                        ChangeOrder = ""
                        rpt = "JobChangeRequests.rpt"
                    End If
                Case "EstAssemblyID"
                    rpt = "JobAssembly.rpt"
                    AssemblyID = Val(.Cell(flexcpText, .Row, 1))
                    s = .Cell(flexcpText, .GetNodeRow(.Row, flexNTParent), 1)
                    If Left(s, 6) = "ZZZZZ" & Chr(1) Then rpt = "JobChangeRequest.rpt"
            End Select
            rpt = PathAppend(HFApp.SystemFolder, "Estimating\Reports", rpt)
'MsgBox rpt & vbCrLf & "Customer=" & Customer & vbCrLf & "ChangeOrder=" & ChangeOrder & vbCrLf & "changeStatus=" & ChangeStatus & vbCrLf & "AssemblyID=" & AssemblyID
            'Set f = New FRptViewer
            'Call f.ShowReport(rpt, True, False, "Customer", Customer, "ChangeOrder", ChangeOrder, "AssemblyID", AssemblyID, "ChangeStatus", ChangeStatus, "Job", Job)
            
            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(rpt, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            Call c.ParameterValue("Customer", Customer)
            Call c.ParameterValue("ChangeOrder", ChangeOrder)
            Call c.ParameterValue("AssemblyID", AssemblyID)
            Call c.ParameterValue("ChangeStatus", ChangeStatus)
            Call c.ParameterValue("Job", Job)
            On Error GoTo eh
            Call c.PrintPreview("Print Preview")
        
        Case mcASMBlY_FINALIZE
        
            If Not SaveData(True) Then Exit Sub
            Select Case .Cell(flexcpData, .Row, 1)
                Case "EstAssemblyID"
                    s = ""
                    s = s & "update estimateassemblies"
                    s = s & "   set budgetslocked=" & IIf(FMain.mnuEstimateItemsAssemblySub(Index).checked, 0, 1)
                    s = s & "      ,budgetslockeddate=getdate()"
                    s = s & "      ,budgetslockedby=" & DbQuote(Str, HFApp.LoginID)
                    s = s & " where estassemblyid=" & DbQuote(Num, .Cell(flexcpText, .Row, 1))
                    Call HFApp.SqlExec(s, dbHomefront)
                    
                Case "ChangeOrder", "CORSort"
                    s = ""
                    s = s & "update estimateassemblies"
                    s = s & "   set budgetslocked=" & IIf(FMain.mnuEstimateItemsAssemblySub(Index).checked, 0, 1)
                    s = s & "      ,budgetslockeddate=getdate()"
                    s = s & "      ,budgetslockedby=" & DbQuote(Str, HFApp.LoginID)
                    s = s & " where job=" & DbQuote(Str, mJob)
                    s = s & "   and DivisionID = " & HFApp.DivisionID & " and isnull(changeorder,'')=" & DbQuote(Str, .Cell(flexcpText, .Row, 1))
                    Call HFApp.SqlExec(s, dbHomefront)
            
                    
                Case "Customer_No"
                    s = ""
                    s = s & "update estimateassemblies"
                    s = s & "   set budgetslocked=" & IIf(FMain.mnuEstimateItemsAssemblySub(Index).checked, 0, 1)
                    s = s & "      ,budgetslockeddate=getdate()"
                    s = s & "      ,budgetslockedby=" & DbQuote(Str, HFApp.LoginID)
                    s = s & " where customer_no=" & DbQuote(Str, .Cell(flexcpText, .Row, 1))
                    Call HFApp.SqlExec(s, dbHomefront)
                    
            End Select
            
            
            'update me and all my children
            Call LockMeAndMyChildren(.Row, Not FMain.mnuEstimateItemsAssemblySub(Index).checked)
            
            
            'update toolbar
            Call BudgetsAreLocked(Not FMain.mnuEstimateItemsAssemblySub(Index).checked, VBA.Date(), HFApp.LoginID)
            
            'clear rhs
            gItems.Rows = 1
            frmAssembly.Visible = False
            frmPOIndex.Visible = False
            Call Form_Resize
            
            
        
        Case mcASMBlY_COPYFROMJOB
            If Not SaveData(True) Then Exit Sub
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Job", "select job_no Job,Description,Municipal_Address Address from tbljobs where DivisionID = " & HFApp.DivisionID) Then Exit Sub
            
            Customer = .Cell(flexcpText, .GetNodeRow(.Row, flexNTRoot), 1)
            SrcJob = FPickList.SelectedItem("Job")
            s = "exec Purch_CopyJobEstimates " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, SrcJob) & "," & DbQuote(Str, mJob) & "," & DbQuote(Str, Customer)
            Call HFApp.SqlExec(s, dbHomefront)
        
            'reload
            Call mnuEstimateItemViewsSub_Click(CInt(mViewIndex + 1))
        
        
        
        Case mcASMBlY_ATTACHQUOTE
            If Not SaveData(True) Then Exit Sub
            
            'select quote
            s = ""
            s = s & "SELECT DISTINCT j.Job_No Quote" & vbCrLf
            s = s & "      ,j.Description" & vbCrLf
            s = s & "      ,j.Community" & vbCrLf
            s = s & "      ,j.CommunityPhase Phase" & vbCrLf
            s = s & "      ,j.Municipal_Address Address" & vbCrLf
            s = s & "      ,j.PM,Purchaser" & vbCrLf
            s = s & "      ,j.Estimator" & vbCrLf
            s = s & "FROM tblCustomers c JOIN tblJobs j ON (c.Job_No = j.Job_No and j.DivisionID = c.DivisionID)" & vbCrLf
            s = s & "    join system_setup s on s.id = j.DivisionID" & vbCrLf
            s = s & "WHERE ISNULL(j.DivisionID,0) = " & HFApp.DivisionID & " and isnull(j.Inactive,0)=0 AND c.job_no<>'' AND " & vbCrLf
            s = s & "((c.purchased=1 AND c.cancelled=0 AND c.inactive=0 AND ((isnull(s.accounting_approve,0)=1 AND c.approved=1) OR (isnull(s.accounting_approve,0)=0 AND c.contract_assigned=1)))" & vbCrLf
            s = s & "  OR " & vbCrLf
            s = s & "(c.home_selection <>'PreSale' OR c.presale_selection<>'PreSale'))" & vbCrLf
            s = s & "  AND j.isquote=1"
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Quote", s) Then Exit Sub
            Quote = FPickList.SelectedItem("Quote")
            
            'get assembly in quote
            s = "select EstAssemblyID,HFDescription Description from estimateassemblies where divisionid = " & HFApp.DivisionID & " and job=" & DbQuote(Str, Quote)
            Set rs = HFApp.SqlExec(s, dbHomefront)
            If rs.EOF Then
                MsgBox "This quote is empty", vbInformation, App.ProductName
                Exit Sub
            Else
                ID = Val("" & rs("EstAssemblyID"))
                rs.MoveNext
                If Not rs.EOF Then
                    'has more than one estassembly
                    If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Quoted Items", s, , , , , "EstAssemblyID") Then Exit Sub
                    ID = Val(FPickList.SelectedItem("EstAssemblyID"))
                End If
            End If
            
            'write quote number on selected assembly
            s = "update estimateassemblies set quote=" & DbQuote(Str, Quote) & " where estassemblyid=" & DbQuote(Num, .Cell(flexcpText, .Row, 1))
            Call HFApp.SqlExec(s, dbHomefront)
            
            ' Copy quote/ID to mJob/assembly
            s = ""
            s = s & "insert into EstimateItems(divisionid,POIndex,Phase,Item,JCExtra,JCCostCode,JCCategory,SortOrder,Description,Comments,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,PONumber,ExcludeFromPO,OriginalJCCategory,BudgetGenerated,BudgetPostingBatch,POGenBatch,BudgetDeleted,PODeleted,BudgetOverridden,POOverridden,Assembly,AssemblyDescription,Model,WBS01,WBS02,WBS03,WBS04,WBS05,WBS06,WBS07,WBS08,WBS09,WBS10,WBS11,WBS12,WBS13,WBS14,WBS15,WBS16,WBS17,WBS18,WBS19,WBS20,WBS21,WBS22,WBS23,WBS24,WBS25,WBS26,WBS27,WBS28,WBS29,WBS30,WBS31,WBS32,WBS33,WBS34,WBS35,WBS36,WBS37,WBS38,WBS39,WBS40,Location,RFP" & vbCrLf
            s = s & "      ,EstAssemblyID" & vbCrLf
            s = s & "      ,Job)" & vbCrLf
            s = s & "select divisionid,POIndex,Phase,Item,JCExtra,JCCostCode,JCCategory,SortOrder,Description,Comments,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,PONumber,ExcludeFromPO,OriginalJCCategory,BudgetGenerated,BudgetPostingBatch,POGenBatch,BudgetDeleted,PODeleted,BudgetOverridden,POOverridden,Assembly,AssemblyDescription,Model,WBS01,WBS02,WBS03,WBS04,WBS05,WBS06,WBS07,WBS08,WBS09,WBS10,WBS11,WBS12,WBS13,WBS14,WBS15,WBS16,WBS17,WBS18,WBS19,WBS20,WBS21,WBS22,WBS23,WBS24,WBS25,WBS26,WBS27,WBS28,WBS29,WBS30,WBS31,WBS32,WBS33,WBS34,WBS35,WBS36,WBS37,WBS38,WBS39,WBS40,Location,RFP" & vbCrLf
            s = s & "      ," & DbQuote(Num, .Cell(flexcpText, .Row, 1)) & vbCrLf
            s = s & "      ," & DbQuote(Str, mJob) & vbCrLf
            s = s & "  from estimateitems where estassemblyid=" & DbQuote(Num, ID) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)

            'load items
            Call gAssemblies_KeyDown(vbKeyReturn, 0)


    
        Case mcASMBlY_DELETE
            b = False
            On Error Resume Next
            b = Val("" & HFApp.SqlExec("select count(*) from estimateitems where (budgetdeleted = 0 or podeleted = 0) and estassemblyid=" & DbQuote(Num, Val(.Cell(flexcpText, .Row, 1))), dbHomefront)(0)) = 0
            On Error GoTo 0
            
            If b Then
                'assembly has no items. it can be removed.
                b = vbYes = MsgBox("Are you sure you want to remove this item?", vbYesNo + vbQuestion, App.ProductName)
            Else
                Call MsgBox(vbQuote & .Cell(flexcpText, .Row, 0) & vbQuote & " has items assigned to it so it cannot" & vbCrLf & "be deleted. Move or delete the items then try again.", vbOK + vbInformation, App.ProductName)
            End If
                
            If b Then
            
                If Me.frmAssembly.Visible And mEstAssemblyID = Val(.Cell(flexcpText, .Row, 1)) Then
                    frmAssembly.Visible = False
                    Toolbar.Buttons("SaveAssembly").Enabled = False
                    frmPOIndex.Visible = False
                    gItems.Rows = 1
                    Call Form_Resize
                End If
            
                s = "delete from estimateassemblies where estassemblyid=" & DbQuote(Num, Val(.Cell(flexcpText, .Row, 1)))
                Call HFApp.SqlExec(s, dbHomefront)
                s = "delete from estimateitems where estassemblyid=" & DbQuote(Num, Val(.Cell(flexcpText, .Row, 1)))
                Call HFApp.SqlExec(s, dbHomefront)
                .RemoveItem
            End If
            
        
        
        Case mcASMBlY_RENAME
            If .Cell(flexcpData, .Row, 1) = "CORSort" Then
                s = Left(InputBox(vbCrLf & vbCrLf & vbCrLf & "Enter a new name for this change order", "Rename Change Order", .Cell(flexcpText, .Row, 0)), 10)
                If s <> "" Then
                    b = Dirty
                    
                    ChangeOrder = .Cell(flexcpText, .Row, 0)
                    .Cell(flexcpText, .Row, 0) = s
                    If Me.frmAssembly.Visible Then Me.txtHFDescription.Text = .Cell(flexcpText, .Row, 0)
                    s = ""
                    s = s & "update estimateassemblies" & vbCrLf
                    s = s & "   set changeorder=" & DbQuote(Str, .Cell(flexcpText, .Row, 0)) & vbCrLf
                    s = s & " where DivisionID = " & HFApp.DivisionID & " and job=" & DbQuote(Str, mJob) & vbCrLf
                    s = s & "   and changeorder=" & DbQuote(Str, ChangeOrder) & vbCrLf
                    Call HFApp.SqlExec(s, dbHomefront)
                    Dirty = b
                End If
            Else
                s = InputBox(vbCrLf & vbCrLf & vbCrLf & "Enter a new name for this item", "Rename Item", .Cell(flexcpText, .Row, 0))
                If s <> "" Then
                    b = Dirty
                    .Cell(flexcpText, .Row, 0) = Left(Trim(s), 200)
                    
                    If Me.frmAssembly.Visible And mEstAssemblyID = Val(.Cell(flexcpText, .Row, 1)) Then
                        Me.txtHFDescription.Text = .Cell(flexcpText, .Row, 0)
                    End If
                    
                    s = ""
                    s = s & "update estimateassemblies" & vbCrLf
                    s = s & "   set hfdescription=" & DbQuote(Str, .Cell(flexcpText, .Row, 0)) & vbCrLf
                    s = s & " where estassemblyid=" & DbQuote(Num, Val(.Cell(flexcpText, .Row, 1))) & vbCrLf
                    Call HFApp.SqlExec(s, dbHomefront)
                    Dirty = b
                End If
            End If
    
    
        Case mcASMBlY_ADD, mcASMBlY_ADDCO, mcASMBlY_ADDCR
            Select Case Index
                Case mcASMBlY_ADDCO
                    'insert change order, add change order node
                    Customer = .Cell(flexcpText, .Row, 1)
                    ChangeStatus = ""
                    coNumber = HFApp.SqlExec("select dbo.GetNextCONumber(" & DbQuote(Str, Customer) & ",1)")(0)
                    s = ""
                    's = s & "update tblcustomers set last_co=isnull(cast(last_co as integer),0)+1 where customer_no=" & DbQuote(Str, Customer) & vbCrLf
                    s = s & "insert into changeordermaster(customer_no,change_order_no,purchasingco,usealtjccodes) Values(" & DbQuote(Str, Customer) & "," & "'CO_" & coNumber & "',1,1)"
                    Call HFApp.SqlExec(s, dbHomefront)
                    ChangeOrder = "CO_" & coNumber 'HFApp.SqlExec("select last_co from tblcustomers where customer_no=" & DbQuote(Str, Customer), dbHomeFront)(0)
                    r = .GetNode().AddNode(flexNTLastChild, ChangeOrder).Row
                    .Cell(flexcpData, r, 0) = "ChangeOrder"
                    .Cell(flexcpText, r, 1) = ChangeOrder
                    .Cell(flexcpData, r, 1) = "CORSort"
                    .Cell(flexcpText, r, 2) = msNone
                    .Cell(flexcpText, r, 3) = msNone
                    .Cell(flexcpText, r, .ColIndex("IsChangeRequest")) = "0"
                    Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(.Cell(flexcpValue, r, 2), .Cell(flexcpValue, r, 3), .Cell(flexcpValue, r, .ColIndex("IsChangeRequest")), .Cell(flexcpText, r, .ColIndex("POIconType")))
                    .RowData(r) = " AND Job_No=" & DbQuote(Str, mJob) & " AND Customer_No=" & DbQuote(Str, Customer) & " AND ChangeOrder=" & DbQuote(Str, ChangeOrder)
                    
                    'add assembly node
                    r = .GetNode(r).AddNode(flexNTLastChild, "untitled").Row
                    
                    
                Case mcASMBlY_ADDCR
                    'add request node
                    Customer = .Cell(flexcpText, .Row, 1)
                    ChangeStatus = "Pending"
                    ChangeOrder = ""
                    
                    X = .FindRow("ZZZZZ" & Chr(1) & "Pending", r, 1, False, True)
                    If X = -1 Then
                        r = .GetNode().AddNode(flexNTLastChild, "Pending requests").Row
                        .Cell(flexcpData, r, 0) = "CORNumber" 'display field
                        .Cell(flexcpText, r, 1) = "ZZZZZ" & Chr(1) & "Pending"  'key value
                        .Cell(flexcpData, r, 1) = "CORSort" 'key field
                        .Cell(flexcpText, r, 2) = msNone
                        .Cell(flexcpText, r, 3) = msNone
                        .Cell(flexcpText, r, .ColIndex("IsChangeRequest")) = "1"
                        Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(.Cell(flexcpValue, r, 2), .Cell(flexcpValue, r, 3), .Cell(flexcpValue, r, .ColIndex("IsChangeRequest")), .Cell(flexcpText, r, .ColIndex("POIconType")))
                        .RowData(r) = " AND Job_No=" & DbQuote(Str, mJob) & " AND Customer_No=" & DbQuote(Str, Customer) & " AND CORSort=" & DbQuote(Str, ChangeOrder)

                    Else
                        r = X
                    End If
                    
                    'add assembly node
                    .Row = r
                    r = .GetNode(.Row).AddNode(flexNTLastChild, "untitled").Row
                    .Cell(flexcpText, r, .ColIndex("IsChangeRequest")) = .Cell(flexcpText, .Row, .ColIndex("IsChangeRequest"))
                    
                Case mcASMBlY_ADD
                    'add assembly node
                    ChangeStatus = ""
                    Customer = .Cell(flexcpText, .GetNodeRow(.Row, flexNTParent), 1)
                    ChangeOrder = .Cell(flexcpText, .Row, 1)
                    
                    If Left(ChangeOrder, 6) = "ZZZZZ" & Chr(1) Then
                        ChangeStatus = Mid(ChangeOrder, 7)
                        ChangeOrder = ""
                    End If
                    r = .GetNode().AddNode(flexNTLastChild, "untitled").Row
                    .Cell(flexcpText, r, .ColIndex("IsChangeRequest")) = .Cell(flexcpText, .Row, .ColIndex("IsChangeRequest"))
                    
            End Select
            
            j = GetJob(Customer)
            
            
            s = ""
            s = s & "INSERT INTO EstimateAssemblies(Customer_No,EstimateIndex,AssemblyType,OptionType,SalesQty,Job,HFDescription,ChangeOrder,ChangeRequestStatus,ChangeRequestedBy,DivisionID,ChangeRequestedDate)" & vbCrLf
            s = s & "SELECT Customer_No" & vbCrLf
            s = s & "      ,isnull(LastEstimateIndex,0)+1" & vbCrLf
            s = s & "      ,-1" & vbCrLf
            s = s & "      ,-1" & vbCrLf
            s = s & "      ,1" & vbCrLf
            s = s & "      ," & DbQuote(Str, j) & vbCrLf
            s = s & "      ,'untitled'" & vbCrLf
            s = s & "      ," & DbQuote(Str, ChangeOrder) & vbCrLf
            s = s & "      ," & DbQuote(Str, ChangeStatus) & vbCrLf
            s = s & "      ," & IIf(ChangeStatus = "", "''", DbQuote(Str, HFApp.LoginID)) & vbCrLf
            s = s & "      ,DivisionID" & vbCrLf
            s = s & "      ," & IIf(ChangeStatus = "", "NULL", DbQuote(DateTime, Now())) & vbCrLf
            s = s & "  FROM tblCustomers" & vbCrLf
            s = s & " WHERE Customer_No=" & DbQuote(Str, Customer)
            Call HFApp.SqlExec(s)
            ID = HFApp.SqlIdentity("EstimateAssemblies")
            
            s = ""
            s = s & "UPDATE tblCustomers" & vbCrLf
            s = s & "   SET LastEstimateIndex=LastEstimateIndex+1" & vbCrLf
            s = s & " WHERE Customer_No=" & DbQuote(Str, Customer) & vbCrLf
            Call HFApp.SqlExec(s)
            
            .Cell(flexcpData, r, 0) = "HFDescription"
            .Cell(flexcpText, r, 1) = ID
            .Cell(flexcpData, r, 1) = "EstAssemblyID"
            .Cell(flexcpText, r, 2) = msNone
            .Cell(flexcpText, r, 3) = msNone
            Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(msNone, msNone, .Cell(flexcpText, r, .ColIndex("IsChangeRequest")) = "1", .Cell(flexcpText, r, .ColIndex("POIconType")))
            
            .RowData(r) = " AND EstAssemblyID = " & DbQuote(Str, ID)
            
            Call SetCtrlFocus(gAssemblies)
            .Row = r
            Call .Select(r, 0)
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
            .Cell(flexcpFontBold, .Row, 0) = True
            Call LoadItems
                        
            Call mnuEstimateItemsAssemblySub_Click(mcASMBlY_RENAME)
        
    End Select
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "mnuEstimateItemsAssemblySub_Click")
End Sub





Private Function GetJob(Customer As String) As String
'recursive lookup to find the last customer in the chain
    Dim s As String
    Dim rs As Recordset
    s = "select job_no,sold,sold_to_customer from tblcustomers where customer_no=" & DbQuote(Str, Customer)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        Err.Raise 9999, "GetJob", "Broken chain of sold_to_customers"
    Else
        If "" & rs("sold") = "True" Then
            GetJob = GetJob("" & rs("sold_to_customer"))
        Else
            GetJob = "" & rs("job_no")
        End If
    End If
    
End Function




Private Sub gItems_RowColChange()
On Error Resume Next
    With gItems
        mFunnyFlag = True
        Call ShowTip(.Row, .Col, .colPos(.Col) + 180, .RowPos(.Row) + .RowHeight(.Row) + 180)
    End With
End Sub

Private Sub ShowTip(Row As Long, Col As Long, X As Long, Y As Long, Optional tip As String)
    With gItems
        If Row < 0 Or Col < 0 Or (.RowSel = .Row And tip <> "") Then
            picWarningMessages.Visible = False
        Else
            lblWarningMessages = IIf(tip <> "", tip, gItems.Cell(flexcpData, Row, Col))

            imgTipIcon.Picture = IIf(tip <> "", imgInfo.Picture, imgWarning.Picture)

            picWarningMessages.Visible = lblWarningMessages <> ""

            Set Me.Font = lblWarningMessages.Font

            lblWarningMessages.Alignment = IIf(tip <> "", 1, 0)
            lblWarningMessages.Width = Me.TextWidth(lblWarningMessages) + lblWarningMessages.Left
            lblWarningMessages.Height = Me.TextHeight(lblWarningMessages) + lblWarningMessages.Top

            picWarningMessages.Width = lblWarningMessages.Width + lblWarningMessages.Left + lblWarningMessages.Top
            picWarningMessages.Height = lblWarningMessages.Height + 2 * lblWarningMessages.Top
            If .Height - Y - picWarningMessages.Height - 365 < 0 Then Y = .Height - picWarningMessages.Height - 365
            If .Width - X - picWarningMessages.Width - 365 < 0 Then X = .Width - picWarningMessages.Width - 365
            picWarningMessages.Top = Y
            picWarningMessages.Left = X
        End If
    End With
End Sub

Private Sub LoadAssemblies(ClearTree As Boolean)
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
    Dim DisplayFld As String
    
    Dim SortFld    As String
    Dim KeyValue   As String
    Dim DisplayValue   As String
    Dim IconKey As String

'PROPERTY                  CONTAINS
'.rowdata()                Where clause
'.cell(flexcpText, r, 0)   DisplayFld Value
'.cell(flexcpData, r, 0)   DisplayFld Field Name (also used as tooltip)
'.Cell(flexcpText, r, 1)   KeyFld Value
'.Cell(flexcpData, r, 1)   KeyFld Field Name
'.Cell(flexcpText, r, 2)   Completed  As MultiStateEnum
'.Cell(flexcpText, r, 3)   Selected   as MultiStateEnum

    With gAssemblies
        .AllowUserResizing = flexResizeColumns
        '.Redraw = flexRDNone
        .WordWrap = True
        
        '.ColWidthMax = 5500
        .TextMatrix(0, 0) = mViews(mViewIndex).Name
        levels = Parse(mViews(mViewIndex).DisplayFlds)
        .TextMatrix(0, .ColIndex("Budgeted")) = "Budgeted"
        .TextMatrix(0, .ColIndex("Changes")) = "Changes"
        .TextMatrix(0, .ColIndex("Commited")) = "Committed"
        
    
        If ClearTree Then .Rows = 1


        If .Rows = 1 Then
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
            IconKey = Parse(mViews(mViewIndex).IconKeys, Level + 2)
            SortFld = Parse(mViews(mViewIndex).SortFlds, Level + 2)
            

            s = ""
            s = s & "SELECT " & KeyFld & "," & DisplayFld & vbCrLf
            If mMode <> fcePO Then
                s = s & "      ,sum(case when Budgetdeleted=1 then 0 else 1 end ) TotalItems" & vbCrLf
                s = s & "      ,SUM(ISNULL(CAST(BudgetGenerated AS INT),0)) CompletedItems" & vbCrLf
            Else
                s = s & "      ,COUNT(*) TotalItems" & vbCrLf
                s = s & "      ,SUM(CASE ISNULL(PONumber,'') WHEN '' THEN 0 ELSE 1 END) Completed" & vbCrLf
            End If
            s = s & "      ,MAX(POStatus) POStatus" & vbCrLf
            s = s & "      ,SUM(CASE WHEN IsChange=0 and BudgetDeleted<>1 THEN isnull(BudgetPretax,0) + isnull(BudgetJCTax,0) ELSE 0 END) Budgeted" & vbCrLf
            s = s & "      ,SUM(CASE WHEN IsChange=1 and BudgetDeleted<>1 THEN isnull(BudgetPretax,0) + isnull(BudgetJCTax,0) ELSE 0 END) Changes" & vbCrLf
            s = s & "      ,SUM(CASE WHEN POGenBatch=0 or poDeleted=1 THEN 0 ELSE isnull(POPretax,0)+isnull(POJCTax,0) END) Commited" & vbCrLf
            s = s & "      ,MIN(CAST(BudgetsLocked AS INT)) BudgetsLocked" & vbCrLf
            s = s & "      ,MIN(HFLocation) HFLocation" & vbCrLf
            s = s & "      ,MIN(IsChangeRequest) IsChangeRequest" & vbCrLf
            s = s & "      ,MIN(POIconType) POIconType" & vbCrLf
            s = s & "  FROM Estimateditems i" & vbCrLf
            On Error Resume Next
            WhereClause = .RowData(.Row)
            On Error GoTo eh
            
            If WhereClause = "" Then
                If ProjectBased Then
                    s = s & " WHERE i.DivisionID = " & HFApp.DivisionID & " and isquote=0 and Community=" & DbQuote(Str, mCommunity) & vbCrLf
                Else
                    If mViews(mViewIndex).Name = "Customer" Then
                        s = s & " WHERE i.DivisionID = " & HFApp.DivisionID & " and AssemblyJob_No=" & DbQuote(Str, mJob) & vbCrLf
                    Else
                        s = s & " WHERE i.DivisionID = " & HFApp.DivisionID & " and Job_No=" & DbQuote(Str, mJob) & vbCrLf
                    End If
                End If
            Else
                s = s & " WHERE " & Mid(WhereClause, 6) & vbCrLf
            End If
            s = s & " and (isnull(budgetdeleted,0)<>1 OR isnull(podeleted,0)<>1)" & vbCrLf
            
            
            If mViews(mViewIndex).Name = "RFQ's" Then
                s = s & "   AND RFP<>0" & vbCrLf
            End If
            
            Select Case KeyFld
                Case ""
                Case "EstAssemblyID"
                    s = s & "GROUP BY " & KeyFld & "," & DisplayFld & ",seq" & vbCrLf
                Case "ChangeOrder"
                    s = s & "GROUP BY " & KeyFld & vbCrLf
                Case "POIndex"
                    s = s & "GROUP BY POIndex,POIndexDescription" & vbCrLf
                Case Else
                    s = s & "GROUP BY " & KeyFld & "," & DisplayFld & vbCrLf
            End Select
                       
            Select Case SortFld
                Case "ChangeOrder", "CORSort"
                    s = s & "ORDER BY 1"
                Case "Seq"
                    s = s & "ORDER BY seq"
                Case Else
                    If HFApp.Options.ValueByName("SortJobAssemblybySeq") = "True" Then
                        s = s & "ORDER BY 1" ' & SortFld
                    Else
                        s = s & "ORDER BY 2"    ' & SortFld
                    End If
            End Select
            

            mViewQuery = s
 
            Set rs = HFApp.SqlExec(s)
            
            
            'remove any children this node already has
            If .Row > 0 Then
                Set n = .GetNode().GetNode(flexNTFirstChild)
                While Not n Is Nothing
                    Call n.RemoveNode
                    Set n = .GetNode().GetNode(flexNTFirstChild)
                Wend
            End If

            'now add new children
            While Not rs.EOF
                KeyValue = "" & rs(0)
                DisplayValue = Trim("" & rs(1))
                If DisplayValue = "" Then DisplayValue = " -- "

                If Level = -1 Then
                    r = .Rows
                    Call .AddItem(DisplayValue, r)
                    
                    If IsIn(KeyFld, "PONumber", "PONumberVendor") Then
                        .Cell(flexcpText, r, .ColIndex("POIconType")) = "" & rs("POIconType")
                        If .Cell(flexcpText, r, .ColIndex("POIconType")) = "" Then .Cell(flexcpText, r, .ColIndex("POIconType")) = "none"
                    End If
                    
                    .Cell(flexcpText, r, .ColIndex("Budgeted")) = format(Val("" & rs("budgeted")), "#,###.00")
                    .Cell(flexcpText, r, .ColIndex("Changes")) = format(Val("" & rs("Changes")), "#,###.00")
                    .Cell(flexcpText, r, .ColIndex("Commited")) = format(Val("" & rs("commited")), "#,###.00")
                    
                    .Cell(flexcpText, r, .ColIndex("IsChangeRequest")) = IIf(mViewIndex = 0, Val("" & rs("IsChangeRequest")), 0)
                    .Cell(flexcpText, r, .ColIndex("BudgetsLocked")) = Val("" & rs("BudgetsLocked"))
                    .Cell(flexcpText, r, .ColIndex("hflocation")) = Val("" & rs("hflocation"))
                    If Val("" & rs("BudgetsLocked")) <> 0 Then .Cell(flexcpForeColor, r, .ColIndex("Budgeted")) = vbGrayText
                    
                    .Cell(flexcpData, r, 0) = DisplayFld
                    .Cell(flexcpText, r, 1) = KeyValue
                    .Cell(flexcpData, r, 1) = KeyFld

                    Select Case True
                        Case Val("" & rs(2)) = Val("" & rs(3)): .Cell(flexcpText, r, .ColIndex("completed")) = msAll
                        Case Val("" & rs(3)) > 0:               .Cell(flexcpText, r, .ColIndex("completed")) = msSome
                        Case Else:                              .Cell(flexcpText, r, .ColIndex("completed")) = msNone
                    End Select

                    .Cell(flexcpText, r, 3) = msNone
                    Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(.Cell(flexcpValue, r, .ColIndex("completed")), .Cell(flexcpValue, r, .ColIndex("selected")), .Cell(flexcpText, r, .ColIndex("IsChangeRequest")), .Cell(flexcpText, r, .ColIndex("POIconType")))

                    .IsSubtotal(r) = True
                    .RowOutlineLevel(r) = Level + 1

                    If Level + 2 <> levels Then
                        Call .GetNode(r).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                    End If

                    Set n = .GetNode(r)
                    n.Expanded = False

                    
                    
                    If ProjectBased Then
                        WhereClause = " AND i.DivisionID = " & HFApp.DivisionID & " and isquote=0 and Community=" & DbQuote(Str, mCommunity)
                    Else
                        If mViews(mViewIndex).Name = "Customer" Then
                            WhereClause = " AND i.DivisionID = " & HFApp.DivisionID & " and AssemblyJob_No=" & DbQuote(Str, mJob) & vbCrLf
                        Else
                            WhereClause = " AND i.DivisionID = " & HFApp.DivisionID & " and Job_No=" & DbQuote(Str, mJob) & vbCrLf
                        End If
                    End If

                    .RowData(r) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Str, KeyValue)

                Else
                    r = .Row
                    Set n = .GetNode(r).AddNode(flexNTLastChild, DisplayValue)
                    
                    .Cell(flexcpText, n.Row, .ColIndex("Budgeted")) = format(Val("" & rs("budgeted")), "#,###.00")
                    .Cell(flexcpText, n.Row, .ColIndex("Changes")) = format(Val("" & rs("Changes")), "#,###.00")
                    .Cell(flexcpText, n.Row, .ColIndex("Commited")) = format(Val("" & rs("commited")), "#,###.00")
                    
                    If IsIn(KeyFld, "PONumber", "PONumberVendor") Then
                        .Cell(flexcpText, n.Row, .ColIndex("POIconType")) = "" & rs("POIconType")
                        If .Cell(flexcpText, n.Row, .ColIndex("POIconType")) = "" Then .Cell(flexcpText, n.Row, .ColIndex("POIconType")) = "none"
                    End If
                    
                    .Cell(flexcpText, n.Row, .ColIndex("IsChangeRequest")) = IIf(mViewIndex = 0, Val("" & rs("IsChangeRequest")), 0)
                    .Cell(flexcpText, n.Row, .ColIndex("BudgetsLocked")) = Val("" & rs("BudgetsLocked"))
                    If Val("" & rs("BudgetsLocked")) <> 0 Then .Cell(flexcpForeColor, n.Row, .ColIndex("Budgeted")) = vbGrayText
                    .Cell(flexcpText, n.Row, .ColIndex("hflocation")) = Val("" & rs("hflocation"))
                    .Cell(flexcpData, n.Row, 0) = DisplayFld
                    .Cell(flexcpText, n.Row, 1) = KeyValue
                    .Cell(flexcpData, n.Row, 1) = KeyFld
                    Select Case True
                        Case rs(2) = rs(3): .Cell(flexcpText, n.Row, .ColIndex("completed")) = msAll
                        Case rs(3) > 0:     .Cell(flexcpText, n.Row, .ColIndex("completed")) = msSome
                        Case Else:          .Cell(flexcpText, n.Row, .ColIndex("completed")) = msNone
                    End Select
                    .Cell(flexcpText, n.Row, .ColIndex("selected")) = .Cell(flexcpText, .Row, .ColIndex("selected"))
                    Set .Cell(flexcpPicture, n.Row, 0) = MultiStateIcon(.Cell(flexcpValue, n.Row, .ColIndex("completed")), .Cell(flexcpValue, n.Row, .ColIndex("selected")), .Cell(flexcpText, n.Row, .ColIndex("IsChangeRequest")), .Cell(flexcpText, n.Row, .ColIndex("POIconType")))


                    If Level + 2 < levels Then
                        Call .GetNode(n.Row).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                    End If

                    n.Expanded = False


                    Select Case KeyFld
                        Case "EstAssemblyID", "EstimateIndex", "POGroup"
                            .RowData(n.Row) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Num, KeyValue)
                        Case Else
                            .RowData(n.Row) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Str, KeyValue)
                    End Select

                    r = n.Row
                End If
                
                
                If IsIn(KeyFld, "PONumberVendor", "PONumber") Then
                    If "" & rs("POStatus") = "Pending" Then
                        .Cell(flexcpForeColor, r, .ColIndex("DisplayField")) = &H80&
                    End If
                End If

                rs.MoveNext
            Wend
            
        End If

        Call Form_Resize
        .Redraw = flexRDBuffered
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "LoadAssemblies", s)
End Sub


Private Sub LoadViews()
    
    Select Case True
        Case IsIn(mMode, fceQuote, fceBudget) And ProjectBased And ProjectPhaseBased
            Call LoadViews_PhaseProjectBudget
            
        Case IsIn(mMode, fceQuote, fceBudget) And ProjectBased
            Call LoadViews_CommunityProjectBudget
            
        Case IsIn(mMode, fceQuote, fceBudget)
            Call LoadViews_JobBudget
            
        Case ProjectBased And ProjectPhaseBased
            Call LoadViews_PhaseProjectPO
            
        Case ProjectBased
            Call LoadViews_CommunityProjectPO
            
        Case Else
            Call LoadViews_JobPO
    End Select
    
    mViewIndex = 0

End Sub

Private Sub LoadViews_CommunityProjectBudget()
    Dim i As Long
'    ReDim mViews(9) As ViewDefs
    ReDim mViews(7) As ViewDefs

    If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
        mViews(i).Name = "Job"
    Else
        mViews(i).Name = "Customer"
    End If
    mViews(i).KeyFlds = "Community,AssemblyJob_No,Customer_no,CORSort,EstAssemblyID"
    mViews(i).DisplayFlds = "CommunityDesc,AssemblyJob_No,CustomerDesc,CORNumber,HFDescription"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "Community,AssemblyJob_No,customer,option,assembly"
    i = i + 1

    mViews(i).Name = "PO Group"
    mViews(i).KeyFlds = "POGroup,POIndex,Job_no,BudgetVendor"
    mViews(i).DisplayFlds = "POGroup,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,Job_no,BudgetVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,purchaseorder,Job_no,vendor"
    i = i + 1

    mViews(i).Name = "PO Index"
    mViews(i).KeyFlds = "POIndex,Job_no,BudgetVendor"
    mViews(i).DisplayFlds = "case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,Job_no,BudgetVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "purchaseorder,Job_no,vendor"
    i = i + 1

    mViews(i).Name = "Cost Code"
    mViews(i).KeyFlds = "JCExtra,JCCostCode,JCCategory,Job_no"
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        mViews(i).DisplayFlds = "JCExtra,JCCostCodeDesc,JCCategoryDesc,Job_no"
    Else
        mViews(i).DisplayFlds = "JCExtra,JCCostCode+' - '+JCCostCodeDesc,JCCategoryDesc,Job_no"
    End If
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,folder,folder"
    i = i + 1

    mViews(i).Name = "Vendor"
    mViews(i).KeyFlds = "BudgetVendor,POIndex,Job_no"
    mViews(i).DisplayFlds = "BudgetVendorName,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,Job_no"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "vendor,purchaseorder,Job_no"
    i = i + 1


    mViews(i).Name = "-"
    i = i + 1

'    mViews(i).Name = "RFQ's"
'    mViews(i).KeyFlds = "RFP"
'    mViews(i).DisplayFlds = "RFPTitle"
'    mViews(i).SortFlds = mViews(i).KeyFlds
'    mViews(i).IconKeys = "purchaseorder"
'    i = i + 1
'
    mViews(i).Name = "Addons"
    i = i + 1
'
'    mViews(i).Name = "-"
'    i = i + 1
    mViews(i).Name = "Contract"
    i = i + 1
End Sub

Private Sub LoadViews_PhaseProjectBudget()
    Dim i As Long
'    ReDim mViews(9) As ViewDefs
    ReDim mViews(7) As ViewDefs
    
    If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
        mViews(i).Name = "Job"
    Else
        mViews(i).Name = "Customer"
    End If
    mViews(i).KeyFlds = "Community,CommunityPhase,AssemblyJob_No,Customer_no,CORSort,EstAssemblyID"
    mViews(i).DisplayFlds = "CommunityDesc,CommunityPhase,AssemblyJob_No,CustomerDesc,CORNumber,HFDescription"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "Community,CommunityPhase,AssemblyJob_No,customer,option,assembly"
    i = i + 1

    mViews(i).Name = "PO Group"
    mViews(i).KeyFlds = "POGroup,POIndex,CommunityPhase,Job_no,BudgetVendor"
    mViews(i).DisplayFlds = "POGroup,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,CommunityPhase,Job_no,BudgetVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,purchaseorder,Job_no,vendor"
    i = i + 1

    mViews(i).Name = "PO Index"
    mViews(i).KeyFlds = "POIndex,CommunityPhase,Job_no,BudgetVendor"
    mViews(i).DisplayFlds = "case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,CommunityPhase,Job_no,BudgetVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "purchaseorder,Job_no,vendor"
    i = i + 1

    mViews(i).Name = "Cost Code"
    mViews(i).KeyFlds = "JCExtra,JCCostCode,JCCategory,CommunityPhase,Job_no"
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        mViews(i).DisplayFlds = "JCExtra,JCCostCodeDesc,JCCategoryDesc,CommunityPhase,Job_no"
    Else
        mViews(i).DisplayFlds = "JCExtra,JCCostCode+' - '+JCCostCodeDesc,JCCategoryDesc,CommunityPhase,Job_no"
    End If
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,folder,folder"
    i = i + 1

    mViews(i).Name = "Vendor"
    mViews(i).KeyFlds = "BudgetVendor,POIndex,CommunityPhase,Job_no"
    mViews(i).DisplayFlds = "BudgetVendorName,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,CommunityPhase,Job_no"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "vendor,purchaseorder,CommunityPhase,Job_no"
    i = i + 1


    mViews(i).Name = "-"
    i = i + 1
'
'    mViews(i).Name = "RFQ's"
'    mViews(i).KeyFlds = "RFP"
'    mViews(i).DisplayFlds = "RFPTitle"
'    mViews(i).SortFlds = mViews(i).KeyFlds
'    mViews(i).IconKeys = "purchaseorder"
'    i = i + 1
'
    mViews(i).Name = "Addons"
    i = i + 1
'
'    mViews(i).Name = "-"
'    i = i + 1
    mViews(i).Name = "Contract"
    i = i + 1
End Sub
Private Sub LoadViews_JobBudget()
    Dim i As Long
'    ReDim mViews(9) As ViewDefs
    ReDim mViews(7) As ViewDefs
    
    If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
        mViews(i).Name = "Job"
    Else
        mViews(i).Name = "Customer"
    End If
    mViews(i).KeyFlds = "Customer_No,CORSort,EstAssemblyID"
    mViews(i).DisplayFlds = "CustomerDesc,CORNumber,HFDescription"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "customer,option,assembly"
    i = i + 1

    mViews(i).Name = "PO Group"
    mViews(i).KeyFlds = "POGroup,POIndex,BudgetVendor"
    mViews(i).DisplayFlds = "POGroup,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,BudgetVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,purchaseorder,vendor"
    i = i + 1

    mViews(i).Name = "PO Index"
    mViews(i).KeyFlds = "POIndex,BudgetVendor"
    mViews(i).DisplayFlds = "case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,BudgetVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "purchaseorder,vendor"
    i = i + 1

    mViews(i).Name = "Cost Code"
    mViews(i).KeyFlds = "JCExtra,JCCostCode,JCCategory"
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        mViews(i).DisplayFlds = "JCExtra,JCCostCodeDesc,JCCategoryDesc"
    Else
        mViews(i).DisplayFlds = "JCExtra,JCCostCode+' - '+JCCostCodeDesc,JCCategoryDesc"
    End If
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,folder,folder"
    i = i + 1

    mViews(i).Name = "Vendor"
    mViews(i).KeyFlds = "BudgetVendor,POIndex"
    mViews(i).DisplayFlds = "BudgetVendorName,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "vendor,purchaseorder"
    i = i + 1


    mViews(i).Name = "-"
    i = i + 1
'
'    mViews(i).Name = "RFQ's"
'    mViews(i).KeyFlds = "RFP"
'    mViews(i).DisplayFlds = "RFPTitle"
'    mViews(i).SortFlds = mViews(i).KeyFlds
'    mViews(i).IconKeys = "purchaseorder"
'    i = i + 1
'
    mViews(i).Name = "Addons"
    i = i + 1
'
'    mViews(i).Name = "-"
'    i = i + 1
    mViews(i).Name = "Contract"
    i = i + 1
End Sub
Private Sub LoadViews_CommunityProjectPO()
    Dim i As Long
'    ReDim mViews(10) As ViewDefs
    ReDim mViews(8) As ViewDefs
    
    mViews(i).Name = "Customer"
    mViews(i).KeyFlds = "Community,AssemblyJob_No,Customer_No,CORSort,EstAssemblyID"
    'mViews(i).DisplayFlds = "CommunityDesc,AssemblyJob_No+' '+CustomerDesc,CustomerDesc,CORNumber,HFDescription"
    mViews(i).DisplayFlds = "CommunityDesc,AssemblyJob_No,CustomerDesc,CORNumber,HFDescription"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "AssemblyJob_No,customer,option,assembly"
    i = i + 1
    
    mViews(i).Name = "PO Group"
    mViews(i).KeyFlds = "POGroup,POIndex,Job_no,PONumberVendor"
    mViews(i).DisplayFlds = "POGroup,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,Job_no,PONumber + ' - ' + POVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,purchaseorder,sendpos"
    i = i + 1

    mViews(i).Name = "PO Index"
    mViews(i).KeyFlds = "POIndex,Job_no,PONumberVendor"
    mViews(i).DisplayFlds = "case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,Job_no,PONumber + ' - ' + POVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "purchaseorder,sendpos"
    i = i + 1

    mViews(i).Name = "Purchase Orders"
    mViews(i).KeyFlds = "PONumber"
    mViews(i).DisplayFlds = "PONumber + CASE PONumber WHEN '' THEN '' ELSE ' -- ' + POVendorName END"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "sendpos"
    i = i + 1
    
    mViews(i).Name = "Vendor"
    mViews(i).KeyFlds = "POVendor,POIndex,PONumber,Job_no"
    mViews(i).DisplayFlds = "POVendorName,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,ponumber,Job_no"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "vendor,sendpos"
    i = i + 1

    mViews(i).Name = "Cost Code"
    mViews(i).KeyFlds = "JCExtra,JCCostCode,JCCategory,Job_no"
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        mViews(i).DisplayFlds = "JCExtra,JCCostCodeDesc,JCCategoryDesc,Job_no"
    Else
        mViews(i).DisplayFlds = "JCExtra,JCCostCode+' - '+JCCostCodeDesc,JCCategoryDesc,Job_no"
    End If
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,folder,folder"
    i = i + 1

    mViews(i).Name = "-"
    i = i + 1
'
'    mViews(i).Name = "RFQ's"
'    mViews(i).KeyFlds = "RFP"
'    mViews(i).DisplayFlds = "RFPTitle"
'    mViews(i).SortFlds = mViews(i).KeyFlds
'    mViews(i).IconKeys = "purchaseorder"
'    i = i + 1
'
    mViews(i).Name = "Addons"
    i = i + 1
'
'    mViews(i).Name = "-"
'    i = i + 1
'
    mViews(i).Name = "Contract"
    i = i + 1

End Sub
Private Sub LoadViews_PhaseProjectPO()
    Dim i As Long
'    ReDim mViews(10) As ViewDefs
    ReDim mViews(8) As ViewDefs
    
    mViews(i).Name = "Customer"
    mViews(i).KeyFlds = "Community,CommunityPhase,AssemblyJob_No,Customer_No,CORSort,EstAssemblyID"
    mViews(i).DisplayFlds = "CommunityDesc,CommunityPhase,AssemblyJob_No+' '+JobDesc,CustomerDesc,CORNumber,HFDescription"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "AssemblyJob_No,customer,option,assembly"
    i = i + 1
    
    mViews(i).Name = "PO Group"
    mViews(i).KeyFlds = "POGroup,POIndex,CommunityPhase,Job_no,PONumberVendor"
    mViews(i).DisplayFlds = "POGroup,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,CommunityPhase,Job_no,PONumber + ' - ' + POVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,purchaseorder,sendpos"
    i = i + 1

    mViews(i).Name = "PO Index"
    mViews(i).KeyFlds = "POIndex,CommunityPhase,Job_no,PONumberVendor"
    mViews(i).DisplayFlds = "case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,CommunityPhase,Job_no,PONumber + ' - ' + POVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "purchaseorder,sendpos"
    i = i + 1

    mViews(i).Name = "Purchase Orders"
    mViews(i).KeyFlds = "PONumber"
    mViews(i).DisplayFlds = "PONumber + CASE PONumber WHEN '' THEN '' ELSE ' -- ' + POVendorName END"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "sendpos"
    i = i + 1
    
    mViews(i).Name = "Vendor"
    mViews(i).KeyFlds = "POVendor,POIndex,PONumber,CommunityPhase,Job_no"
    mViews(i).DisplayFlds = "POVendorName,POIndex,PONUmber + ' - ' + case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,CommunityPhase,Job_no"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "vendor,sendpos"
    i = i + 1

    mViews(i).Name = "Cost Code"
    mViews(i).KeyFlds = "JCExtra,JCCostCode,JCCategory,CommunityPhase,Job_no"
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        mViews(i).DisplayFlds = "JCExtra,JCCostCodeDesc,JCCategoryDesc,CommunityPhase,Job_no"
    Else
        mViews(i).DisplayFlds = "JCExtra,JCCostCode+' - '+JCCostCodeDesc,JCCategoryDesc,CommunityPhase,Job_no"
    End If
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,folder,folder"
    i = i + 1

    mViews(i).Name = "-"
    i = i + 1
'
'    mViews(i).Name = "RFQ's"
'    mViews(i).KeyFlds = "RFP"
'    mViews(i).DisplayFlds = "RFPTitle"
'    mViews(i).SortFlds = mViews(i).KeyFlds
'    mViews(i).IconKeys = "purchaseorder"
'    i = i + 1
'
    mViews(i).Name = "Addons"
    i = i + 1
'
'    mViews(i).Name = "-"
'    i = i + 1
'
    mViews(i).Name = "Contract"
    i = i + 1
End Sub
Private Sub LoadViews_JobPO()
    Dim i As Long
'    ReDim mViews(10) As ViewDefs
    ReDim mViews(8) As ViewDefs
    
    mViews(i).Name = "Customer"
    mViews(i).KeyFlds = "Customer_No,CORSort,EstAssemblyID"
    mViews(i).DisplayFlds = "CustomerDesc,CORNumber,HFDescription"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "customer,option,assembly"
    i = i + 1
    
    mViews(i).Name = "PO Group"
    mViews(i).KeyFlds = "POGroup,POIndex,PONumberVendor"
    mViews(i).DisplayFlds = "POGroup,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,PONumber + ' - ' + POVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,purchaseorder,sendpos"
    i = i + 1

    mViews(i).Name = "PO Index"
    mViews(i).KeyFlds = "POIndex,PONumberVendor"
    mViews(i).DisplayFlds = "case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,PONumber + ' - ' + POVendorName"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "purchaseorder,sendpos"
    i = i + 1

    mViews(i).Name = "Purchase Orders"
    mViews(i).KeyFlds = "PONumber"
    mViews(i).DisplayFlds = "PONumber + CASE PONumber WHEN '' THEN '' ELSE ' -- ' + POVendorName END"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "sendpos"
    i = i + 1
    
    mViews(i).Name = "Vendor"
    mViews(i).KeyFlds = "POVendor,POIndex,PONumber"
    mViews(i).DisplayFlds = "POVendorName,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,PONumber"
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "vendor,sendpos"
    i = i + 1

    mViews(i).Name = "Cost Code"
    mViews(i).KeyFlds = "JCExtra,JCCostCode,JCCategory"
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        mViews(i).DisplayFlds = "JCExtra,JCCostCodeDesc,JCCategoryDesc"
    Else
        mViews(i).DisplayFlds = "JCExtra,JCCostCode+' - '+JCCostCodeDesc,JCCategoryDesc"
    End If
    mViews(i).SortFlds = mViews(i).KeyFlds
    mViews(i).IconKeys = "folder,folder,folder"
    i = i + 1

    mViews(i).Name = "-"
    i = i + 1
'
'    mViews(i).Name = "RFQ's"
'    mViews(i).KeyFlds = "RFP"
'    mViews(i).DisplayFlds = "RFPTitle"
'    mViews(i).SortFlds = mViews(i).KeyFlds
'    mViews(i).IconKeys = "purchaseorder"
'    i = i + 1
'
    mViews(i).Name = "Addons"
    i = i + 1
'
'    mViews(i).Name = "-"
'    i = i + 1
'
    mViews(i).Name = "Contract"
    i = i + 1
End Sub

Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    Dim r           As Long
    Dim c           As Long
    Dim rsel As Long
    Dim csel As Long
    Dim rowBudgetsLocked  As Boolean
    Dim rowBudgeted As Boolean
    Dim rowPOed     As Boolean
    Dim rowCorrection As Boolean
    Dim rate        As Double
    Dim rs          As Recordset
    Dim s           As String
    Dim i           As Long
    
    

    With gItems
    
    For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
    
        If r > 0 Then
    
        rowBudgetsLocked = .TextMatrix(r, .ColIndex("BudgetsLocked")) = "True"
        rowBudgeted = .ValueMatrix(r, .ColIndex("BudgetGenerated")) <> 0
        rowPOed = Trim(.TextMatrix(r, .ColIndex("PONumber"))) <> ""
        rowCorrection = Trim(.TextMatrix(r, .ColIndex("IsCorrectingItem"))) = "True"
        
        
        Select Case .ColKey(Col)
            'case "EstItemID":
            'case "EstAssemblyID":
            Case "Selected":
            'case "Job_No":
            'case "JobDesc":
            Case "JCExtra":
            Case "JCCostCode":
            Case "JCCostCodeDesc":
            'case "OriginalJCCategory":
            Case "JCCategory", "JCCategoryDesc":
                s = ""
                s = s & "SELECT * FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & vbCrLf
                s = s & "dbo.Purch_GetDefaultTaxGroup(" & DbQuote(Str, .TextMatrix(r, .ColIndex("Job_No"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, mCommunity) & vbCrLf
                s = s & "   ," & DbQuote(Str, mCommunityPhase) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("BudgetVendor"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("JCCategory"))) & "," & HFApp.DivisionID & ")"
                Set rs = HFApp.SqlExec(s)
                If Not rs.EOF Then
                
                    If Not (rowBudgeted Or rowBudgetsLocked) Then
                        .TextMatrix(r, .ColIndex("BudgetTaxGroup")) = "" & rs("TaxGroup")
                        .TextMatrix(r, .ColIndex("BudgetJCTaxRate")) = "" & rs("JCRate")
                        .TextMatrix(r, .ColIndex("BudgetNJCTaxRate")) = "" & rs("NJCRate")
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    End If
                    
                    If Not rowPOed And Not rowCorrection Then
                        .TextMatrix(r, .ColIndex("POTaxGroup")) = "" & rs("TaxGroup")
                        .TextMatrix(r, .ColIndex("POJCTaxRate")) = "" & rs("JCRate")
                        .TextMatrix(r, .ColIndex("PONJCTaxRate")) = "" & rs("NJCRate")
                        .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    End If
                    
                End If
                    

                    
                    
                
            'case "HFDescription":
            'case "EstPhase":
            'case "EstItem":
            Case "ItemDesc":
            Case "ItemComments":
            
            Case "BudgetVendor", "BudgetVendorName":
                If Not rowBudgeted Then
                    If GetVendorCost(r, .TextMatrix(r, .ColIndex("BudgetVendor")), rate) Then
                        .TextMatrix(r, .ColIndex("BudgetOverridden")) = "False"
                        .TextMatrix(r, .ColIndex("BudgetRate")) = rate
                        .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                        If Not rowPOed Then
                            .TextMatrix(r, .ColIndex("POOverridden")) = "False"
                            .TextMatrix(r, .ColIndex("PORate")) = .TextMatrix(r, .ColIndex("BudgetRate"))
                            .TextMatrix(r, .ColIndex("POPretax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                            .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                            .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                            .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                            .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                        End If
                    End If
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("POVendor")) = .TextMatrix(r, .ColIndex("BudgetVendor"))
                        .TextMatrix(r, .ColIndex("POVendorName")) = .TextMatrix(r, .ColIndex("BudgetVendorName"))
                    End If
                End If
            Case "POVendor", "POVendorName":
                 If Not rowPOed Then
                    If GetVendorCost(r, .TextMatrix(r, .ColIndex("POVendor")), rate) Then
                        .TextMatrix(r, .ColIndex("PORate")) = rate
                        .TextMatrix(r, .ColIndex("POOverridden")) = "False"
                        .TextMatrix(r, .ColIndex("POPretax")) = Round(.ValueMatrix(r, .ColIndex("POQty")) * .ValueMatrix(r, .ColIndex("PORate")), 2)
                        .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                        If rowBudgetsLocked Then
                            If .TextMatrix(r, .ColIndex("POPretax")) <> .TextMatrix(r, .ColIndex("BudgetPretax")) Then
                                If HFApp.Options.ValueByName("UseVarianceReporting") = "False" Then
                                    .TextMatrix(r, .ColIndex("VarianceJCCategory")) = .TextMatrix(r, .ColIndex("JCCategory"))
                                    .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = .TextMatrix(r, .ColIndex("JCCategoryDesc"))
                                Else
                                    If FPickList.Choose(HFApp.Databases(dbHomefront), "Variance Category", "Select [Category], [Description] from StandardCategories where isvariance=1 and DivisionID =" & HFApp.DivisionID, VarianceCatDef) Then
                                        .TextMatrix(r, .ColIndex("VarianceJCCategory")) = FPickList.SelectedItem("Category")
                                        .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = FPickList.SelectedItem("Description")
                                    End If
                                End If
                            End If
                        End If
                        If Not (rowBudgeted Or rowBudgetsLocked) Then
                            .TextMatrix(r, .ColIndex("BudgetOverridden")) = "False"
                            .TextMatrix(r, .ColIndex("BudgetRate")) = .TextMatrix(r, .ColIndex("PORate"))
                            .TextMatrix(r, .ColIndex("BudgetPretax")) = .TextMatrix(r, .ColIndex("POPretax"))
                            .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                            .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                            .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                            .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                        End If
                    End If
                    If Not (rowBudgeted Or rowBudgetsLocked) Then
                        .TextMatrix(r, .ColIndex("BudgetVendor")) = .TextMatrix(r, .ColIndex("POVendor"))
                        .TextMatrix(r, .ColIndex("BudgetVendorName")) = .TextMatrix(r, .ColIndex("POVendorName"))
                    End If
                End If
                
                
            Case "TakeoffQty", "ConversionFactor":
                If mMode <> fcePO Then
                    If Not rowBudgeted Then
                        .TextMatrix(r, .ColIndex("BudgetQty")) = RoundTo(.ValueMatrix(r, .ColIndex("TakeoffQty")) * (100 + .ValueMatrix(r, .ColIndex("WastePercent"))) / 100 * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
                        .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                        If Not rowPOed Then
                            .TextMatrix(r, .ColIndex("POQty")) = .TextMatrix(r, .ColIndex("BudgetQTY"))
                            .TextMatrix(r, .ColIndex("POPretax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                            .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                            .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                            .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                            .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                        End If
                    End If
                Else
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("POQty")) = RoundTo(.ValueMatrix(r, .ColIndex("TakeoffQty")) * (100 + .ValueMatrix(r, .ColIndex("WastePercent"))) / 100 * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
                        .TextMatrix(r, .ColIndex("POPretax")) = Round(.ValueMatrix(r, .ColIndex("POQty")) * .ValueMatrix(r, .ColIndex("PORate")), 2)
                        .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                        If Not (rowBudgeted Or rowBudgetsLocked) Then
                            .TextMatrix(r, .ColIndex("BudgetQty")) = .TextMatrix(r, .ColIndex("POQTY"))
                            .TextMatrix(r, .ColIndex("BudgetPretax")) = .TextMatrix(r, .ColIndex("POPretax"))
                            .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                            .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                            .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                            .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                        End If
                    End If
                End If
                        
                If Not rowPOed Or Not rowBudgeted Then
                    If rowBudgetsLocked And .TextMatrix(r, .ColIndex("POPretax")) <> .TextMatrix(r, .ColIndex("BudgetPretax")) And Trim(.TextMatrix(r, .ColIndex("VarianceJCCategory"))) = "" Then
                        If HFApp.Options.ValueByName("UseVarianceReporting") = "False" Then
                            .TextMatrix(r, .ColIndex("VarianceJCCategory")) = .TextMatrix(r, .ColIndex("JCCategory"))
                            .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = .TextMatrix(r, .ColIndex("JCCategoryDesc"))
                        Else
                            If FPickList.Choose(HFApp.Databases(dbHomefront), "Variance Category", "Select [Category], [Description] from StandardCategories where isvariance=1 and DivisionID =" & HFApp.DivisionID, VarianceCatDef) Then
                                .TextMatrix(r, .ColIndex("VarianceJCCategory")) = FPickList.SelectedItem("Category")
                                .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = FPickList.SelectedItem("Description")
                            End If
                        End If
                    End If
                End If
            
            Case "BudgetQty":
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("TakeoffQty")) = .ValueMatrix(r, .ColIndex("BudgetQty")) / .ValueMatrix(r, .ColIndex("ConversionFactor"))
                    .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("POQty")) = .TextMatrix(r, .ColIndex("BudgetQTY"))
                        .TextMatrix(r, .ColIndex("POPretax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                        .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                        .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    End If
                End If
                
            Case "POQty":
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("TakeoffQty")) = .ValueMatrix(r, .ColIndex("POQty")) / .ValueMatrix(r, .ColIndex("ConversionFactor"))
                    .TextMatrix(r, .ColIndex("POPretax")) = Round(.ValueMatrix(r, .ColIndex("POQty")) * .ValueMatrix(r, .ColIndex("PORate")), 2)
                    .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    If rowBudgetsLocked And .TextMatrix(r, .ColIndex("POPretax")) <> .TextMatrix(r, .ColIndex("BudgetPretax")) And Trim(.TextMatrix(r, .ColIndex("VarianceJCCategory"))) = "" Then
                        If HFApp.Options.ValueByName("UseVarianceReporting") = "False" Then
                            .TextMatrix(r, .ColIndex("VarianceJCCategory")) = .TextMatrix(r, .ColIndex("JCCategory"))
                            .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = .TextMatrix(r, .ColIndex("JCCategoryDesc"))
                        Else
                            If FPickList.Choose(HFApp.Databases(dbHomefront), "Variance Category", "Select [Category], [Description] from StandardCategories where isvariance=1 and DivisionID =" & HFApp.DivisionID, VarianceCatDef) Then
                                .TextMatrix(r, .ColIndex("VarianceJCCategory")) = FPickList.SelectedItem("Category")
                                .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = FPickList.SelectedItem("Description")
                            End If
                        End If
                    End If
                    
                    
                    If Not (rowBudgeted Or rowBudgetsLocked) Then
                        .TextMatrix(r, .ColIndex("BudgetQty")) = .TextMatrix(r, .ColIndex("POQty"))
                        .TextMatrix(r, .ColIndex("BudgetPretax")) = .TextMatrix(r, .ColIndex("POPretax"))
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    End If
                End If
                
            Case "OrderUOM":
            
            Case "BudgetRate":
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                    .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                        .TextMatrix(r, .ColIndex("PORate")) = .TextMatrix(r, .ColIndex("BudgetRate"))
                        .TextMatrix(r, .ColIndex("POPretax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                        .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                        .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    End If
                End If
                
            Case "PORate":
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                    .TextMatrix(r, .ColIndex("POPretax")) = Round(.ValueMatrix(r, .ColIndex("POQty")) * .ValueMatrix(r, .ColIndex("PORate")), 2)
                    .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    If Not (rowBudgeted Or rowBudgetsLocked) Then
                        .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                        .TextMatrix(r, .ColIndex("BudgetRate")) = .TextMatrix(r, .ColIndex("PORate"))
                        .TextMatrix(r, .ColIndex("BudgetPretax")) = .TextMatrix(r, .ColIndex("POPretax"))
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    End If
                    
                    If rowBudgetsLocked And .TextMatrix(r, .ColIndex("POPretax")) <> .TextMatrix(r, .ColIndex("BudgetPretax")) And Trim(.TextMatrix(r, .ColIndex("VarianceJCCategory"))) = "" Then
                        If HFApp.Options.ValueByName("UseVarianceReporting") = "False" Then
                            .TextMatrix(r, .ColIndex("VarianceJCCategory")) = .TextMatrix(r, .ColIndex("JCCategory"))
                            .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = .TextMatrix(r, .ColIndex("JCCategoryDesc"))
                        Else
                            If FPickList.Choose(HFApp.Databases(dbHomefront), "Variance Category", "Select [Category], [Description] from StandardCategories where isvariance=1 and DivisionID =" & HFApp.DivisionID, VarianceCatDef) Then
                                .TextMatrix(r, .ColIndex("VarianceJCCategory")) = FPickList.SelectedItem("Category")
                                .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = FPickList.SelectedItem("Description")
                            End If
                        End If
                    End If
                End If
            
            Case "BudgetPretax":
                If Not rowBudgeted Then
                    If .ValueMatrix(r, .ColIndex("BudgetQty")) = 0 Then .TextMatrix(r, .ColIndex("BudgetQty")) = 1
                    .TextMatrix(r, .ColIndex("BudgetRate")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) / .ValueMatrix(r, .ColIndex("BudgetQty")), 4)
                    .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                        .TextMatrix(r, .ColIndex("POQty")) = .TextMatrix(r, .ColIndex("BudgetQty"))
                        .TextMatrix(r, .ColIndex("PORate")) = .TextMatrix(r, .ColIndex("BudgetRate"))
                        .TextMatrix(r, .ColIndex("POPreTax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                        .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                        .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    End If
                End If
                
            Case "POPretax":
                If Not rowPOed Then
                    If .ValueMatrix(r, .ColIndex("POQty")) = 0 Then .TextMatrix(r, .ColIndex("POQty")) = 1
                    .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                    .TextMatrix(r, .ColIndex("PORate")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) / .ValueMatrix(r, .ColIndex("POQty")), 4)
                    .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    If rowBudgetsLocked Then
                        If .TextMatrix(r, .ColIndex("POPretax")) <> .TextMatrix(r, .ColIndex("BudgetPretax")) Then
                            If HFApp.Options.ValueByName("UseVarianceReporting") = "False" Then
                                .TextMatrix(r, .ColIndex("VarianceJCCategory")) = .TextMatrix(r, .ColIndex("JCCategory"))
                                .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = .TextMatrix(r, .ColIndex("JCCategoryDesc"))
                            Else
                                If FPickList.Choose(HFApp.Databases(dbHomefront), "Variance Category", "Select [Category], [Description] from StandardCategories where isvariance=1 and DivisionID =" & HFApp.DivisionID, VarianceCatDef) Then
                                    .TextMatrix(r, .ColIndex("VarianceJCCategory")) = FPickList.SelectedItem("Category")
                                    .TextMatrix(r, .ColIndex("VarianceJCCategoryDesc")) = FPickList.SelectedItem("Description")
                                End If
                            End If
                        End If
                    End If
                    
                    If Not (rowBudgeted Or rowBudgetsLocked) Then
                        .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                        .TextMatrix(r, .ColIndex("BudgetQty")) = .TextMatrix(r, .ColIndex("POQty"))
                        .TextMatrix(r, .ColIndex("BudgetRate")) = .TextMatrix(r, .ColIndex("PORate"))
                        .TextMatrix(r, .ColIndex("BudgetPreTax")) = .TextMatrix(r, .ColIndex("POPretax"))
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    End If
                End If
            
            Case "BudgetTaxGroup":
                If Not rowBudgeted Then
                    s = "SELECT JCRate,NJCRate FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("BudgetTaxGroup")))
                    Set rs = HFApp.SqlExec(s)
                    If rs.EOF Then
                        .TextMatrix(r, .ColIndex("BudgetJCTaxRate")) = 0
                        .TextMatrix(r, .ColIndex("BudgetNJCTaxRate")) = 0
                    Else
                        .TextMatrix(r, .ColIndex("BudgetJCTaxRate")) = Val("" & rs("JCRate"))
                        .TextMatrix(r, .ColIndex("BudgetNJCTaxRate")) = Val("" & rs("NJCRate"))
                    End If
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    If Not rowPOed Then
                        .Cell(flexcpText, r, .ColIndex("POTaxGroup")) = .Cell(flexcpText, r, .ColIndex("BudgetTaxGroup"))
                        .Cell(flexcpText, r, .ColIndex("POJCTaxRate")) = .Cell(flexcpText, r, .ColIndex("BudgetJCTaxRate"))
                        .Cell(flexcpText, r, .ColIndex("PONJCTaxRate")) = .Cell(flexcpText, r, .ColIndex("BudgetNJCTaxRate"))
                        .Cell(flexcpText, r, .ColIndex("POJCTax")) = .Cell(flexcpText, r, .ColIndex("BudgetJCTax"))
                        .Cell(flexcpText, r, .ColIndex("PONJCTax")) = .Cell(flexcpText, r, .ColIndex("BudgetNJCTax"))
                        .Cell(flexcpText, r, .ColIndex("POTax")) = .Cell(flexcpText, r, .ColIndex("BudgetTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    End If
                End If
                
            Case "POTaxGroup":
                If Not rowPOed Then
                    s = "SELECT JCRate,NJCRate FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("POTaxGroup")))
                    Set rs = HFApp.SqlExec(s)
                    If rs.EOF Then
                        .TextMatrix(r, .ColIndex("POJCTaxRate")) = 0
                        .TextMatrix(r, .ColIndex("PONJCTaxRate")) = 0
                    Else
                        .TextMatrix(r, .ColIndex("POJCTaxRate")) = Val("" & rs("JCRate"))
                        .TextMatrix(r, .ColIndex("PONJCTaxRate")) = Val("" & rs("NJCRate"))
                    End If
                    .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    If Not (rowBudgeted Or rowBudgetsLocked) Then
                        .Cell(flexcpText, r, .ColIndex("BudgetTaxGroup")) = .Cell(flexcpText, r, .ColIndex("POTaxGroup"))
                        .Cell(flexcpText, r, .ColIndex("BudgetJCTaxRate")) = .Cell(flexcpText, r, .ColIndex("POJCTaxRate"))
                        .Cell(flexcpText, r, .ColIndex("BudgetNJCTaxRate")) = .Cell(flexcpText, r, .ColIndex("PONJCTaxRate"))
                        .Cell(flexcpText, r, .ColIndex("BudgetJCTax")) = .Cell(flexcpText, r, .ColIndex("POJCTax"))
                        .Cell(flexcpText, r, .ColIndex("BudgetNJCTax")) = .Cell(flexcpText, r, .ColIndex("PONJCTax"))
                        .Cell(flexcpText, r, .ColIndex("BudgetTax")) = .Cell(flexcpText, r, .ColIndex("POTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    End If
                End If
            
            Case "BudgetJCTax":
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                        .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    End If
                End If
                
            
            Case "POJCTax":
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    If Not (rowBudgeted Or rowBudgetsLocked) Then
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    End If
                End If
                
            Case "BudgetNJCTax":
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                        .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    End If
                End If
                
            Case "PONJCTax":
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("POTotal")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("POPreTax"))
                    If Not (rowBudgeted Or rowBudgetsLocked) Then
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                        .TextMatrix(r, .ColIndex("BudgetTotal")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetPreTax"))
                    End If
                End If
                
            Case "POIndex":
            Case "ExcludeFromPO":
        End Select
        
        
        Select Case .ColKey(Col)
            Case "BudgetVendor", "BudgetVendorName", "POVendor", "POVendorName", "TakeoffQty", "BudgetQty", "POQty", "BudgetRate", "PORate", "BudgetPretax", "POPretax", "BudgetTaxGroup", "POTaxGroup", "BudgetJCTax", "POJCTax", "BudgetNJCTax", "PONJCTax":
        End Select
        
        
        If .ColKey(Col) <> "Selected" Then
            Dirty = True
            For i = r To .RowSel
                If .RowData(i) <> "NEW" Then .RowData(i) = "DIRTY"
            Next
        End If
        
        End If
    Next
    End With
    
    
    Select Case gItems.ColKey(Col)
    Case "POQty", "BudgetQty", "BudgetRate", "PORate", "TakeoffQty", "BudgetVendor", "BudgetVendorName", "POVendor", "POVendorName":
    On Error Resume Next
        i = gItems.TopRow
        r = gItems.Row
        c = gItems.Col
        rsel = gItems.RowSel
        csel = gItems.ColSel
        
        Call GroupGrid
        gItems.Row = r
        gItems.Col = c
        gItems.RowSel = rsel
        gItems.ColSel = csel
        gItems.TopRow = i
        
    End Select
    
    
    'Call GroupGrid
End Sub








Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
Static inHere As Boolean
If inHere Or mFunnyFlag Then Exit Sub
inHere = True
    gItems.ColSel = gItems.Col
inHere = False
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rowBudgeted As Boolean
    Dim rowBudgetsLocked As Boolean
    Dim rowPOed As Boolean
    Dim rowEither As Boolean
    Dim qtyLocked As Boolean
    
    
    If ReadOnly Then
        Cancel = True
        Exit Sub
    End If
    With gItems
        .ComboList = ""
        .EditMaxLength = 0
        .AutoSearch = flexSearchNone
        
        If Row >= .Rows Then Exit Sub
        If .IsSubtotal(Row) Then
            Cancel = True
            Exit Sub
        End If
        
        rowBudgetsLocked = .TextMatrix(Row, .ColIndex("BudgetsLocked")) = "True"
        rowBudgeted = .ValueMatrix(Row, .ColIndex("BudgetGenerated")) <> 0
        rowPOed = Trim(.TextMatrix(Row, .ColIndex("PONumber"))) <> "" Or (HFApp.UserPermission("IssuePOs") = False And mMode = fcePO)
        rowEither = rowBudgeted Or rowPOed Or (rowBudgetsLocked And .TextMatrix(Row, .ColIndex("BudgetDeleted")) = "False")
        
        'prevent changes to qty/rate/amount when budgets are finalized and PostPOQtyToAccounting
        qtyLocked = mMode = fcePO And rowBudgetsLocked And PostPOQtyToAccounting And .TextMatrix(Row, .ColIndex("BudgetDeleted")) = "False"
        
        'prevent changes to cancelled orig and reversal rows
        
        If IsIn(.TextMatrix(Row, .ColIndex("RowType")), "Original - Cancelled", "Original - Reversal") Then
            Cancel = True
            Exit Sub
        End If
        
        If ((.ColKey(Col) <> "Selected" And rowBudgetsLocked And mMode <> fcePO) Or (rowBudgeted And mMode <> fcePO) Or (rowPOed And mMode = fcePO)) And .ColKey(Col) <> "ExcludeFromPO" Then
            Cancel = True
            Exit Sub
        End If
        
        Select Case .ColKey(Col)
            
            'qtylock was added for PostPOQtyToAccounting July 2024, MK
            Case "PORate":             Cancel = rowPOed Or qtyLocked
            Case "POPretax":           Cancel = rowPOed Or qtyLocked
            Case "POJCTax":            Cancel = rowPOed Or qtyLocked:             Cancel = .TextMatrix(Row, .ColIndex("POTaxGroup")) = ""
            Case "PONJCTax":           Cancel = rowPOed Or qtyLocked:             Cancel = .TextMatrix(Row, .ColIndex("POTaxGroup")) = ""
            Case "POQty":              Cancel = rowPOed Or qtyLocked
            Case "POTaxGroup":         Cancel = rowPOed Or qtyLocked:             .ComboList = "|..."
            
            
            Case "Assembly":            Cancel = rowEither:        .ComboList = "..."
            Case "AssemblyDescription": Cancel = rowEither:        .ComboList = "..."
            Case "Selected":           Cancel = .TextMatrix(Row, .ColIndex("EstItemID")) = "" Or .Cell(flexcpChecked, Row, Col) = flexNoCheckbox
            Case "BudgetApproved":     Cancel = .TextMatrix(Row, .ColIndex("EstItemID")) = ""
            Case "JCExtra":            Cancel = rowEither:         .ComboList = IIf(HFApp.Options(Use_Timberline), "|...", ""):   .EditMaxLength = 10
            Case "JCCostCode":         Cancel = rowEither:         .ComboList = "|..."
            Case "JCCostCodeDesc":     Cancel = rowEither:         .ComboList = "..."
            Case "JCCategory":         Cancel = rowEither:         .ComboList = "|..."
            Case "JCCategoryDesc":     Cancel = rowEither:         .ComboList = "..."
            Case "VarianceJCCategory":         Cancel = rowPOed:         .ComboList = "|..."
            Case "VarianceJCCategoryDesc":     Cancel = rowPOed:         .ComboList = "..."
            Case "ItemDesc":                                       .ComboList = "|...":           .EditMaxLength = 200
            Case "ItemComments":                                   .ComboList = "|...":
            Case "BudgetVendor":       Cancel = rowBudgeted:       .ComboList = "|..."
            Case "POVendor":           Cancel = rowPOed:           .ComboList = "|..."
            Case "BudgetVendorName":   Cancel = rowBudgeted:       .ComboList = "..."
            Case "POVendorName":       Cancel = rowPOed:           .ComboList = "..."
            Case "Job_No", "JobDesc":  Cancel = rowEither:         .ComboList = "..."
            Case "TakeoffQty":         Cancel = (rowBudgeted And mMode <> fcePO) Or (rowPOed And mMode = fcePO)
            Case "ConversionFactor":   Cancel = (rowBudgeted And mMode <> fcePO) Or (rowPOed And mMode = fcePO)
            Case "BudgetQty":          Cancel = rowBudgeted
            Case "Unit":                                                                          .EditMaxLength = 10
            Case "Formula":                                        .ComboList = "|..."
            Case "BudgetRate":         Cancel = rowBudgeted
            Case "OrderUOM":                                       .ComboList = "|...":           .EditMaxLength = 10
            Case "BudgetPretax":       Cancel = rowBudgeted
            Case "BudgetTaxGroup":     Cancel = rowBudgeted:       .ComboList = "|..."
            Case "BudgetJCTax":        Cancel = rowBudgeted:        Cancel = .TextMatrix(Row, .ColIndex("BudgetTaxGroup")) = ""
            Case "BudgetNJCTax":       Cancel = rowBudgeted:        Cancel = .TextMatrix(Row, .ColIndex("BudgetTaxGroup")) = ""
            Case "POIndex":            Cancel = rowPOed:           .ComboList = "|..."
            Case "POIndexDescription": Cancel = rowEither:         .ComboList = "..."
            Case "ExcludeFromPO":
            Case "SalesQty":

            Case "Location":           .ComboList = GetComboList("Location"):                     .EditMaxLength = 50
            Case Else:                 Cancel = True
        End Select
        
        If Left(.ColKey(Col), 3) = "WBS" Then
            .ComboList = GetComboList(.ColKey(Col))
            .EditMaxLength = 50
            Cancel = False
        End If
        
        .AutoSearch = IIf(Cancel, flexSearchFromCursor, flexSearchNone)
        
    End With
End Sub



Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Dim i As Long
    Dim r As Long
    Dim s As String
    
    Dim grpd As Boolean
    Dim fixd As Boolean
    
    With gItems
    If Button = vbRightButton Then
        If .MouseRow = 0 Then
            Cancel = True
            
            
            FMain.mnuGrid2Sub(mcGRID_GROUP).checked = .ColData(.MouseCol) = "GROUPED"
            grpd = FMain.mnuGrid2Sub(mcGRID_GROUP).checked
            fixd = IsIn(.ColKey(.MouseCol), "Selected", "WarningMessages")
            Call FMain.ShowColumnMenu(gItems, Not (grpd Or fixd), , , Not fixd)
            
            
            'this is a wbs column save description in case they changed it.
            If .MouseCol > -1 Then
                If Left(.ColKey(.MouseCol), 3) = "WBS" Then
                    i = Val(Mid(.ColKey(.MouseCol), 4))
                    s = "update tbljobs set wbsdesc" & format(i, "00") & "=" & DbQuote(Str, .TextMatrix(0, .MouseCol)) & " where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mJob)
                    Call HFApp.SqlExec(s)
                End If
            End If
            
        Else
            If .Row < 1 And .MouseRow > 1 Then .Row = .MouseRow
            If .Row > 0 Then
            Set MouseCtrl = gItems
            MouseCol = .MouseCol
                                  
            FMain.mnuEstimateItemsGridSub(mcITEM_UPDATEPRICELIST).Enabled = HFApp.UserPermission("SetupVendorPricing")
            FMain.mnuEstimateItemsGridSub(mcITEM_CANCELBUDGETS).Enabled = .TextMatrix(.Row, .ColIndex("BudgetGenerated")) = "True" And .TextMatrix(.Row, .ColIndex("BudgetsLocked")) = "False" And mMode = fceBudget
            
            If mMode = fcePO Then
                If HFApp.UserPermission("IssuePOs") = False Then
                    FMain.mnuEstimateItemsGridSub(mcITEM_SPLITITEMS).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_COPYITEMS).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_REMOVEITEMS).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_MODIFYQTY).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_SAVEONETIMETODB).Enabled = False
                Else
                    FMain.mnuEstimateItemsGridSub(mcITEM_SPLITITEMS).Enabled = .TextMatrix(.Row, .ColIndex("PONumber")) = "" And Not IsIn(.TextMatrix(.Row, .ColIndex("RowType")), "Original - Cancelled", "Original - Reversal", "Original - Correction", "")
                    FMain.mnuEstimateItemsGridSub(mcITEM_COPYITEMS).Enabled = Not IsIn(.TextMatrix(.Row, .ColIndex("RowType")), "Original - Cancelled", "Original - Reversal", "Original - Correction", "")
                    FMain.mnuEstimateItemsGridSub(mcITEM_REMOVEITEMS).Enabled = .TextMatrix(.Row, .ColIndex("PONumber")) = "" And Not IsIn(.TextMatrix(.Row, .ColIndex("RowType")), "Original - Cancelled", "Original - Reversal", "")
                    
                    
                    FMain.mnuEstimateItemsGridSub(mcITEM_MODIFYQTY).Enabled = .ValueMatrix(.Row, .ColIndex("EstItemID")) <> 0 And .TextMatrix(.Row, .ColIndex("ReversingItemID")) = "" And .TextMatrix(.Row, .ColIndex("PONumber")) = "" And _
                                                                              .TextMatrix(.Row, .ColIndex("BudgetsLocked")) = "True" And .TextMatrix(.Row, .ColIndex("BudgetDeleted")) <> "True" And mMode = fcePO And PostPOQtyToAccounting
                    
                    'can do this only if item is onetime
                    FMain.mnuEstimateItemsGridSub(mcITEM_SAVEONETIMETODB).Enabled = .TextMatrix(.Row, .ColIndex("EstPhase")) = ""
                    
                    FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = Trim(.TextMatrix(.Row, .ColIndex("PONumber"))) = "" _
                                                                                  And Not IsIn(.TextMatrix(.Row, .ColIndex("RowType")), "Original - Cancelled", "Original - Reversal", "Original - Correction", "") _
                                                                                  And (.TextMatrix(.Row, .ColIndex("BudgetDeleted")) = "True" Or (.TextMatrix(.Row, .ColIndex("BudgetGenerated")) = "False" And _
                                                                                                                                                   .TextMatrix(.Row, .ColIndex("BudgetsLocked")) = "False") _
                                                                                       )
                End If
            Else
                FMain.mnuEstimateItemsGridSub(mcITEM_MODIFYQTY).Enabled = False
                If HFApp.UserPermission("IssueBudgets") = False Then
                    FMain.mnuEstimateItemsGridSub(mcITEM_SPLITITEMS).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_COPYITEMS).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_REMOVEITEMS).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = False
                    FMain.mnuEstimateItemsGridSub(mcITEM_SAVEONETIMETODB).Enabled = False
                Else
                    FMain.mnuEstimateItemsGridSub(mcITEM_SPLITITEMS).Enabled = .TextMatrix(.Row, .ColIndex("BudgetsLocked")) = "False" And .TextMatrix(.Row, .ColIndex("BudgetGenerated")) = "False"
                    FMain.mnuEstimateItemsGridSub(mcITEM_COPYITEMS).Enabled = .TextMatrix(.Row, .ColIndex("BudgetsLocked")) = "False"
                    FMain.mnuEstimateItemsGridSub(mcITEM_REMOVEITEMS).Enabled = .TextMatrix(.Row, .ColIndex("BudgetsLocked")) = "False" And .TextMatrix(.Row, .ColIndex("BudgetGenerated")) = "False"
                    
                    'can do this only if item is onetime
                    FMain.mnuEstimateItemsGridSub(mcITEM_SAVEONETIMETODB).Enabled = .TextMatrix(.Row, .ColIndex("EstPhase")) = ""
                    
                    'can do this if budgeted not generated and po not generated
                    FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = .TextMatrix(.Row, .ColIndex("BudgetGenerated")) = "False" _
                                                                        And Trim(.TextMatrix(.Row, .ColIndex("PONumber"))) = "" _
                                                                        And (.TextMatrix(.Row, .ColIndex("BudgetsLocked")) = "False")
                End If
            End If
            
            
        
            If ReadOnly Then
                FMain.mnuEstimateItemsGridSub(mcITEM_SPLITITEMS).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_COPYITEMS).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_REMOVEITEMS).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_MODIFYQTY).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = False
                FMain.mnuEstimateItemsGridSub(mcITEM_CANCELBUDGETS).Enabled = False
            End If
            
            PopupMenu FMain.mnuEstimateItemsGrid
            End If
        End If
    End If
    End With
End Sub






Private Sub gItems_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Dim fixd As Boolean
    
    With gItems
        fixd = IsIn(.ColKey(Col), "Selected", "WarningMessages")
        If fixd Or Col < gItems.FrozenCols Then
            Position = Col
        Else
            If Position < gItems.FrozenCols Then
                Position = gItems.FrozenCols
            End If
        End If
    End With

End Sub



Private Sub gItems_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case gItems.ColKey(Col)
        Case "WarningMessages", "Selected"
            Cancel = True
    End Select
End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo eh
    Dim s As String
    Dim i As Long
    Dim OldVendorCSV As String
    Dim NewVendor As String
    Dim NewVendorName As String
    Dim b As Boolean
    Dim rCount As Long
    Dim r As Long
    
    With gItems
    
        Select Case .ColKey(Col)
            
            Case "Assembly", "AssemblyDescription"
                s = ""
                If ProjectBased Then
                    s = s & "select j.job_no Job,j.description JobDesc,c.customer_no Customer,c.description CustomerDesc,a.EstAssemblyID,a.Assembly,a.hfDescription AssemblyDesc" & vbCrLf
                    s = s & "from estimateassemblies a" & vbCrLf
                    s = s & "left outer join tbljobs j on a.job=j.job_no" & vbCrLf
                    s = s & "left outer join tblcustomers c on a.customer_no=c.customer_no" & vbCrLf
                    s = s & "where j.community=" & DbQuote(Str, mCommunity)
                Else
                    s = s & "select Assembly,hfdescription AssemblyDesc,estassemblyid" & vbCrLf
                    s = s & "from estimateassemblies" & vbCrLf
                    s = s & "where job=" & DbQuote(Str, mJob)
                End If
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Assembly", s, , , , , "estassemblyid") Then
                    .Cell(flexcpText, Row, .ColIndex("estassemblyid"), .RowSel, .ColIndex("estassemblyid")) = FPickList.SelectedItem("estassemblyid")
                    .Cell(flexcpText, Row, .ColIndex("assembly"), .RowSel, .ColIndex("assembly")) = FPickList.SelectedItem("Assembly")
                    .Cell(flexcpText, Row, .ColIndex("assemblydescription"), .RowSel, .ColIndex("assemblydescription")) = FPickList.SelectedItem("AssemblyDesc")
                    
                    If ProjectBased Then
                        .Cell(flexcpText, Row, .ColIndex("Job_no"), .RowSel, .ColIndex("Job_no")) = FPickList.SelectedItem("Job")
                        .Cell(flexcpText, Row, .ColIndex("JobDesc"), .RowSel, .ColIndex("JobDesc")) = FPickList.SelectedItem("JobDesc")
                    End If
                    
                End If
            
            Case "Formula"
                s = .Text
                If FFormulaEditor.EditFormula("", s) Then
                    .Text = s
                End If
        
            Case "JCExtra"
                If HFApp.Options(AccountingSystem) = asTimberline Then
                    s = "select Extra,Description from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No"))))
                    If FPickList.Choose(HFApp.Databases(dbAccounting), "Extra", s, gItems) Then
                        For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                            Call gItems_BeforeEdit(Row, Col, b)
                            If Not b Then .Cell(flexcpText, Row, .ColIndex("JCExtra")) = FPickList.SelectedItem("Extra")
                        Next
                    End If
                End If
            Case "JCCostCode"
                s = "SELECT CostCode,Description FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                            .Cell(flexcpText, Row, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                        End If
                    Next
                End If
                
            Case "JCCostCodeDesc"
                s = "SELECT Description,CostCode FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                            .Cell(flexcpText, Row, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                        End If
                    Next
                End If
            
            Case "JCCategory", "JCCategoryDesc"
                If .ColKey(Col) = "JCCategory" Then
                    s = "SELECT Category,Description FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                Else
                    s = "SELECT Description,Category FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                End If
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems, , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Category", "")) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                            .Cell(flexcpText, Row, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                            
                            'if you updated a reversing or correcting entry then also update its pairing
                            i = -1
                            If .TextMatrix(Row, .ColIndex("IsReversingItem")) = "true" Then i = GetCorrectingEntryRowNum(, Row)
                            If .TextMatrix(Row, .ColIndex("IsCorrectingItem")) = "true" Then i = GetReversingEntryRowNum(, Row)
                            If i > 0 Then
                                .TextMatrix(i, .ColIndex("JCCategory")) = .TextMatrix(Row, .ColIndex("JCCategory"))
                                .TextMatrix(i, .ColIndex("JCCategoryDesc")) = .TextMatrix(Row, .ColIndex("JCCategoryDesc"))
                                If .RowData(i) <> "New" Then .RowData(i) = "Dirty"
                            End If
                            
                        End If
                    Next
                End If
                
               
                
                
                
            Case "VarianceJCCategory", "VarianceJCCategoryDesc"
                If .ColKey(Col) = "VarianceJCCategory" Then
                    s = "SELECT Category,Description FROM StandardCategories where isvariance=1 and DivisionID = " & HFApp.DivisionID
                Else
                    s = "SELECT Description,Category FROM StandardCategories where isvariance=1 and DivisionID = " & HFApp.DivisionID
                End If
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, .ColIndex("VarianceJCCategory")) = FPickList.SelectedItem("Category")
                            .Cell(flexcpText, Row, .ColIndex("VarianceJCCategoryDesc")) = FPickList.SelectedItem("Description")
                        End If
                    Next
                End If
                
                
                
            Case "Job_no"
                s = "SELECT Job_No Job,Description,Municipal_Address Address FROM tblJobs where DivisionID = " & HFApp.DivisionID & " and inactive = 0 and isquote = 0"
                If ProjectBased Then s = s & " and community=" & DbQuote(Str, mCommunity)
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Job", s, gItems, , , , , False) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, .ColIndex("Job_no")) = FPickList.SelectedItem("Job")
                            .Cell(flexcpText, Row, .ColIndex("JobDesc")) = FPickList.SelectedItem("Description")
                        End If
                    Next
                End If
                
            Case "JobDesc"
                s = "SELECT Job_No Job,Description,Municipal_Address Address FROM tblJobs where DivisionID = " & HFApp.DivisionID & " and inactive = 0 and isquote = 0"
                If ProjectBased Then s = s & " and community=" & DbQuote(Str, mCommunity)
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Job", s, gItems, , , , , False) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, .ColIndex("Job_no")) = FPickList.SelectedItem("Job")
                            .Cell(flexcpText, Row, .ColIndex("JobDesc")) = FPickList.SelectedItem("Description")
                        End If
                    Next
                End If
            
            
            Case "OrderUOM"
                s = "SELECT DISTINCT OrderUOM Unit FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Units", s, gItems) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, .ColIndex("OrderUOM")) = FPickList.SelectedItem("Unit")
                        End If
                    Next
                End If
            
            Case "ItemComments"
                s = gItems.Text
                If FComments.Edit(s, gItems) Then
                    gItems.Text = s
                End If
                
            Case "ItemDesc":
                s = gItems.Text
                If FComments.Edit(s, gItems, , "Description", 200) Then
                    gItems.Text = s
                End If
                
            Case "BudgetVendor", "BudgetVendorName"
                s = ""
                s = s & "SELECT DISTINCT" & vbCrLf
                s = s & "       case when isnull(c.vendor,'')<>'' then 'Approved' else '' end Status" & vbCrLf
                s = s & "      ,v.vendorgroupid Trade" & vbCrLf
                s = s & "      ,v.Vendor_Name Company" & vbCrLf
                s = s & "      ,v.Vendor_ID Vendor" & vbCrLf
                s = s & "      ,v.City" & vbCrLf
                s = s & "      ,v.Phone " & vbCrLf
                s = s & "  FROM tblVendors v" & vbCrLf
                s = s & "       LEFT OUTER JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor and v.DivisionID = c.DivisionID" & vbCrLf
                s = s & "                                          AND (ISNULL(Community,'')='' OR Community=" & DbQuote(Str, mCommunity) & ")" & vbCrLf
                s = s & "                                          AND (ISNULL(Assembly,'')='' OR Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ")" & vbCrLf
                s = s & "                                          AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & vbCrLf
                s = s & "                                          AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & ")" & vbCrLf
                s = s & " WHERE v.inactive=0 and v.DivisionID = " & HFApp.DivisionID & vbCrLf
                If .ColKey(Col) = "BudgetVendor" Then
                    s = s & "order by 1 desc,4" & vbCrLf
                Else
                    s = s & "order by 1 desc,3" & vbCrLf
                End If
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Vendor", "")) Then
                
                    NewVendor = FPickList.SelectedItem("Vendor")
                    NewVendorName = FPickList.SelectedItem("Company")
                    
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(i, Col, b)
                        If Not b Then
                            OldVendorCSV = OldVendorCSV & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetVendor")))
                            .Cell(flexcpText, i, .ColIndex("BudgetVendor"), i, .ColIndex("BudgetVendor")) = NewVendor
                            .Cell(flexcpText, i, .ColIndex("BudgetVendorName"), i, .ColIndex("BudgetVendorName")) = NewVendorName
                        End If
                    Next
                    OldVendorCSV = Mid(OldVendorCSV, 2)
                    Call UpdateScheduleVendor(.TextMatrix(.Row, .ColIndex("Job_No")), OldVendorCSV, NewVendor)
                End If
            
            Case "POVendor", "POVendorName"
                s = ""
                s = s & "SELECT DISTINCT" & vbCrLf
                s = s & "       case when isnull(c.vendor,'')<>'' then 'Approved' else '' end Status" & vbCrLf
                s = s & "      ,v.vendorgroupid Trade" & vbCrLf
                s = s & "      ,v.Vendor_Name Company" & vbCrLf
                s = s & "      ,v.Vendor_ID Vendor" & vbCrLf
                s = s & "      ,v.City" & vbCrLf
                s = s & "      ,v.Phone " & vbCrLf
                s = s & "  FROM tblVendors v" & vbCrLf
                s = s & "       LEFT OUTER JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor and v.DivisionID = c.DivisionID" & vbCrLf
                s = s & "                                          AND (ISNULL(Community,'')='' OR Community=" & DbQuote(Str, mCommunity) & ")" & vbCrLf
                s = s & "                                          AND (ISNULL(Assembly,'')='' OR Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ")" & vbCrLf
                s = s & "                                          AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & vbCrLf
                s = s & "                                          AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & ")" & vbCrLf
                s = s & " WHERE v.inactive=0 and v.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "order by 1 desc,3" & vbCrLf
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Vendor", "")) Then
                    
                    NewVendor = FPickList.SelectedItem("Vendor")
                    NewVendorName = FPickList.SelectedItem("Company")
                
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(i, Col, b)
                        If Not b Then
                            OldVendorCSV = OldVendorCSV & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("poVendor")))
                            If .TextMatrix(i, .ColIndex("PONumber")) = "" Then
                                .Cell(flexcpText, i, .ColIndex("POVendor"), i, .ColIndex("POVendor")) = NewVendor
                                .Cell(flexcpText, i, .ColIndex("POVendorName"), i, .ColIndex("POVendorName")) = NewVendorName
                            End If
                        End If
                    Next
                    OldVendorCSV = Mid(OldVendorCSV, 2)
                    Call UpdateScheduleVendor(.TextMatrix(.Row, .ColIndex("Job_No")), OldVendorCSV, NewVendor)
                End If

            Case "BudgetTaxGroup", "POTaxGroup"
                s = "SELECT TaxGroup,Description FROM TaxGroups Where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gItems) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, Col) = FPickList.SelectedItem("TaxGroup")
                        End If
                    Next
                End If
        
            
            Case "POIndex", "POIndexDescription"
                s = "SELECT POIndex,Description FROM tblPOIndex where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "PO Index", s, gItems) Then
                    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_BeforeEdit(Row, Col, b)
                        If Not b Then
                            .Cell(flexcpText, Row, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                            .Cell(flexcpText, Row, .ColIndex("POIndexDescription")) = FPickList.SelectedItem("Description")
                        End If
                    Next
                End If


        End Select
        
        
        Call gItems_AfterEdit(Row, Col)
        
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "gItems_CellButtonClick")
End Sub

Private Function GetCorrectingEntryRowNum(Optional OrigRowNum As Long, Optional ReversingRowNum As Long) As Long
    Dim i As Long
    Dim r As Long
    With gItems
        
        'logic is same whether you provided a reversing entry or an original entry row num
        r = -1
        If OrigRowNum <> 0 Then r = OrigRowNum
        If ReversingRowNum <> 0 Then r = ReversingRowNum
        If r < 0 Or r > .Rows - 1 Then
            GetCorrectingEntryRowNum = -1
        Else
            'look for id in the correctingitemid column
            GetCorrectingEntryRowNum = .FindRow(.TextMatrix(r, .ColIndex("CorrectingItemID")), , .ColIndex("EstItemID"))
        End If
            
    End With
End Function



Private Function GetOriginalEntryRowNum(Optional ReversingRowNum As Long, Optional CorrectingRowNum As Long) As Long
    Dim i As Long
    Dim r As Long
    With gItems
        
        'you provided a correcting entry row num
        If CorrectingRowNum > 0 Then
            r = CorrectingRowNum
            'correcting item id will be set for the reversing item and the original item so you will need to check that you have found the correct row
            
            'out of bounds
            If r < 0 Then
                GetOriginalEntryRowNum = -1
            Else
                i = .FindRow(.TextMatrix(r, .ColIndex("EstItemID")), , .ColIndex("CorrectingItemID"))
                If i < 0 Or i = r Then
                    GetOriginalEntryRowNum = -1
                    Exit Function
                End If
                'check row type
                If .TextMatrix(i, .ColIndex("IsReversingItem")) <> "true" Then
                    GetOriginalEntryRowNum = i
                Else
                    'search again
                    GetOriginalEntryRowNum = .FindRow(.TextMatrix(r, .ColIndex("EstItemID")), i + 1, .ColIndex("CorrectingItemID"))
                End If
            
            End If

        Else
        'you provided a reversing entry row num
            r = ReversingRowNum
            
            'out of bounds
            If r < 0 Then
                GetOriginalEntryRowNum = -1
            Else
                'look for id in the reversingitemid column
                GetOriginalEntryRowNum = .FindRow(.TextMatrix(r, .ColIndex("EstItemID")), , .ColIndex("ReversingItemID"))
            End If
        
        End If
    End With
    
End Function


Private Function GetReversingEntryRowNum(Optional OrigRowNum As Long, Optional CorrectingRowNum As Long) As Long
    Dim i As Long
    Dim r As Long
    With gItems
        
        'you provided a correcting entry row num
        If CorrectingRowNum > 0 Then
            r = CorrectingRowNum
            'correcting item id will be set for the reversing item and the original item so you will need to check that you have found the correct row
            
            'check out of bounds
            If r < 0 Then
                GetReversingEntryRowNum = -1
                Exit Function
            End If
            
            'search for id
            i = .FindRow(.TextMatrix(r, .ColIndex("EstItemID")), , .ColIndex("CorrectingItemID"))
            If i < 0 Or i = r Then
                GetReversingEntryRowNum = -1
                Exit Function
            End If
            'check row type
            If .TextMatrix(i, .ColIndex("IsReversingItem")) = "true" Then
                GetReversingEntryRowNum = i
            Else
                'search again
                GetReversingEntryRowNum = .FindRow(.TextMatrix(r, .ColIndex("EstItemID")), i + 1, .ColIndex("CorrectingItemID"))
            End If
        
        
        Else
        'you provided an original entry row num
            MsgBox "not yet implemented"
            GetReversingEntryRowNum = -1
        
        End If
        
    End With
End Function

Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim s As String
    Dim r As Long
    Dim i As Long
    Dim CanDeleteBudget As Boolean
    Dim CanDeletePO     As Boolean
    
    
    With gItems
        Select Case True
                
            Case KeyCode = vbKeyF And Shift = vbCtrlMask
                Call FFind.ShowForm(gItems)
                
            Case KeyCode = vbKeyDelete And Shift = vbCtrlMask
                
                If mReadOnly Then Exit Sub
                For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                
                    CanDeleteBudget = .TextMatrix(r, .ColIndex("BudgetGenerated")) <> "true" _
                                       And .TextMatrix(r, .ColIndex("BudgetsLocked")) <> "True" _
                                       And .TextMatrix(r, .ColIndex("Budget1Qty")) = "" _
                                       And .TextMatrix(r, .ColIndex("Budget2Qty")) = "" _
                                       And .TextMatrix(r, .ColIndex("Budget3Qty")) = ""
                                       
                    'not po'd and not modified and not reversal entry
                    CanDeletePO = .TextMatrix(r, .ColIndex("PONumber")) = "" And .ValueMatrix(r, .ColIndex("ReversingItemID")) = 0 And .TextMatrix(r, .ColIndex("IsReversingItem")) <> "True"
                    
                    If CanDeleteBudget And Not (mMode = fcePO And Not CanDeletePO) Then .TextMatrix(r, .ColIndex("BudgetDeleted")) = "True"
                    If CanDeletePO And Not (mMode = fceBudget And Not CanDeleteBudget) Then .TextMatrix(r, .ColIndex("PODeleted")) = "True"
                    
                    If (CanDeleteBudget And mMode <> fcePO) Or (CanDeletePO And mMode = fcePO) Then
                        Dirty = True
                        
                        'if you deleted a correcting entry then correct original and hide its reversing entry
                        If .TextMatrix(r, .ColIndex("IsCorrectingItem")) = "true" Then
                            
                            'hide reversing
                            i = GetReversingEntryRowNum(, r)
                            If i > 0 Then
                                .RowHidden(i) = True
                                .TextMatrix(i, .ColIndex("PODeleted")) = "True"
                                If .RowData(i) <> "New" Then .RowData(i) = "Dirty"
                            End If
                            'correct original
                            i = GetOriginalEntryRowNum(, r)
                            If i > 0 Then
                                .TextMatrix(i, .ColIndex("RowType")) = "Original"
                                .TextMatrix(i, .ColIndex("ReversingItemID")) = ""
                                .TextMatrix(i, .ColIndex("CorrectingItemID")) = ""
                                If .RowData(i) <> "New" Then .RowData(i) = "Dirty"
                                Call ColorizeItems(i)
                            End If
                        End If
                        
                        'hide the row you deleted
                        .RowHidden(r) = True
                        If .RowData(r) <> "New" Then .RowData(r) = "Dirty"
                        
                    End If
                Next
                .Row = r
        End Select
    End With
    
End Sub

Private Sub gItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button <> 0 Then Exit Sub
    With gItems
        If mFunnyFlag Then
            mFunnyFlag = False
            Exit Sub
        End If
        If .MouseRow >= Min(.Row, .RowSel) And .MouseRow <= Max(.Row, .RowSel) And .MouseCol = .Col Then
            'dont do this. it causes the cell edit to end after a single keystroke
            'Call gItems_SelChange
        Else
            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315)
        End If
    End With
End Sub


Private Sub gItems_SelChange()
    Dim r As Long
    Dim prefix As String

    Dim pretax As Double
    Dim NJCTax As Double
    Dim tax As Double
    Dim tip As String
    Dim qty As Double
    
    gItems.ColSel = gItems.Col
    
    If mMode <> fcePO Then
        prefix = "Budget"
    Else
        prefix = "PO"
    End If
    
    With gItems
        If .Row <> .RowSel Then
            For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            If .RowHidden(r) = False And Not .IsSubtotal(r) Then
                pretax = pretax + .ValueMatrix(r, .ColIndex(prefix & "Pretax"))
                NJCTax = NJCTax + .ValueMatrix(r, .ColIndex(prefix & "NJCTax"))
                tax = tax + .ValueMatrix(r, .ColIndex(prefix & "JCTax"))
                If IsIn(.ColKey(.Col), "TakeoffQty", "OrderQty", "BudgetQty") Then
                    qty = qty + .ValueMatrix(r, .Col)
                End If
            End If
            Next
            
                

            
            tip = IIf(IsIn(.ColKey(.Col), "TakeoffQty", "OrderQty", "BudgetQty"), .ColKey(.Col) & ":" & vbTab & qty & vbCrLf & vbCrLf, "") & _
            "Pretax:" & vbTab & format(pretax, "#,##0.00") & vbCrLf & _
            "Costed Tax:" & vbTab & format(tax, "#,##0.00") & vbCrLf & vbCrLf & _
            "Subtotal:" & vbTab & format(pretax + tax, "#,##0.00") & vbCrLf
            If NJCTax <> 0 Then
                tip = tip & "Non Costed Tax:" & vbTab & format(NJCTax, "#,##0.00") & vbCrLf & _
                "Total:" & vbTab & format(pretax + tax + NJCTax, "#,##0.00")
            End If
                  
            
            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315, tip)
            
        End If
    End With
End Sub


Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh

    Dim s As String
    Dim i As Long
    Dim Description As String
    Dim rs As Recordset
    Dim rowPOed As Boolean
    Dim b As Boolean
    
    With gItems
    
    'grid used to use repeat fillstyle to do multirow edits but that doesnt prevent edits to locked rows. changed to single fillstyle and implementing pre-check on each row
    For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
    Call gItems_BeforeEdit(Row, Col, b)
    If Not b Then
    
        s = .EditText
        rowPOed = Trim(.TextMatrix(Row, .ColIndex("PONumber"))) <> ""
        Select Case .ColKey(Col)
            Case "SalesQty": s = Val(s)
            
            Case "JCExtra"
                If HFApp.Options(Use_Timberline) And Trim(s) <> "" Then
                    Set rs = HFApp.SqlExec("select extra from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No")))) & " AND extra=" & DbQuote(Str, s), dbAccounting)
                    If rs.EOF Then
                        Description = InputBox(vbCrLf & "Extra '" & s & "' could not be found." & vbCrLf & "Do you want to add it?" & vbCrLf & vbCrLf & vbCrLf & "Enter a description for the extra. (30 characters or less)", App.ProductName, txtHFOption.Text)
                        If Description = "" Then
                            Cancel = True
                            Exit Sub
                        Else
                            s = ""
                            s = s & "INSERT INTO jcm_master__extra(Job,Extra,Description,Status)" & vbCrLf
                            s = s & "VALUES(" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No")))) & vbCrLf
                            s = s & "      ," & DbQuote(Str, .EditText) & vbCrLf
                            s = s & "      ," & DbQuote(Str, Left(Description, 30)) & vbCrLf
                            s = s & "      ,'In progress')" & vbCrLf
                            Set rs = HFApp.SqlExec(s, dbAccounting)
                            s = .EditText
                        End If
                    Else
                        s = "" & rs(0)
                    End If
                End If
            
                                   
            Case "JCCostCode":              Cancel = Not ValidateField(gItems, s, "Cost Code not found", "SELECT CostCode,Description FROM StandardCostCodes WHERE DivisionID = " & HFApp.DivisionID & " and CostCode=" & DbQuote(Str, s), "JCCostCodeDesc")
            Case "JCCategory":
                Cancel = Not ValidateField(gItems, s, "Category not found", "SELECT Category,Description FROM StandardCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, s), "JCCategoryDesc")
                'if you updated a reversing or correcting entry then also update its pairing
                i = -1
                If .TextMatrix(Row, .ColIndex("IsReversingItem")) = "true" Then i = GetCorrectingEntryRowNum(, Row)
                If .TextMatrix(Row, .ColIndex("IsCorrectingItem")) = "true" Then i = GetReversingEntryRowNum(, Row)
                If i > 0 Then
                    .TextMatrix(i, .ColIndex("JCCategory")) = .TextMatrix(Row, .ColIndex("JCCategory"))
                    .TextMatrix(i, .ColIndex("JCCategoryDesc")) = .TextMatrix(Row, .ColIndex("JCCategoryDesc"))
                    If .RowData(i) <> "New" Then .RowData(i) = "Dirty"
                End If
            
            Case "VarianceJCCategory":      Cancel = Not ValidateField(gItems, s, "Category not found", "SELECT Category,Description FROM StandardCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, s), "VarianceJCCategoryDesc")
            Case "ItemDesc":
            Case "ItemComments":
            Case "BudgetVendor":            Cancel = Not ValidateField(gItems, s, "Vendor not found or is inactive", "SELECT Vendor_ID,Vendor_Name FROM tblVendors WHERE DivisionID =" & HFApp.DivisionID & " and Vendor_ID=" & DbQuote(Str, s), "BudgetVendorName")
            Case "POVendor":                Cancel = Not ValidateField(gItems, s, "Vendor not found or is inactive", "SELECT Vendor_ID,Vendor_Name FROM tblVendors WHERE DivisionID =" & HFApp.DivisionID & " and Vendor_ID=" & DbQuote(Str, s), "POVendorName") And Not rowPOed
            Case "POVendorName":            Cancel = Not rowPOed
            Case "ConversionFactor":        Cancel = Not IsNumeric(s)
            Case "BudgetQty":               Cancel = Not IsNumeric(s)
            Case "POQty":                   Cancel = Not IsNumeric(s)
            Case "OrderUOM":
            Case "BudgetRate":              Cancel = Not IsNumeric(s)
            Case "PORate":                  Cancel = Not IsNumeric(s)
            Case "BudgetPretax":            Cancel = Not IsNumeric(s)
            Case "POPretax":                Cancel = Not IsNumeric(s)
            
            Case "BudgetTaxGroup":          Cancel = Not ValidateField(gItems, s, "Tax Group not found", "SELECT TaxGroup FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, s))
                
            Case "POTaxGroup":              Cancel = Not ValidateField(gItems, s, "Tax Group not found", "SELECT TaxGroup FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, s))
            Case "BudgetJCTax":             Cancel = Not IsNumeric(s)
            'case "BudgetJCTaxRate":
            Case "POJCTax":                 Cancel = Not IsNumeric(s)
            'case "POJCTaxRate":
            Case "BudgetNJCTax":            Cancel = Not IsNumeric(s)
            'case "BudgetNJCTaxRate":
            Case "PONJCTax":                Cancel = Not IsNumeric(s)
            'case "PONJCTaxRate":
            'case "BudgetTax":
            'case "POTax":
            Case "POIndex":                 Cancel = Not ValidateField(gItems, s, "PO Index not found", "SELECT POIndex,Description POIndexDescription FROM tblPOIndex WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, s), "POIndexDescription")
            'case "PONumber":
            'case "BudgetGenerated":
            Case "ExcludeFromPO":
            'case "SalesWorksheet":
        End Select
        .EditText = s
        .TextMatrix(Row, Col) = s
        If .ColKey(Col) <> "Selected" Then
            Dirty = True
            If .RowData(Row) <> "NEW" Then .RowData(Row) = "DIRTY"
        End If
        
    End If
    Next
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "gItems_ValidateEdit")
End Sub

Private Sub LoadJob(Job As String)
On Error GoTo eh
Dim dbg As String
    Dim u As String
    Dim s As String
    Dim i As Long
    Dim rs As Recordset
    Dim QuoteNo As String
    If Not SaveData(True) Then Exit Sub
    
    Call Clear
    
    If Job = "" Then
        
        If mMode = fceQuote Then
            s = ""
            s = s & "SELECT DISTINCT j.Job_No Quote" & vbCrLf
            s = s & "      ,j.Description" & vbCrLf
            s = s & "      ,l.description ""Area/Project""" & vbCrLf
            s = s & "      ,j.CommunityPhase Phase" & vbCrLf
            s = s & "      ,j.Municipal_Address Address" & vbCrLf
            s = s & "      ,j.PM,j.Purchaser" & vbCrLf
            s = s & "      ,j.Estimator" & vbCrLf
            s = s & "      ,j.Notes" & vbCrLf
            s = s & "FROM tblCustomers c JOIN tblJobs j ON (c.Job_No = j.Job_No and c.DivisionID = j.DivisionID)" & vbCrLf
            s = s & "     left outer join tbllocality l on(j.community = l.area)" & vbCrLf
            s = s & "    join system_setup s on s.id = c.DivisionID" & vbCrLf
            s = s & "WHERE ISNULL(j.Inactive,0)=0 AND c.job_no<>'' AND " & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "(c.DivisionID = " & HFApp.DivisionID & ") AND " & vbCrLf
            End If
            s = s & "((c.purchased=1 AND c.cancelled=0 AND c.inactive=0 AND ((isnull(s.accounting_approve,0)=1 AND c.approved=1) OR (isnull(s.accounting_approve,0)=0 AND c.contract_assigned=1) OR isnull(s.PurchasingRequiresSalesApproval,0)=0))" & vbCrLf
            s = s & "  OR " & vbCrLf
            s = s & "(c.home_selection <>'PreSale' OR c.presale_selection<>'PreSale'))" & vbCrLf
            s = s & "  AND j.isquote=1" & vbCrLf
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), IIf(mMode = fceQuote, "Quote", "Job"), s, mJob, , True) Then Exit Sub
            Job = FPickList.SelectedItem("Quote")
        Else
           If ProjectBased Then
                s = ""
                s = s & "SELECT distinct c.area Project" & vbCrLf
                s = s & "      ,c.Description" & vbCrLf
                s = s & "FROM tblLocality c join tbljobs j on c.area=j.community"
                If HFApp.DivisionID <> "" Then
                     s = s & " where (j.DivisionID = " & HFApp.DivisionID & ")"
                End If
                If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Project", s, mCommunity) Then Exit Sub
                mCommunity = FPickList.SelectedItem("Project")
                On Error Resume Next
                ProjectPhaseBased = "" & HFApp.SqlExec("Select count(*) from CommunityPhase where Community = " & DbQuote(Str, mCommunity))(0) <> 0
                On Error GoTo eh
                Call LoadViews
           
           Else
                s = ""
                s = s & "Open " & IIf(mMode = fceQuote, "Quotes", "Jobs") & Chr(1) & vbCrLf
                s = s & "SELECT DISTINCT j.Job_No Job" & vbCrLf
                s = s & "      ,j.Description" & vbCrLf
                s = s & "      ,l.description """ & lblCommunity.Caption & """" & vbCrLf
                s = s & "      ,j.CommunityPhase Phase" & vbCrLf
                s = s & "      ,j.Municipal_Address Address" & vbCrLf
                s = s & "      ,j.PM,j.Purchaser" & vbCrLf
                s = s & "      ,j.Estimator" & vbCrLf
                s = s & "      ,j.Notes" & vbCrLf
                s = s & "FROM tblCustomers c JOIN tblJobs j ON (c.Job_No = j.Job_No and c.DivisionID = j.DivisionID)" & vbCrLf
                s = s & "     left outer join tbllocality l on(j.community = l.area)" & vbCrLf
                s = s & " join system_setup s on s.id = c.DivisionID" & vbCrLf
                s = s & "WHERE j.Inactive=0 AND c.job_no<>'' AND " & vbCrLf
                If HFApp.DivisionID <> "" Then
                    s = s & "(c.DivisionID = " & HFApp.DivisionID & ") AND " & vbCrLf
                End If
                s = s & "((c.purchased=1 AND c.cancelled=0 AND c.inactive=0 AND ((isnull(s.accounting_approve,0)=1 AND c.approved=1) OR (isnull(s.accounting_approve,0)=0 AND c.contract_assigned=1) OR isnull(s.PurchasingRequiresSalesApproval,0)=0))" & vbCrLf
                s = s & "  OR " & vbCrLf
                s = s & "(c.home_selection <>'PreSale' OR c.presale_selection<>'PreSale'))" & vbCrLf
                s = s & "  AND j.isquote=0" & vbCrLf
                s = s & Chr(0) & "Closed " & IIf(mMode = fceQuote, "Quotes", "Jobs") & Chr(1) & vbCrLf
                s = s & "SELECT DISTINCT " & vbCrLf
                s = s & "  j.Job_No Job" & vbCrLf
                s = s & " ,j.Description" & vbCrLf
                s = s & " ,l.description """ & lblCommunity.Caption & """" & vbCrLf
                s = s & " ,j.CommunityPhase Phase" & vbCrLf
                s = s & " ,j.Municipal_Address Address" & vbCrLf
                s = s & " ,j.PM" & vbCrLf
                s = s & " ,j.Purchaser" & vbCrLf
                s = s & " ,j.Estimator" & vbCrLf
                s = s & "" & vbCrLf
                s = s & "FROM tblJobs j " & vbCrLf
                s = s & "left outer join tbllocality l on j.community=l.area" & vbCrLf
                s = s & "join system_setup s on s.id=j.DivisionID" & vbCrLf
                s = s & "" & vbCrLf
                s = s & "WHERE j.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "and j.inactive=1" & vbCrLf
                
                
                If Not FPickList.Choose(HFApp.Databases(dbHomefront), IIf(mMode = fceQuote, "Quote", "Job"), s, mJob) Then Exit Sub
                Job = FPickList.SelectedItem("Job")
            End If
        End If
    End If
        
        
    'lock new job, unlock old job, no locking on projects...
    If Job <> "" And Not ProjectBased Then
        u = HFApp.RecordLockedBy("BudgetsAndPOs", Job)
        If u <> "" Then
            MsgBox "Unable to open this job. It is in use by " & u, vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    
    'unlock record if is locked by me. this looks wrong but it isnt.
    'RecordLockedBy() only returns other user names. If is locked by me then it returns nothing.
    u = HFApp.RecordLockedBy("BudgetsAndPOs", mJob)
    If u = "" Then Call HFApp.UnLockRecord("BudgetsAndPOs", mJob)
    
    If ProjectBased Then
        mJob = HFApp.SqlExec("Select top 1 job_no from tbljobs where Divisionid = " & HFApp.DivisionID & " and community=" & DbQuote(Str, mCommunity))(0)
    Else
        mJob = Job
        Call HFApp.LockRecord("BudgetsAndPOs", mJob)
    End If
    FMain.CurrentJob = mJob
        
        
        
        
        
        
        
        
dbg = "1"
        
    'reset screen
    mTakeoffSettings = ""
    Toolbar.Buttons("takeoffsettings").value = tbrUnpressed
    frmAssembly.Visible = False
    frmPOIndex.Visible = False
    Call Form_Resize
    gAssemblies.Rows = 1
    gItems.Rows = 1
    
dbg = "2"
    
    
    'read in job
    Me.Tag = mMode & Chr(0) & mJob
    s = ""
    s = s & "select j.*" & vbCrLf
    s = s & "      ,c.ar_customer_deposit ARCustomer,m.Description ModelDesc" & vbCrLf
    s = s & "  from tbljobs j" & vbCrLf
    s = s & "  left outer join tblcustomers c on(j.job_no=c.job_no and j.DivisionID = c.DivisionID)" & vbCrLf
    s = s & "  left outer join DistinctModelsbyDivision m on(m.Model=j.model and m.DivisionID = j.DivisionID)" & vbCrLf
    s = s & " where j.DivisionID = " & HFApp.DivisionID & " and j.job_no=" & DbQuote(Str, mJob) & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)

    If rs.EOF Then
        
        'if this is a new quote then get next quote number
        If mMode = fceQuote Then
            On Error Resume Next
            QuoteNo = ""
            QuoteNo = HFApp.SqlExec("select max(cast(job_no as float)) + 1 from tbljobs where DivisionID = " & HFApp.DivisionID & " and isnumeric(job_no)=1", dbHomefront)(0)
            On Error GoTo 0
        End If
        
        Me.Caption = Choose(mMode, "Prepare Quote", "Issue Budgets", "Issue PO's")
        mCommunity = ""
        mCommunityPhase = ""
        mTakeoffSystemBidID = 0
        mTakeoffSystemProjectName = ""
        ReadOnly = False
        txtJob.Text = QuoteNo
        txtAddress.Text = ""
        txtDescription.Text = ""
        txtNotes.Text = ""
        cboCommunity.ListIndex = -1
        cboPhase.Text = ""
        cboModel.Text = ""
        txtLot.Text = ""
        txtBlock.Text = ""
        txtLotPlan.Text = ""
        gAssemblies.Rows = 1
        gItems.Rows = 1
        
    Else
dbg = "3"
        mJob = "" & rs("Job_No")
        ReadOnly = "" & rs("Status") = "Locked"
        mTakeoffSystemBidID = Val("" & rs("OnScreenBidID"))
        mTakeoffSystemProjectName = "" & rs("OnScreenProjectName")
        Me.Caption = Choose(CSng(mMode), "Prepare Quote", "Issue Budgets", "Issue PO's") & " - " & mJob
        mCommunity = "" & rs("Community")
        mCommunityPhase = "" & rs("CommunityPhase")
        QuoteNo = "" & rs("Quote_No")
        txtAddress.Text = "" & rs("Municipal_Address")
        txtJob.Text = "" & rs("job_no")
        txtDescription.Text = "" & rs("Description")
        txtNotes.Text = "" & rs("Notes")
        Call SetComboBoxListIndex(cboCommunity, , mCommunity)
        cboPhase.Text = mCommunityPhase
        cboModel.Text = "" & rs("Model")
        Me.lblModelDesc.Caption = "" & rs("ModelDesc")
        txtLot.Text = "" & rs("Lot")
        txtBlock.Text = "" & rs("Block")
        txtLotPlan.Text = "" & rs("LotPlan")
    
        
    End If
    
    
    
dbg = "5"
    Call LoadWBSDescriptions
    Call LoadProperties
    
dbg = "6"
        
    'load assemblies and items
    Call EnableAssemblies(mJob <> "")
    If mJob <> "" Then
    
        'if job has nothing in estimateitems then add a dummy row
        s = ""
        s = s & "INSERT INTO EstimateAssemblies(Customer_No,Job,HFDescription,AssemblyType,DivisionID)" & vbCrLf
        s = s & "SELECT TOP 1 c.Customer_No,c.Job_No,'Manual Estimates',-1,c.DivisionID" & vbCrLf
        s = s & "  FROM tblCustomers c LEFT OUTER JOIN EstimateAssemblies a ON(c.DivisionID = a.DivisionID and c.Job_No=a.Job)" & vbCrLf
        s = s & " WHERE a.EstAssemblyID IS NULL" & vbCrLf
        s = s & "   AND c.Job_no=" & DbQuote(Str, mJob) & " and c.DivisionID = " & HFApp.DivisionID & vbCrLf
        Call HFApp.SqlExec(s)
        
dbg = "7"
        Dirty = False
        Call mnuEstimateItemViewsSub_Click(CInt(mViewIndex) + 1)
        'Call LoadAssemblies(True)
dbg = "8"
    End If
    
    
    
    Dirty = False
Exit Sub
eh:
MsgBox Err.Description & vbCrLf & dbg
End Sub

Private Sub LoadWBSDescriptions()
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    If mJob = "" Then Exit Sub
    s = "select * from tbljobs where DivisionID = " & HFApp.DivisionID & " and  job_no=" & DbQuote(Str, mJob)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    If Not rs.EOF Then
    With gItems
        For i = 1 To 40
            .TextMatrix(0, .ColIndex("WBS" & format(i, "00"))) = "" & rs("WBSDesc" & format(i, "00"))
            If .TextMatrix(0, .ColIndex("WBS" & format(i, "00"))) = "" Then .ColHidden(.ColIndex("WBS" & format(i, "00"))) = True
        Next
    End With
    End If
End Sub

Private Function GetComboList(fieldname As String) As String
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    s = "select distinct " & fieldname & " from estimateitems where job=" & DbQuote(Str, mJob) & " order by 1"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    s = ""
    While Not rs.EOF
        If Trim("" & rs(0)) <> "" Then s = s & "|" & rs(0)
        rs.MoveNext
    Wend
    If s = "|" Then s = ""
    GetComboList = s
    
End Function


Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim s As String
    Dim r As Long
    Dim community As String
    Dim CommunityPhase As String
    Dim i As Long
    
    Dim ei As Long
    Dim es As String
    
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
        
        
    txtJob.Text = Trim(txtJob.Text)
    If Trim(txtJob.Text) = "" Then
        MsgBox "A quote number is required.", vbExclamation, App.ProductName
        Call SetCtrlFocus(txtJob)
        Exit Function
    End If
    If mJob = "" Then
        'create new quote records
        'create job record first - stop on duplicate key. job must not exist otherwise you can end up deleting customer data
        s = "INSERT INTO tblJobs(DivisionID,Job_no) VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, txtJob.Text) & ")"
        On Error Resume Next
        HFApp.SqlExec s
        If Err.Number <> 0 Then
            If Err.Description Like "*Duplicate*" Then
                MsgBox "A job or quote with this number already exists. Please use a different number.", vbExclamation, App.ProductName
                Call SetCtrlFocus(txtJob)
                Exit Function
            Else
                ei = Err.Number
                es = Err.Description
                On Error GoTo eh
                Err.Raise ei, , es
            End If
        End If
        On Error GoTo eh
    End If
    
    
    Screen.MousePointer = vbHourglass
    community = GetComboBoxListKey(cboCommunity)
    CommunityPhase = cboPhase.Text
    
    With gProperties
        If mJob = "" Then
            'create new quote records
            s = ""
            s = s & "INSERT INTO tblcustomers(DivisionID,Customer_No,Job_no,Description,community,phase,purchased,approved"
            s = s & "                        ,contract_assigned,cancelled,inactive,lot,block,lotplan,model,sale_posted)" & vbCrLf
            s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, community) & vbCrLf
            s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
            s = s & "      ,1,1,1,0,0"
            s = s & "      ," & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "      ,0" & vbCrLf
            s = s & ")"
            HFApp.SqlExec s
            
            s = ""
            s = s & "UPDATE tblJobs" & vbCrLf
            s = s & "SET IsQuote=" & DbQuote(Bit, mMode = fceQuote) & vbCrLf
            s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "   ,Notes=" & DbQuote(Str, txtNotes.Text) & vbCrLf
            If "" & community <> "" Then
                s = s & "   ,Community=" & DbQuote(Str, community) & vbCrLf
                s = s & "   ,CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf
            End If
            s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "   ,Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "   ,OnScreenBidID=" & DbQuote(Num, mTakeoffSystemBidID) & vbCrLf
            s = s & "   ,OnScreenProjectName=" & DbQuote(Str, mTakeoffSystemProjectName) & vbCrLf
            
            'properties
            If HFApp.Options.ValueByName("EnableTarionFields") = "true" Then
                s = s & "   ,TarionEnrollmentNumber=" & DbQuote(Str, .TextMatrix(mprop_TarionEnrollmentNumber, .ColIndex("value"))) & vbCrLf
                s = s & "   ,TarionBuilderNumber=" & DbQuote(Str, .TextMatrix(mprop_TarionBuilderNumber, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,PermitNumber=" & DbQuote(Str, .TextMatrix(mprop_PermitNumber, .ColIndex("value"))) & vbCrLf
            s = s & "   ,PermitReceivedDate=" & DbQuote(Date, .TextMatrix(mprop_PermitDate, .ColIndex("value"))) & vbCrLf
            If IsMultiFamily Then
                s = s & "   ,ShellScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_ShellTemplate, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,ScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_UnitTemplate, .ColIndex("value"))) & vbCrLf
            s = s & "   ,Start_Date=" & DbQuote(Date, .TextMatrix(mprop_ConstStart, .ColIndex("value"))) & vbCrLf
            If HFApp.Options(AccountingSystem) = asIntacct Then
                s = s & "   ,IntacctDepartment=" & DbQuote(Str, .TextMatrix(mprop_IntacctDepartment, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,LabourTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_LabTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,MaterialTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_MatTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,SubContractTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_SubTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,EquipmentTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_EquTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OverheadTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_OvrTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OtherTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_OthTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,ARTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_ARTax, .ColIndex("value"))) & vbCrLf
            
            
            s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Job_No=" & DbQuote(Str, txtJob.Text)
            HFApp.SqlExec s
            mJob = txtJob.Text
            Call EnableCtrl(txtJob, False)
            
            
            s = ""
            s = s & "INSERT INTO EstimateAssemblies(Customer_No,Job,HFDescription,AssemblyType,DivisionID)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ,'Quoted Items',-1," & HFApp.DivisionID & ")" & vbCrLf
            Call HFApp.SqlExec(s)
            
            Call EnableAssemblies(True)
            Call LoadAssemblies(True)
            
        Else
            s = ""
            s = s & "UPDATE tblJobs" & vbCrLf
            s = s & "SET Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "   ,Notes=" & DbQuote(Str, txtNotes.Text) & vbCrLf
            If "" & community <> "" Then
                s = s & "   ,Community=" & DbQuote(Str, community) & vbCrLf
                s = s & "   ,CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf
            End If
            s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "   ,Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "   ,OnScreenBidID=" & DbQuote(Num, mTakeoffSystemBidID) & vbCrLf
            s = s & "   ,OnScreenProjectName=" & DbQuote(Str, mTakeoffSystemProjectName) & vbCrLf
            
            'properties
            If HFApp.Options.ValueByName("EnableTarionFields") = "true" Then
                s = s & "   ,TarionEnrollmentNumber=" & DbQuote(Str, .TextMatrix(mprop_TarionEnrollmentNumber, .ColIndex("value"))) & vbCrLf
                s = s & "   ,TarionBuilderNumber=" & DbQuote(Str, .TextMatrix(mprop_TarionBuilderNumber, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,PermitNumber=" & DbQuote(Str, .TextMatrix(mprop_PermitNumber, .ColIndex("value"))) & vbCrLf
            s = s & "   ,PermitReceivedDate=" & DbQuote(Date, .TextMatrix(mprop_PermitDate, .ColIndex("value"))) & vbCrLf
            If IsMultiFamily Then
                s = s & "   ,ShellScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_ShellTemplate, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,ScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_UnitTemplate, .ColIndex("value"))) & vbCrLf
            s = s & "   ,Start_Date=" & DbQuote(Date, .TextMatrix(mprop_ConstStart, .ColIndex("value"))) & vbCrLf
            If HFApp.Options(AccountingSystem) = asIntacct Then
                s = s & "   ,IntacctDepartment=" & DbQuote(Str, .TextMatrix(mprop_IntacctDepartment, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,LabourTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_LabTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,MaterialTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_MatTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,SubContractTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_SubTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,EquipmentTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_EquTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OverheadTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_OvrTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OtherTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_OthTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,ARTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_ARTax, .ColIndex("value"))) & vbCrLf
                        
            s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Job_No=" & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
            HFApp.SqlExec s

            If HFApp.Options(SalesSystem) = SalesSystems.asNone Or HFApp.Options(SalesSystem) = SalesSystems.asHomeFront Then
                s = ""
                s = s & "UPDATE tblcustomers" & vbCrLf
                s = s & "SET Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
                If "" & community <> "" Then
                    s = s & "   ,Community=" & DbQuote(Str, community) & vbCrLf
                End If
                s = s & "   ,Phase=" & DbQuote(Str, CommunityPhase) & vbCrLf
                s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
                s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
                s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
                s = s & "   ,Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
                s = s & "WHERE Customer_No=" & DbQuote(Str, txtJob.Text)
                HFApp.SqlExec s
            End If
            
        End If
        
        
        'save properties
        s = ""
        For r = mprop_UserFlds To .Rows - 1
            If .TextMatrix(r, .ColIndex("DataType")) <> "" Then
                s = s & "," & vbQuote & .TextMatrix(r, .ColIndex("Name")) & vbQuote & "=" & DbQuote(.ValueMatrix(r, .ColIndex("DataType")), .TextMatrix(r, .ColIndex("Value")))
            End If
        Next
        If s <> "" Then
            On Error Resume Next
            Call HFApp.SqlExec("INSERT INTO JobCustomFields(Job_No,divisionid) VALUES(" & DbQuote(Str, mJob) & "," & DbQuote(Num, HFApp.DivisionID) & ")", dbHomefront)
            Call HFApp.SqlExec("UPDATE JobCustomFields SET " & Mid(s, 2) & " WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and Job_No=" & DbQuote(Str, mJob), dbHomefront)
            On Error GoTo eh
        End If
        
    End With
    
    Call HFApp.WriteJobToAccounting(mJob)
    
    Toolbar.Buttons("NewRFQ").Enabled = mJob <> ""
    If SaveItems And SaveBids And SaveContract Then
        SaveData = True
        Dirty = False
    End If
    Screen.MousePointer = vbDefault
    
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function

Private Sub gProperties_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh

    
    With gProperties
        .ComboList = ""

        If Col <> .ColIndex("Value") Then
            Cancel = True
            Exit Sub
        End If

        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt
            Case Num
            Case Cur
            Case Bit:      .ComboList = "Yes|No"
            Case Date:     .ComboList = "|..."
            Case DateTime: .ComboList = "|..."
            Case Str:      .EditMaxLength = .ValueMatrix(Row, .ColIndex("Length"))
                           .ComboList = .TextMatrix(Row, .ColIndex("PickList"))
            Case Else:     Cancel = True
        End Select
        
    End With

Exit Sub
eh: Call errHandler(SRCFILE & "gProperties_BeforeEdit")
End Sub

Private Sub gProperties_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo eh
    With gProperties
        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt:
            Case Num:
            Case Cur:
            Case Bit:
            Case DateTime, Date
                Call DCalendar.Popup(gProperties, .RowPos(.Row) + .RowHeight(.Row), .colPos(.Col))
                .RowData(Row) = "DIRTY"
                Dirty = True
                
            Case Str:
            Case Else:
        End Select
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "gProperties_CellButtonClick")
End Sub


Private Sub gProperties_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gProperties
        If .Row < 0 Then Exit Sub
        Select Case KeyCode

            Case vbKeySpace
                .IsCollapsed(.Row) = IIf(.IsCollapsed(.Row) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
                
            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
                
            Case vbKeyDelete
                If .Col = .ColIndex("value") Then
                    .Text = ""
                    Call gProperties_ValidateEdit(.Row, .Col, False)
                    Dirty = True
                    
                    Select Case .ValueMatrix(.Row, .ColIndex("DataType"))
                        Case Cur, Num, NumInt:  .Text = "0"
                    End Select
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
End Sub

Private Sub gProperties_RowColChange()
On Error GoTo eh

    With gProperties
        If .Row >= 0 Then
            If .TextMatrix(.Row, .ColIndex("DataType")) = "" Then
                .Col = .ColIndex("Name")
            Else
                .Col = .ColIndex("Value")
            End If
        End If
    End With

Exit Sub
eh: Call errHandler(SRCFILE & "gProperties_ValidateEdit")
End Sub

Private Sub gProperties_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh
    
    With gProperties
        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt:    .EditText = Int(Val(.EditText))
            Case Num:       .EditText = Val(.EditText)
            Case Cur:       .EditText = format(Round(Val("" & Replace(Replace(Replace(.EditText, "%", ""), ",", ""), "$", "")), 2), "Currency")
            Case Bit:
            Case DateTime:
                If IsDate(.EditText) Or .EditText = "" Then
                    .EditText = format(.EditText, HFApp.Options(DateFormat))
                Else
                    Cancel = True
                End If
            Case Str:
            Case Else:       Cancel = True
        End Select
    
        If Not Cancel Then
            .RowData(Row) = "DIRTY"
            Dirty = True
        End If
    
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "gProperties_ValidateEdit")
End Sub


Private Property Get Dirty() As Boolean
    Dirty = mDirty
End Property
Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
End Property

Private Function CleanJob(FormatedJob As String) As String
    If HFApp.Options(AccountingSystem) = asTimberline Then
        CleanJob = Trim(Replace(Replace(Replace(Replace(Replace(FormatedJob, "\", ""), ",", ""), "/", ""), ".", ""), "-", ""))
    Else
        CleanJob = FormatedJob
    End If
End Function

Private Sub EnableAssemblies(RHS As Boolean)

    Call EnableCtrl(gAssemblies, RHS)
    Call EnableCtrl(gItems, RHS)
    
    Toolbar.Buttons("View").Enabled = RHS
    Toolbar.Buttons("ViewPOs").Enabled = RHS
    Toolbar.Buttons("ViewBudgets").Enabled = RHS
    
    If ReadOnly Then RHS = False
    
    Toolbar.Buttons("TakeoffOneTime").Enabled = RHS
    Toolbar.Buttons("TakeoffItem").Enabled = RHS
    Toolbar.Buttons("TakeoffAssembly").Enabled = RHS
    Toolbar.Buttons("TakeoffPlanSwift").Enabled = RHS
    Toolbar.Buttons("TakeoffCustom").Enabled = RHS
    Toolbar.Buttons("takeoffsettings").Enabled = RHS
    Toolbar.Buttons("RePrice").Enabled = RHS
    Toolbar.Buttons("Generate").Enabled = RHS
    
    Toolbar.Buttons("snapshots").Enabled = RHS
    

End Sub

Private Sub DeleteQuote()
    Dim i As Long
    
    If mJob = "" Or mMode <> fceQuote Then Exit Sub
    If MsgBox("Are you sure you want to permanently delete this quote?", vbYesNo + vbExclamation, App.ProductName) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Call HFApp.SqlExec("delete from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mJob))
    Call HFApp.SqlExec("delete from tblcustomers where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mJob))
    Call HFApp.SqlExec("delete from estimateassemblies where job=" & DbQuote(Str, mJob) & " and DivisionID = " & HFApp.DivisionID)
    Call HFApp.SqlExec("delete from estimateitems where job=" & DbQuote(Str, mJob) & " and DivisionID = " & HFApp.DivisionID)
    Call HFApp.SqlExec("delete from addons where job=" & DbQuote(Str, mJob))
    
    'clear everything
    mJob = ""
    txtJob.Text = ""
    txtDescription.Text = ""
    cboCommunity.ListIndex = -1
    cboPhase.Text = ""
    cboModel.ListIndex = -1
    txtLot.Text = ""
    txtLotPlan.Text = ""
    txtBlock.Text = ""
    txtNotes.Text = ""
    gProperties.Rows = 0
    gAssemblies.Rows = 1
    gItems.Rows = 1
    Me.Caption = Choose(mMode, "Prepare Quote", "Issue Budgets", "Issue PO's")
    Dirty = False
    
    'disable everything
    EnableAssemblies False
    EnableCtrl txtJob, False
    EnableCtrl txtDescription, False
    EnableCtrl cboCommunity, False
    EnableCtrl cboPhase, False
    EnableCtrl cboModel, False
    EnableCtrl txtLot, False
    EnableCtrl txtLotPlan, False
    EnableCtrl txtBlock, False
    EnableCtrl gProperties, False
    txtNotes.Enabled = False
    
    For i = 2 To Toolbar.Buttons.Count
        Toolbar.Buttons(i).Enabled = False
    Next
    Toolbar.Buttons("Open").Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub SaveAs(Quote As Boolean)
On Error GoTo eh
    Dim newJob As String
    Dim s As String
    Dim community As String
    Dim CommunityPhase As String
    Dim rs As Recordset
    
    Dim AssemblyID As Long
    Dim NewAssemblyID As Long
    
    If mJob = "" Or mMode <> fceQuote Then Exit Sub
    
GetJob:
    Do
        newJob = Trim(InputBox(vbCrLf & vbCrLf & vbCrLf & IIf(Quote, "Save quote as", "Create job number"), IIf(Quote, "Save As", "Create Job"), mJob))
        
        If newJob <> "" Then
            Set rs = HFApp.SqlExec("Select Job_No from tbljobs where DivisionID = " & HFApp.DivisionID & " and Job_No = " & DbQuote(Str, newJob))
            If Not rs.EOF Then
                MsgBox "Job Number already exists. Please try again", vbOKOnly
                GoTo GetJob:
            End If
        End If
        If Len(newJob) > 12 Then MsgBox IIf(Quote, "Quote", "Job") & "numbers must be 12 characters or less.", vbInformation, App.ProductName
    
        If newJob <> "" And (Not Quote) Then
            If (Not ValidateJobNumber(newJob)) Then
                MsgBox "Incorrect job format." & vbCrLf & "Correct format is " & Replace(HFApp.Options(Job_Mask), "&", "x") & vbCrLf & "Reenter job.", vbExclamation, App.ProductName
                newJob = "bigenoughtocauselooptogoagain"
                
            End If
        End If
    
    Loop While Len(newJob) > 12
    If Len(newJob) = 0 Then Exit Sub
    newJob = StripFormating(newJob)
    
    Screen.MousePointer = vbHourglass
    community = GetComboBoxListKey(cboCommunity)
    CommunityPhase = cboPhase.Text
    
    
    'create new job/customer
    With gProperties
        s = ""
        s = s & "INSERT INTO tblJobs(DivisionID,IsQuote,Quote_No,Job_No,Description,Notes,Community,CommunityPhase" & vbCrLf
        s = s & "                   ,Lot,Block,LotPlan,Model" & vbCrLf
        s = s & "                   ,LabourTaxGroup,MaterialTaxGroup,SubContractTaxGroup,EquipmentTaxGroup,OverheadTaxGroup,OtherTaxGroup)" & vbCrLf
        s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Bit, Quote) & vbCrLf
        s = s & "      ," & DbQuote(Str, IIf(Quote, "", mJob)) & vbCrLf
        s = s & "      ," & DbQuote(Str, newJob) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtNotes.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, community) & vbCrLf
        s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtLot.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtBlock.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtLotPlan.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(1, .ColIndex("value"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(2, .ColIndex("value"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(3, .ColIndex("value"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(4, .ColIndex("value"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(5, .ColIndex("value"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(6, .ColIndex("value"))) & ")"
        HFApp.SqlExec s
    

        s = ""
        s = s & "INSERT INTO tblcustomers(DivisionID,Customer_No,Job_no,Description,community,phase,purchased,approved,contract_assigned,cancelled,inactive,lot,block,lotplan,model,sale_posted)" & vbCrLf
        s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, newJob) & vbCrLf
        s = s & "      ," & DbQuote(Str, newJob) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, community) & vbCrLf
        s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
        s = s & "      ,1,1,1,0,0"
        s = s & "      ," & DbQuote(Str, txtLot.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtBlock.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtLotPlan.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
        s = s & "      ,0" & vbCrLf
        s = s & ")"
        HFApp.SqlExec s
                
    End With
    'copy over job contacts
    s = "insert into JobContacts(Job, Role, ContactID) select " & DbQuote(Str, CleanJob(newJob)) & ", Role, ContactID from JobContacts where Job = " & DbQuote(Str, CleanJob(mJob))
    HFApp.SqlExec s
        
    'Update Pm,Purchaser, Estimator on new job
    HFApp.SqlExec "Update tblJobs set Pm= (select pm from tbljobs where DivisionID = " & HFApp.DivisionID & " and isnull(pm,'')<>'' and job_no =" & DbQuote(Str, mJob) & ") where DivisionID = " & HFApp.DivisionID & " and job_no = " & DbQuote(Str, CleanJob(newJob))
    HFApp.SqlExec "Update tblJobs set Purchaser= (select purchaser from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no =" & DbQuote(Str, mJob) & ") where DivisionID = " & HFApp.DivisionID & " and job_no = " & DbQuote(Str, CleanJob(newJob))
    HFApp.SqlExec "Update tblJobs set Estimator= (select Estimator from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no =" & DbQuote(Str, mJob) & ") where DivisionID = " & HFApp.DivisionID & " and job_no = " & DbQuote(Str, CleanJob(newJob))
    
    
    'copy over the Addons except the ones that are tied to an assembly
    s = "insert into addons( Job, AssemblyID, Sequence, Title, Basis, RateType, Percentage, AmountPer1, AmountPer2, LumpSumAmt, Mat, Eq, Lab, Sub, Oth, Amount, TotalAmount, PercentOfTotal)"
    s = s & " select " & DbQuote(Str, CleanJob(newJob)) & ", AssemblyID, Sequence, Title, Basis, RateType, Percentage, AmountPer1, AmountPer2, LumpSumAmt, Mat, Eq, Lab, Sub, Oth, Amount, TotalAmount, PercentOfTotal from Addons where isnull(AssemblyID,0)=0 and Job = " & DbQuote(Str, CleanJob(mJob))
    HFApp.SqlExec s
    
    
    'copy over assemblies/items
    s = "select estassemblyid from estimateassemblies where job=" & DbQuote(Str, mJob) & " and DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        AssemblyID = Val("" & rs("EstAssemblyID"))
        s = ""
        'Used the job number field as the customer number as well. Before it was retaining the customer number from the quote, rather than the new customer which was created. This caused problems when adding assemblies afterwards because it would look up the job number from the customer
        s = s & "insert into estimateassemblies(Job,Customer_No,EstimateIndex,HFLocation,HFDescription,HFComments,HFCategory,AssemblyType,OptionType,ChangeOrder,Assembly,EstimateComplete,SalesQty,SalesRate,JCExtra,SalesWorksheet,Seq,Model,OptionID,IncludedInSpec,AssemblyUOM,DivisionID)" & vbCrLf
        s = s & "select " & DbQuote(Str, newJob) & ", " & DbQuote(Str, newJob) & ",EstimateIndex,HFLocation,HFDescription,HFComments,HFCategory,AssemblyType,OptionType,ChangeOrder,Assembly,EstimateComplete,SalesQty,SalesRate,JCExtra,SalesWorksheet,Seq,Model,OptionID,IncludedInSpec,AssemblyUOM" & vbCrLf
        s = s & ",DivisionID" & vbCrLf
        s = s & "from estimateassemblies " & vbCrLf
        s = s & "where EstAssemblyID=" & DbQuote(Num, AssemblyID) & vbCrLf
        Call HFApp.SqlExec(s)
        
        NewAssemblyID = HFApp.SqlIdentity("EstimateAssemblies")
        'copy over the Addons that are tied to the assembly
        s = "insert into addons( Job, AssemblyID, Sequence, Title, Basis, RateType, Percentage, AmountPer1, AmountPer2, LumpSumAmt, Mat, Eq, Lab, Sub, Oth, Amount, TotalAmount, PercentOfTotal)"
        s = s & " select " & DbQuote(Str, CleanJob(newJob)) & "," & NewAssemblyID & ", Sequence, Title, Basis, RateType, Percentage, AmountPer1, AmountPer2, LumpSumAmt, Mat, Eq, Lab, Sub, Oth, Amount, TotalAmount, PercentOfTotal from Addons where AssemblyID=" & AssemblyID & " and Job = " & DbQuote(Str, CleanJob(mJob))
        HFApp.SqlExec s
        
        s = ""
        s = s & "insert into estimateitems(EstAssemblyID,Job,POIndex,Phase,Item,JCExtra,JCCostCode,JCCategory,SortOrder,Description,Comments,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,PONumber,ExcludeFromPO,OriginalJCCategory,BudgetGenerated,BudgetPostingBatch,POGenBatch,BudgetDeleted,PODeleted,BudgetOverridden,POOverridden,Assembly,AssemblyDescription,Model,WBS01,WBS02,WBS03,WBS04,WBS05,WBS06,WBS07,WBS08,WBS09,WBS10,WBS11,WBS12,WBS13,WBS14,WBS15,WBS16,WBS17,WBS18,WBS19,WBS20,WBS21,WBS22,WBS23,WBS24,WBS25,WBS26,WBS27,WBS28,WBS29,WBS30,WBS31,WBS32,WBS33,WBS34,WBS35,WBS36,WBS37,WBS38,WBS39,WBS40,Location,DivisionID)" & vbCrLf
        s = s & "select " & DbQuote(Num, NewAssemblyID) & "," & DbQuote(Str, newJob) & ",POIndex,Phase,Item,JCExtra,JCCostCode,JCCategory,SortOrder,Description,Comments,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,PONumber,ExcludeFromPO,OriginalJCCategory,BudgetGenerated,BudgetPostingBatch,POGenBatch,BudgetDeleted,PODeleted,BudgetOverridden,POOverridden,Assembly,AssemblyDescription,Model,WBS01,WBS02,WBS03,WBS04,WBS05,WBS06,WBS07,WBS08,WBS09,WBS10,WBS11,WBS12,WBS13,WBS14,WBS15,WBS16,WBS17,WBS18,WBS19,WBS20,WBS21,WBS22,WBS23,WBS24,WBS25,WBS26,WBS27,WBS28,WBS29,WBS30,WBS31,WBS32,WBS33,WBS34,WBS35,WBS36,WBS37,WBS38,WBS39,WBS40,Location," & HFApp.DivisionID & vbCrLf
        s = s & "from estimateitems" & vbCrLf
        s = s & "where EstAssemblyID=" & DbQuote(Num, AssemblyID) & vbCrLf
        Call HFApp.SqlExec(s)
        
        rs.MoveNext
    Wend
    
    mJob = newJob
    txtJob.Text = mJob
    
    'save changes in case they choose not to save current quote
    Call SaveData(False)
    
    If Not Quote Then
        mMode = fceBudget
        Call Form_Load
    End If
    Call LoadJob(mJob)
    
    Screen.MousePointer = vbDefault
    

Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        MsgBox "Unable to save this quote. Number " & vbQuote & newJob & vbQuote & " has already been used." & vbCrLf & vbCrLf & "Please use a different number.", vbExclamation, App.ProductName
        Screen.MousePointer = vbDefault
        GoTo GetJob
    Else
        Call errHandler(SRCFILE & "SaveAs", s)
    End If
End Sub

Private Function ValidateJobNumber(JobNumber As String) As Boolean
    Select Case HFApp.Options(AccountingSystem)
    
        Case asTimberline
            ValidateJobNumber = IsBetween(Len(CleanJob(JobNumber)), Val(HFApp.Options(Job_Section2)) + Val(HFApp.Options(Job_Section3)), Val(HFApp.Options(Job_Section1)) + Val(HFApp.Options(Job_Section2)) + Val(HFApp.Options(Job_Section3)))
            
        Case asMasterBuilder
            If IsNumeric(JobNumber) And Val(JobNumber) - Int(Val(JobNumber)) = 0 Then
                ValidateJobNumber = True
                JobNumber = Int(Val(JobNumber))
            End If
            
        Case Else
            ValidateJobNumber = True
            
    End Select
End Function

Public Property Get Job() As String
    Job = mJob
End Property

Private Function SaveBids() As Boolean
    Dim r As Long
    Dim c As Long
    Dim s As String
    
    If mRFP = 0 Then
        SaveBids = True
        Exit Function
    End If
    
    s = ""
    s = s & "UPDATE RFPs" & vbCrLf
    s = s & "SET Description=" & DbQuote(Str, txtRFPDescription.Text) & vbCrLf
    s = s & "   ,Comments=" & DbQuote(Str, txtRFPComments.Text) & vbCrLf
    s = s & "WHERE RFP=" & DbQuote(Num, mRFP) & vbCrLf
    Call HFApp.SqlExec(s, dbHomefront)
    
    
    With gBids
    For r = 1 To .Rows - 1
    If .RowData(r) = "DIRTY" Then
    
        s = ""
        s = s & "update vendorbids" & vbCrLf
        s = s & "   set ExpiryDate=" & DbQuote(Date, .Cell(flexcpText, r, .ColIndex("ExpiryDate"))) & vbCrLf
        s = s & "      ,Comments=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("Comments"))) & vbCrLf
        s = s & " where RFP=" & DbQuote(Num, mRFP) & vbCrLf
        s = s & "   and Vendor=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor"))) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
    
    End If
    Next
    End With
    
    
    With gBidItems
    For r = 1 To .Rows - 1
    For c = 0 To .Cols - 1
    If .ColData(c) = "VendorRate" Then
    
        s = ""
        s = s & "update VendorBidItems" & vbCrLf
        s = s & "   set Status=" & DbQuote(Num, .Cell(flexcpData, r, c)) & vbCrLf
        s = s & "      ,comment=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("vdC~" & Mid(gBidItems.ColKey(c), 5)))) & vbCrLf
        If .Cell(flexcpText, r, c) = "" Then
            s = s & "      ,Rate = null" & vbCrLf
        Else
            s = s & "      ,Rate = " & DbQuote(Num, .Cell(flexcpText, r, c)) & vbCrLf
        End If
        s = s & " where RFP=" & DbQuote(Num, mRFP) & vbCrLf
        s = s & "   and Vendor=" & DbQuote(Str, Mid(gBidItems.ColKey(c), 5)) & vbCrLf
        s = s & "   and EstItemID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("EstItemID"))) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        
        
        If .Cell(flexcpData, r, c) = bsAccepted Then
            s = ""
            s = s & "update estimateitems" & vbCrLf
            s = s & "set budgetvendor=" & DbQuote(Str, Mid(gBidItems.ColKey(c), 5)) & vbCrLf
            s = s & "   ,budgetrate=" & DbQuote(Num, .Cell(flexcpText, r, c)) & vbCrLf
            s = s & "   ,budgetpretax=round(budgetqty*" & DbQuote(Num, .Cell(flexcpText, r, c)) & ",2)" & vbCrLf
            s = s & "   ,budgetjctax=round(round(budgetqty*" & DbQuote(Num, .Cell(flexcpText, r, c)) & ",2) * budgetjctaxrate/100,2)" & vbCrLf
            s = s & "   ,budgetnjctax=round(round(budgetqty*" & DbQuote(Num, .Cell(flexcpText, r, c)) & ",2) * budgetnjctaxrate/100,2)" & vbCrLf
            s = s & "where EstItemID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("EstItemID"))) & vbCrLf
            s = s & "  and budgetgenerated=0" & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
        
            s = ""
            s = s & "update estimateitems" & vbCrLf
            s = s & "set POvendor=" & DbQuote(Str, Mid(gBidItems.ColKey(c), 5)) & vbCrLf
            s = s & "   ,POrate=" & DbQuote(Num, .Cell(flexcpText, r, c)) & vbCrLf
            s = s & "   ,POpretax = round(POqty*" & DbQuote(Num, .Cell(flexcpText, r, c)) & ",2)" & vbCrLf
            s = s & "   ,POjctax = round(round(POqty*" & DbQuote(Num, .Cell(flexcpText, r, c)) & ",2) * POjctaxrate/100,2)" & vbCrLf
            s = s & "   ,POnjctax = round(round(POqty*" & DbQuote(Num, .Cell(flexcpText, r, c)) & ",2) * POnjctaxrate/100,2)" & vbCrLf
            s = s & "where EstItemID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("EstItemID"))) & vbCrLf
            s = s & "  and pogenbatch=0" & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
       End If
        
    End If
    Next
    Next
    End With
    
    
    SaveBids = True
    
End Function


Public Sub mnuEstimateBidItemsGridSub_Click(Index As Integer)
    Dim s As String
    Dim r As Long
    Dim c As Long


    With gBidItems
        If .ColData(.Col) <> "VendorRate" Then Exit Sub
    
        Select Case Index
            Case mcBIDITEM_EXPORT
                Call ExportBidSheets
                
                    
            Case mcBIDITEM_IMPORT
                Call ImportBidSheet
        
            
            Case mcBIDITEM_FORMATTING:
                Call FEstimateItemsFormatting.ShowBidRules(AcceptedForeColor, AcceptedBackColor, AcceptedStyle, DeclinedForeColor, DeclinedBackColor, DeclinedStyle, MinForeColor, MinBackColor, MinStyle, MaxForeColor, MaxBackColor, MaxStyle)
                Call UpdateBidTotals
            
            
            Case mcBIDITEM_COMMENTS:
                On Error Resume Next
                s = .TextMatrix(.Row, .ColIndex("vdC~" & Mid(gBidItems.ColKey(gBidItems.Col), 5)))
                If FComments.Edit(s, gBidItems, , , 4000) Then
                    Dirty = True
                    s = Trim(s)
                    If Replace(s, vbCrLf, "") = "" Then s = ""
                    .Cell(flexcpPicture, .Row, .Col) = IIf(s = "", Nothing, FMain.SmallIcons.ListImages("cellcomments").Picture)
                    .Cell(flexcpPictureAlignment, .Row, .Col) = flexPicAlignRightTop
                    .TextMatrix(.Row, .ColIndex("vdC~" & Mid(gBidItems.ColKey(gBidItems.Col), 5))) = s
                End If
            
            
            Case mcBIDITEM_ACCEPTED
                    For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        If .Cell(flexcpText, r, .Col) <> "" Then
                            If .Cell(flexcpData, r, .Col) <> bsDeclined Then
                               .Cell(flexcpData, r, .Col) = IIf(FMain.mnuEstimateBidItemsGridSub(Index).checked, bsProposed, bsAccepted)
                            End If
                            If .Cell(flexcpData, r, .Col) = bsAccepted Then
                                For c = 0 To .Cols - 1
                                    Select Case .Cell(flexcpData, r, c)
                                        Case bsDeclined: 'do nothing
                                        Case Else:       .Cell(flexcpData, r, c) = IIf(c = .Col, bsAccepted, bsProposed)
                                    End Select
                                Next
                            End If
                        End If
                    Next
                    Dirty = True
                
                
            Case mcBIDITEM_DECLINED
                For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                For c = Min(.Col, .ColSel) To Max(.Col, .ColSel)
                    If FMain.mnuEstimateBidItemsGridSub(Index).checked Then
                        If .Cell(flexcpData, r, .Col) = bsDeclined Then .Cell(flexcpData, r, .Col) = bsProposed
                    Else
                        .Cell(flexcpData, r, .Col) = bsDeclined
                    End If
                Next
                Next
                Dirty = True
            
            
            
            
            Case mcBIDITEM_VENDOR
                Call HFApp.EditVendor(Mid(gBidItems.ColKey(gBidItems.Col), 5))
                
                
        End Select
        UpdateBidTotals
    End With
End Sub


Private Sub UpdateBidTotals()
On Error GoTo ExitSub
    Dim r As Long
    Dim c As Long
    Dim accepted As Boolean
    Dim vMin As Double
    Dim vMax As Double
    Dim vTotal As Double
    Dim vCount As Long
    
    
    gBids.Redraw = flexRDNone
    gBidItems.Redraw = flexRDNone
    
    With gBidItems
        
        .Cell(flexcpFontBold, 1, 0, .Rows - 1, .Cols - 1) = False
        .Cell(flexcpFontItalic, 1, 0, .Rows - 1, .Cols - 1) = False
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, .Cols - 1) = vbWindowText
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = vbWindowBackground
        
        'for each bid item
        For r = 1 To .Rows - 1
            
            .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbWindowBackground
            vMin = 99999999
            vMax = -99999999
            For c = 1 To .Cols - 1
                If .ColData(c) = "VendorRate" And .TextMatrix(r, c) <> "" And .Cell(flexcpData, r, c) <> bsDeclined Then
                    vMax = Max(vMax, .ValueMatrix(r, c))
                    vMin = Min(vMin, .ValueMatrix(r, c))
                End If
            Next
            
            'color each cell
            For c = 1 To .Cols - 1
                If .ColData(c) = "VendorRate" Then
            
                    Select Case True
                        'accepted
                        Case .Cell(flexcpData, r, c) = bsAccepted
                            .Cell(flexcpForeColor, r, c) = AcceptedForeColor
                            .Cell(flexcpBackColor, r, c) = AcceptedBackColor
                            .Cell(flexcpFontBold, r, c) = InStr(1, AcceptedStyle, "Bold")
                            .Cell(flexcpFontItalic, r, c) = InStr(1, AcceptedStyle, "Italic")
                            
                            .Cell(flexcpForeColor, r, .ColIndex("ItemDesc")) = AcceptedForeColor
                            .Cell(flexcpBackColor, r, .ColIndex("ItemDesc")) = AcceptedBackColor
                            .Cell(flexcpFontBold, r, .ColIndex("ItemDesc")) = InStr(1, AcceptedStyle, "Bold")
                            .Cell(flexcpFontItalic, r, .ColIndex("ItemDesc")) = InStr(1, AcceptedStyle, "Italic")
                
                        'declined
                        Case .Cell(flexcpData, r, c) = bsDeclined
                            .Cell(flexcpForeColor, r, c) = DeclinedForeColor
                            .Cell(flexcpBackColor, r, c) = DeclinedBackColor
                            .Cell(flexcpFontBold, r, c) = InStr(1, DeclinedStyle, "Bold")
                            .Cell(flexcpFontItalic, r, c) = InStr(1, DeclinedStyle, "Italic")
                        
                        'min/max
                        Case .TextMatrix(r, c) <> ""
                            If vMax = .ValueMatrix(r, c) Then
                                .Cell(flexcpForeColor, r, c) = MaxForeColor
                                .Cell(flexcpBackColor, r, c) = MaxBackColor
                                .Cell(flexcpFontBold, r, c) = InStr(1, MaxStyle, "Bold")
                                .Cell(flexcpFontItalic, r, c) = InStr(1, MaxStyle, "Italic")
                            End If
                            
                            If vMin = .ValueMatrix(r, c) Then
                                .Cell(flexcpForeColor, r, c) = MinForeColor
                                .Cell(flexcpBackColor, r, c) = MinBackColor
                                .Cell(flexcpFontBold, r, c) = InStr(1, MinStyle, "Bold")
                                .Cell(flexcpFontItalic, r, c) = InStr(1, MinStyle, "Italic")
                            End If
                            
                    End Select
                End If
            Next
        Next
           
        'update vendor totals
        For c = 1 To .Cols - 1
            If .ColData(c) = "VendorRate" Then
                vTotal = 0
                vCount = 0
                accepted = False
                For r = 1 To .Rows - 1
                    accepted = accepted Or .Cell(flexcpData, r, c) = bsAccepted
                    If .Cell(flexcpData, r, c) <> bsDeclined Then
                        vTotal = vTotal + .ValueMatrix(r, c) * .ValueMatrix(r, .ColIndex("POQty"))
                        If .TextMatrix(r, c) <> "" Then
                            vCount = vCount + 1
                        End If
                    End If
                Next
        
                r = gBids.FindRow(Mid(.ColKey(c), 5), , gBids.ColIndex("Vendor"), True, True)
                If r <> -1 Then
                    gBids.TextMatrix(r, gBids.ColIndex("amount")) = IIf(vTotal = 0, "", format(vTotal, "#,##0.00"))
                    gBids.TextMatrix(r, gBids.ColIndex("status")) = IIf(vCount = 0, "", IIf(vCount = .Rows - 1, "Full", "Partial"))
                
                    
                    gBids.Cell(flexcpForeColor, r, gBids.ColIndex("VendorName")) = IIf(accepted, AcceptedForeColor, vbWindowText)
                    gBids.Cell(flexcpBackColor, r, gBids.ColIndex("VendorName")) = IIf(accepted, AcceptedBackColor, vbWindowBackground)
                    gBids.Cell(flexcpFontBold, r, gBids.ColIndex("VendorName")) = IIf(accepted, InStr(1, AcceptedStyle, "Bold"), False)
                    gBids.Cell(flexcpFontItalic, r, gBids.ColIndex("VendorName")) = IIf(accepted, InStr(1, AcceptedStyle, "Italic"), False)
                                
                End If
            End If
        Next
    
    End With
    
'    'show min/max on vendor totals
'    With gBids
'        vMin = 99999999
'        vMax = -99999999
'        c = .ColIndex("Amount")
'        .Cell(flexcpBackColor, 1, c, .Rows - 1, c) = vbWindowBackground
'        For r = 1 To .Rows - 1
'            If .TextMatrix(r, c) <> "" Then vMax = Max(vMax, .ValueMatrix(r, c))
'            If .TextMatrix(r, c) <> "" Then vMin = Min(vMin, .ValueMatrix(r, c))
'        Next
'        For r = 1 To .Rows - 1
'            If vMax = .ValueMatrix(r, c) Then .Cell(flexcpBackColor, r, c) = MAXCOLOR
'            If vMin = .ValueMatrix(r, c) Then .Cell(flexcpBackColor, r, c) = MINCOLOR
'        Next
'    End With
'
ExitSub:
    gBids.Redraw = flexRDBuffered
    gBidItems.Redraw = flexRDBuffered
End Sub


Public Sub mnuEstimateBidsGridSub_Click(Index As Integer)
    Dim s As String
    Dim i As Long
    Dim r As Long
    Dim v As String
    
    With gBids
    Select Case Index
            
        
        Case mcBID_ADD
            If Not SaveData(False) Then Exit Sub

            s = "select vendorgroupid [Trade], vendor_id Vendor,Vendor_name Company from tblvendors where DivisionID =" & HFApp.DivisionID & " and inactive=0"
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, , , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Vendor", ""), True) Then Exit Sub
            For i = 1 To FPickList.SelectedItems
                v = v & "," & DbQuote(Str, FPickList.SelectedItem("Vendor", i))
            Next
            v = Mid(v, 2)

            s = ""
            s = s & "INSERT INTO VendorBids(RFP,Vendor,CreatedDate) " & vbCrLf
            s = s & "SELECT " & DbQuote(Num, mRFP) & ",vendor_id,getdate() " & vbCrLf
            s = s & "  FROM tblvendors" & vbCrLf
            s = s & " WHERE DivisionID =" & HFApp.DivisionID & " and vendor_id in(" & v & ")" & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)

            s = ""
            s = s & "INSERT INTO VendorBidItems(RFP,Vendor,EstItemID)" & vbCrLf
            s = s & "SELECT e.rfp" & vbCrLf
            s = s & "      ,v.vendor_id" & vbCrLf
            s = s & "      ,e.estitemid" & vbCrLf
            s = s & "  FROM estimateitems e" & vbCrLf
            s = s & "       JOIN tblvendors v ON(1=1)" & vbCrLf
            s = s & " WHERE v.DivisionID =" & HFApp.DivisionID & " and vendor_id in(" & v & ")" & vbCrLf
            s = s & "   AND e.rfp=" & DbQuote(Num, mRFP) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)

            Call LoadRFQ
            
        
        Case mcBID_REMOVE
            If vbYes <> MsgBox("Are you sure you want to remove " & IIf(.SelectedRows > 1, "these bidders?", "this bidder?"), vbQuestion + vbYesNo, App.ProductName) Then Exit Sub
            If Not SaveData(False) Then Exit Sub
            For i = 0 To .SelectedRows - 1
                r = .SelectedRow(i)
                s = "delete from vendorbiditems where vendor=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor"))) & " and rfp=" & DbQuote(Num, mRFP)
                Call HFApp.SqlExec(s, dbHomefront)
                s = "delete from vendorbids where vendor=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor"))) & " and rfp=" & DbQuote(Num, mRFP)
                Call HFApp.SqlExec(s, dbHomefront)
            Next
            Call LoadRFQ
        
            
        
        Case mcBID_VENDOR
            Call HFApp.EditVendor(.TextMatrix(.Row, .ColIndex("Vendor")))
        
        Case mcBID_IMPORT
            Call ImportBidSheet
        
        Case mcBID_EXPORT
            Call ExportBidSheets
        
    End Select
    End With
End Sub


Private Sub ImportBidSheet()
    Dim FileName As String
    Dim r As Long
    Dim c As Long
    Dim IDCol   As Long
    Dim RateCol As Long
    Dim s As String
    
    If Not VBGetOpenFileName(FileName, , , , , True, "Microsoft Excel File (*.xls)|*.xls", , , , "xls", Me.hwnd) Then Exit Sub
    
    With gImport
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        If InIde Then .Visible = True
        Call .LoadGrid(FileName, flexFileExcel)
        
        IDCol = -1
        RateCol = -1
        For r = 1 To .Rows - 1
        
            .Row = r
            .Col = 1
            If IDCol = -1 Then
                For c = 1 To .Cols - 1
                    .Col = c
                    Select Case Trim(UCase(.TextMatrix(r, c)))
                        Case "REF#":                                      IDCol = c
                        Case "RATE", "UNIT RATE", "UNIT PRICE", "PRICE":  RateCol = c
                    End Select
                Next
                If IDCol = -1 And RateCol <> -1 Then
                    MsgBox "Invalid format" & vbCrLf & vbCrLf & "The Ref# column could not be found on row " & r, vbExclamation, App.ProductName
                    Exit Sub
                ElseIf RateCol = -1 And IDCol <> -1 Then
                    MsgBox "Invalid format" & vbCrLf & vbCrLf & "The Unit Price column could not be found on row " & r, vbExclamation, App.ProductName
                    Exit Sub
                End If
            Else
                If Trim(.TextMatrix(r, RateCol)) = "" Then
                    s = "update vendorbiditems set rate=null where vendorbiditemid=" & DbQuote(Num, .TextMatrix(r, IDCol))
                Else
                    s = "update vendorbiditems set rate=" & DbQuote(Num, .TextMatrix(r, RateCol)) & " where vendorbiditemid=" & DbQuote(Num, .TextMatrix(r, IDCol))
                End If
                Call HFApp.SqlExec(s, dbHomefront)
            End If
        Next
        .Visible = False
    End With
    Call LoadRFQ
    
End Sub

Private Sub ExportBidSheet(RFP As Long, Vendor As String, FileName As String)
On Error GoTo eh
    
    Dim RptFile  As String
    Dim c As New ZybUtil.Crystal
    Dim i        As Long
    Dim s           As String
    
    
    RptFile = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\rfq.rpt")
    Call c.LoadODBCReport(RptFile, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
    On Error Resume Next
    Call c.ParameterValue("DivisionID", HFApp.DivisionID)
    Call c.ParameterValue("RFP", RFP)
    Call c.ParameterValue("Vendor", Vendor)
    On Error GoTo 0
    
    
    
    If FileExists(FileName) Then Call Kill(FileName)
    Call c.SaveAsXLS(FileName)
   
Exit Sub
eh: Call errHandler(SRCFILE & "SendGroup")
End Sub



Public Sub mnuEstimateRFPItemsGridSub_Click(Index As Integer)
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    Dim ids As String
    
    Select Case Index
        Case mcRFPITEM_ADD
            'build picklist
            s = "EstItemID,EstPhase Phase,EstItem Item,ItemDesc Description,POIndex,Location,HFDescription"
            Set rs = HFApp.SqlExec("select * from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mJob))
            If Not rs.EOF Then
                For r = 1 To 40
                    If "" & rs("WBSDesc" & format(r, "00")) <> "" Then s = s & ",WBS" & format(r, "00") & " " & Quote(rs("WBSDesc" & format(r, "00")))
                Next
            End If
            s = "select " & s & " from estimateditems where rfp=0 and DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mJob)
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s, , , , , "EstItemID", True) Then Exit Sub
            
            Call SaveData(False)
            
            'build selected items list
            For r = 1 To FPickList.SelectedItems
                ids = ids & "," & DbQuote(Num, FPickList.SelectedItem("EstItemID", r))
            Next
            ids = Mid(ids, 2)
            
            'add selected items to rfp
            s = "update estimateitems set rfp=" & DbQuote(Num, mRFP) & " where estitemid in(" & ids & ")"
            Call HFApp.SqlExec(s)
            
            'add selected items to vendor bids
            s = ""
            s = s & "INSERT INTO VendorBidItems(RFP,Vendor,EstItemID)" & vbCrLf
            s = s & "select b.rfp,b.vendor,i.estitemid" & vbCrLf
            s = s & "  from vendorbids b" & vbCrLf
            s = s & "       join estimateitems i on (1=1)" & vbCrLf
            s = s & " where b.rfp=" & DbQuote(Num, mRFP) & vbCrLf
            s = s & "   and i.estitemid in(" & ids & ")" & vbCrLf
            Call HFApp.SqlExec(s)
            
            Call LoadRFQ
            
            
        Case mcRFPITEM_REMOVE
            With gBidItems
                If vbYes <> MsgBox("Are you sure you want to remove this item from the RFP?", vbQuestion + vbYesNo, App.ProductName) Then Exit Sub
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    If .Rows = 2 Then
                        If vbYes <> MsgBox("This is the last item on the RFP. If you remove it the RFP will also be removed." & vbCrLf & vbCrLf & "Are you sure this is what you want to do?", vbExclamation + vbYesNo, App.ProductName) Then Exit Sub
                        Call HFApp.SqlExec("delete from vendorbiditems where rfp=" & DbQuote(Num, mRFP))
                        Call HFApp.SqlExec("delete from vendorbids where rfp=" & DbQuote(Num, mRFP))
                        Call HFApp.SqlExec("delete from rfps where rfp=" & DbQuote(Num, mRFP))
                        Call HFApp.SqlExec("update estimateitems set rfp=0 where rfp=" & DbQuote(Num, mRFP))
                        Call mnuEstimateItemViewsSub_Click(1 + mViewIndex)
                    Else
                        Call HFApp.SqlExec("delete from vendorbiditems where estitemid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("EstItemID"))))
                        Call HFApp.SqlExec("update estimateitems set rfp=0 where estitemid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("EstItemID"))))
                        .RemoveItem
                    End If
                Next
            End With
    End Select
End Sub



Private Sub Clear()
    mRFP = 0
    txtRFPDescription.Text = ""
    txtRFPComments.Text = ""
    gBids.Rows = 1
    gBidItems.Rows = 1
    frmAssembly.Visible = False
    Toolbar.Buttons("SaveAssembly").Enabled = False
    frmPOIndex.Visible = False
    gItems.Rows = 1
    gBidItems.Rows = 1
    gBids.Rows = 1
    Call Form_Resize
End Sub


Private Sub ReloadAssemblyTotals(Row As Long)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    With gAssemblies
    If mViews(mViewIndex).Name = "Contract" Then Exit Sub
    
        If Row = -1 Then Row = .Row
        If Row < 1 Then Exit Sub
    
        s = ""
        s = s & "SELECT SUM(CASE WHEN IsChange=0 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) Budgeted" & vbCrLf
        s = s & "      ,SUM(CASE WHEN IsChange=1 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) Changes" & vbCrLf
        s = s & "      ,SUM(CASE WHEN POGenBatch=0 or poDeleted=1 THEN 0 ELSE POPretax+POJCTax END) Commited" & vbCrLf
        s = s & "  FROM EstimatedItems i" & vbCrLf
        s = s & "WHERE " & Mid(.RowData(Row), 6)
        Set rs = HFApp.SqlExec(s, dbHomefront)
        If Not rs.EOF Then
            .Cell(flexcpText, Row, .ColIndex("Budgeted")) = format(Val("" & rs("budgeted")), "#,###.00")
            .Cell(flexcpText, Row, .ColIndex("Changes")) = format(Val("" & rs("Changes")), "#,###.00")
            .Cell(flexcpText, Row, .ColIndex("Commited")) = format(Val("" & rs("commited")), "#,###.00")
            
        End If
        
        r = .GetNodeRow(Row, flexNTParent)
        
        If r >= 0 And mViews(mViewIndex).Name <> "Contract" Then Call ReloadAssemblyTotals(r)
                    
    End With
    
End Sub

Public Sub mnuEstimateRFPsSub_Click(Index As Integer)
    Select Case Index
        Case mcRFP_DELETE: Call DeleteRFQ
    End Select
End Sub

Private Sub DeleteRFQ()
    Dim i As Long
    With gAssemblies
        i = Val(Parse(.Cell(flexcpText, .Row, 1), 1, Chr(2)))
        If i = 0 Then Exit Sub
        If vbOK <> MsgBox("Are you sure you want to delete this RFQ?", vbOKCancel + vbQuestion, App.ProductName) Then Exit Sub
        Call HFApp.SqlExec("exec dbo.Purch_DeleteRFQ " & DbQuote(Num, i))
        Call .RemoveItem(.Row)
        Call LoadRFQ
    End With
End Sub


Private Sub ExportBidSheets()
    Dim i As Long
    Dim r As Long
    Dim Folder As String
    Dim FileName As String
    Dim Vendor As String
    
    
    With gBids
    
        If Not VBChooseFolder(0, BIF_RETURNONLYFSDIRS, Folder, "Select Export Location", Me.hwnd) Then Exit Sub
    
        For i = 0 To .SelectedRows - 1
            r = .SelectedRow(i)
            Vendor = .TextMatrix(r, .ColIndex("Vendor"))
            FileName = PathAppend(Folder, CleanFileName(.TextMatrix(r, .ColIndex("VendorName"))) & ".xls")
            Call ExportBidSheet(mRFP, Vendor, FileName)
        Next
    
    End With

End Sub


Private Function CheckVendorInsurance(WhereClause As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    Dim things As String
    Dim msg As String
    
    If mMode <> fcePO Then
        CheckVendorInsurance = vbOK = MsgBox("Generate these budgets?", vbOKCancel + vbQuestion, App.ProductName)
        Exit Function
    End If
    
    s = ""
    s = s & "select distinct v.vendor_name Name" & vbCrLf
    s = s & "      ,case when v.glinsrequired=1 and v.glinsexpdate<getdate() then v.glinsexpdate else null end gl" & vbCrLf
    s = s & "      ,case when v.wcinsrequired=1 and v.wcinsexpdate<getdate() then v.wcinsexpdate else null end wc" & vbCrLf
    s = s & "      ,case when v.umbinsrequired=1 and v.umbinsexpdate<getdate() then v.umbinsexpdate else null end umb" & vbCrLf
    s = s & "      ,case when v.autoinsrequired=1 and v.autoinsexpdate<getdate() then v.autoinsexpdate else null end auto" & vbCrLf
    s = s & "from estimateditems i" & vbCrLf
    s = s & "join tblvendors v on (i.povendor=v.vendor_id and v.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
    s = s & "where ((v.glinsrequired=1 and v.glinsexpdate<getdate()) or" & vbCrLf
    s = s & "       (v.wcinsrequired=1 and v.wcinsexpdate<getdate()) or" & vbCrLf
    s = s & "       (v.umbinsrequired=1 and v.umbinsexpdate<getdate()) or" & vbCrLf
    s = s & "       (v.autoinsrequired=1 and v.autoinsexpdate<getdate()))" & vbCrLf
    s = s & "   and " & WhereClause
    s = s & " order by 1"
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        CheckVendorInsurance = vbOK = MsgBox("Generate these " & IIf(mMode = fceBudget, "budgets?", "purchase orders?"), vbOKCancel + vbQuestion, App.ProductName)
        Exit Function
    End If
    
    
    s = "These vendors have expired insurance policies. Do you want to generate these " & IIf(mMode = fceBudget, "budgets", "purchase orders") & " anyway?" & vbCrLf & vbCrLf
    While Not rs.EOF
        s = s & "" & rs("name") & vbCrLf
        If "" & rs("gl") <> "" Then s = s & "    General liablity insurance expired " & format("" & rs("gl"), "mmm d, yyyy") & vbCrLf
        If "" & rs("wc") <> "" Then s = s & "    Workers compensation insurance expired " & format("" & rs("wc"), "mmm d, yyyy") & vbCrLf
        If "" & rs("umb") <> "" Then s = s & "    Umbrella insurance expired " & format("" & rs("umb"), "mmm d, yyyy") & vbCrLf
        If "" & rs("auto") <> "" Then s = s & "    Automobile insurance expired " & format("" & rs("auto"), "mmm d, yyyy") & vbCrLf
        s = s & vbCrLf
        rs.MoveNext
    Wend
    
    CheckVendorInsurance = vbOK = MsgBox(s, vbOKCancel + vbQuestion, App.ProductName)

End Function

Private Sub BudgetsAreLocked(Locked As Boolean, LockedDate As String, LockedBy As String)

    If mMode = fceBudget Then
        If Locked Then
            lblBudgetLock.Caption = "Budgets finalized " & format(LockedDate, "mmm dd/yy") & " by " & LockedBy
        Else
            lblBudgetLock.Caption = "Budgets are open. Changes are permitted."
        End If
        imgLockBudgets(0).Visible = Not Locked
        imgLockBudgets(1).Visible = Locked
        Toolbar.Buttons("TakeoffOneTime").Enabled = Not Locked
        Toolbar.Buttons("TakeoffItem").Enabled = Not Locked
        Toolbar.Buttons("TakeoffAssembly").Enabled = Not Locked
        Toolbar.Buttons("TakeoffCustom").Enabled = Not Locked
        Toolbar.Buttons("RePrice").Enabled = Not Locked
    Else
        lblBudgetLock.Visible = False
        imgLockBudgets(0).Visible = False
        imgLockBudgets(1).Visible = False
    End If
End Sub

Private Sub LockMeAndMyChildren(Row As Long, Locked As Boolean)
    Dim n As VSFlexNode
    With gAssemblies

        'lock me
        .Cell(flexcpText, Row, .ColIndex("BudgetsLocked")) = IIf(Locked, 1, 0)
        .Cell(flexcpForeColor, Row, 4) = IIf(Locked, vbGrayText, vbWindowText)
    
        Set n = .GetNode(Row).GetNode(flexNTFirstChild)
        While Not n Is Nothing
            'lock this child
            .Cell(flexcpText, n.Row, .ColIndex("BudgetsLocked")) = IIf(Locked, 1, 0)
            .Cell(flexcpForeColor, n.Row, 4) = IIf(Locked, vbGrayText, vbWindowText)
            'lock this child's children
            Call LockMeAndMyChildren(n.Row, Locked)
            'next child
            Set n = n.GetNode(flexNTNextSibling)
        Wend
    End With
End Sub



Private Sub SplitItems()
On Error GoTo eh
    Dim Percents() As Double
    Dim WriteUnits As Boolean
    Dim UnitStart As Long
    Dim UnitIncrement As Long
    
    Dim r As Long
    Dim i As Long
    Dim c As Long
    Dim NewRow As Long
    
    If Not FSplitItems.ShowForm(Percents(), WriteUnits, UnitStart, UnitIncrement) Then Exit Sub
    
Dim t As Single
t = Timer()
mSplitting = True

    With gItems
    For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
        'for each copy to be made
        For i = 1 To UBound(Percents)
            Dirty = True
            
            'create new row. copy all values from this row
            NewRow = Max(.Row, .RowSel) + 1
            .AddItem "", NewRow
            .RowData(NewRow) = "NEW"
            For c = 0 To .Cols - 1
                .TextMatrix(NewRow, c) = .TextMatrix(r, c)
            Next
            
            If WriteUnits Then
                .TextMatrix(NewRow, .ColIndex("Unit")) = UnitStart + ((i - 1) * UnitIncrement)
            End If
            
            'change id of new item
            .TextMatrix(NewRow, .ColIndex("EstItemID")) = "0"
            If mMode = fcePO Then
                .TextMatrix(NewRow, .ColIndex("POQty")) = .ValueMatrix(r, .ColIndex("POQty")) * Percents(i)
                Call .Select(NewRow, .ColIndex("POQty"))
                Call gItems_AfterEdit(NewRow, .ColIndex("POQty"))
                .TextMatrix(NewRow, .ColIndex("BudgetGenerated")) = "False"
                .TextMatrix(NewRow, .ColIndex("BudgetPostingBatch")) = "0"
                .TextMatrix(NewRow, .ColIndex("BudgetDeleted")) = "True"
                .TextMatrix(NewRow, .ColIndex("BudgetQty")) = 0
                .TextMatrix(NewRow, .ColIndex("BudgetRate")) = 0
            Else
                .TextMatrix(NewRow, .ColIndex("BudgetQty")) = .ValueMatrix(r, .ColIndex("BudgetQty")) * Percents(i)
                If .TextMatrix(r, .ColIndex("PONumber")) <> "" Then
                    .TextMatrix(NewRow, .ColIndex("PODeleted")) = "True"
                    .TextMatrix(NewRow, .ColIndex("POQty")) = 0
                    .TextMatrix(NewRow, .ColIndex("PORate")) = 0
                End If
                Call .Select(NewRow, .ColIndex("BudgetQty"))
                Call gItems_AfterEdit(NewRow, .ColIndex("BudgetQty"))
            End If
            
            Call ColorizeItems(NewRow)
        Next
        
        
        'delete source po item
        .RowHidden(r) = True
        .RowData(r) = "DIRTY"
        If mMode = fcePO Then
            .TextMatrix(r, .ColIndex("PODeleted")) = "True"
            .TextMatrix(r, .ColIndex("POQty")) = 0
            .TextMatrix(r, .ColIndex("PORate")) = 0
        Else
            If .TextMatrix(r, .ColIndex("PONumber")) = "" Then
                .TextMatrix(r, .ColIndex("PODeleted")) = "True"
                .TextMatrix(r, .ColIndex("POQty")) = 0
                .TextMatrix(r, .ColIndex("PORate")) = 0
            End If
            .TextMatrix(r, .ColIndex("BudgetDeleted")) = "True"
            .TextMatrix(r, .ColIndex("BudgetQty")) = 0
            .TextMatrix(r, .ColIndex("BudgetRate")) = 0
        End If
    Next
    End With


mSplitting = False
Call GroupGrid
'Call MsgBox(Timer() - t, "0.000")

Exit Sub
eh: MsgBox Err.Description:
Resume Next
End Sub



Private Sub LoadContract()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    
    s = ""
    s = s & "select" & vbCrLf
    s = s & "  b.BillingItemID" & vbCrLf
    s = s & " ,b.Description" & vbCrLf
    s = s & " ,SUM(CASE WHEN IsChange=0 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) Budget" & vbCrLf
    s = s & " ,SUM(CASE WHEN IsChange=1 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) Changes" & vbCrLf
    s = s & " ,SUM(CASE WHEN POGenBatch=0 or poDeleted=1 THEN 0 ELSE POPretax+POJCTax END) Commited" & vbCrLf
    s = s & " ,0 JTDCost" & vbCrLf
    s = s & " ,b.Price" & vbCrLf
    s = s & " ,b.amtbilled Billed" & vbCrLf
    s = s & " ,b.HoldbackRate" & vbCrLf
    s = s & " ,b.HoldbackTotalHeld" & vbCrLf
    s = s & " ,b.TaxGroup" & vbCrLf
    s = s & " ,a.account RevenueAccount" & vbCrLf
    s = s & " ,a.description RevenueAccountDescription" & vbCrLf
    
    s = s & "from billingitems b" & vbCrLf
    s = s & "left outer join estimateditems i on(b.BillingItemID=i.BillingItemID)" & vbCrLf
    s = s & "left outer join glaccounts a on(i.DivisionID = a.DivisionID and b.revenueaccount=a.account)" & vbCrLf
    
    s = s & "where b.job=" & DbQuote(Str, mJob) & vbCrLf
    s = s & "group by b.BillingItemID,b.Description,b.CostCode,b.Category,b.Price,b.SortOrder,b.amtbilled,b.holdbackrate" & vbCrLf
    s = s & "        ,b.holdbacktotalheld,b.TaxGroup,a.Account,a.description" & vbCrLf
    s = s & "order by b.SortOrder,b.Description" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    With gContract
        .Rows = 1
        r = 0
        While Not rs.EOF
            .AddItem ""
            r = r + 1
            .TextMatrix(r, .ColIndex("BillingItemID")) = "" & rs("BillingItemID")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("Budget")) = Val("" & rs("Budget"))
            .TextMatrix(r, .ColIndex("Changes")) = Val("" & rs("Changes"))
            .TextMatrix(r, .ColIndex("Committed")) = Val("" & rs("Commited"))
            .TextMatrix(r, .ColIndex("JTDCost")) = Val("" & rs("JTDCost"))
            .TextMatrix(r, .ColIndex("Price")) = Val("" & rs("Price"))
            .TextMatrix(r, .ColIndex("Billed")) = Val("" & rs("Billed"))
            .TextMatrix(r, .ColIndex("TaxGroup")) = "" & rs("TaxGroup")
            If Val("" & rs("Price")) = 0 Then
                .TextMatrix(r, .ColIndex("PercentComplete")) = 0
            Else
                .TextMatrix(r, .ColIndex("PercentComplete")) = Val("" & rs("Billed")) / Val("" & rs("Price"))
            End If
            .TextMatrix(r, .ColIndex("Remaining")) = Val("" & rs("Price")) - Val("" & rs("Billed"))
            
            .TextMatrix(r, .ColIndex("HoldbackRate")) = Val("" & rs("HoldbackRate"))
            .TextMatrix(r, .ColIndex("HoldbackTotalHeld")) = Val("" & rs("HoldbackTotalHeld"))
            .TextMatrix(r, .ColIndex("RevenueAccount")) = "" & rs("RevenueAccount")
            .TextMatrix(r, .ColIndex("RevenueAccountDescription")) = "" & rs("RevenueAccountDescription")
            
            rs.MoveNext
        Wend
        .AddItem ""
    End With
    
    Call LoadInvoiceHistory
    
    
End Sub



Private Sub LoadInvoiceHistory()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    s = ""
    s = s & "select m.arcustomer,m.invoice,isnull(m.RetainageInvoice,0) RetainageInvoice,m.invoice,m.invoicedate,m.description,m.draw,m.notes,m.scopeofwork,m.PostedDate,m.PostedBy,m.Cancelled,m.CancelledBy" & vbCrLf
    s = s & "      ,sum(d.amtthisperiod) Amount" & vbCrLf
    s = s & "      ,sum(d.holdbackamount) Holdback" & vbCrLf
    s = s & "      ,sum(case when m.RetainageInvoice =1 then d.holdbackamount else 0 end) HoldbackBilled" & vbCrLf
    s = s & "from arinvoices m" & vbCrLf
    s = s & "join arinvoiceitems d on(m.invoice=d.invoice)" & vbCrLf
    s = s & "where m.DivisionID = " & HFApp.DivisionID & " and m.job=" & DbQuote(Str, mJob) & vbCrLf
    s = s & "group by m.arcustomer,m.invoice,m.RetainageInvoice,m.invoice,m.invoicedate,m.description,m.draw,m.notes,m.scopeofwork,m.PostedDate,m.PostedBy,m.Cancelled,m.CancelledBy" & vbCrLf
    s = s & "order by m.Invoice" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    With gInvoices
        .Rows = 1
        r = 0
        While Not rs.EOF
            .AddItem ""
            r = r + 1
            .TextMatrix(r, .ColIndex("ARCustomer")) = "" & rs("ARCustomer")
            .TextMatrix(r, .ColIndex("Invoice")) = "" & rs("Invoice")
            .TextMatrix(r, .ColIndex("InvoiceDate")) = "" & rs("InvoiceDate")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("Draw")) = Val("" & rs("Draw"))
            .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
            .TextMatrix(r, .ColIndex("ScopeofWork")) = "" & rs("ScopeOfWork")
            .TextMatrix(r, .ColIndex("PostedDate")) = format("" & rs("PostedDate"), "medium date")
            .TextMatrix(r, .ColIndex("PostedBy")) = "" & rs("PostedBy")
            .TextMatrix(r, .ColIndex("Cancelled")) = "" & rs("Cancelled")
            .TextMatrix(r, .ColIndex("CancelledBy")) = "" & rs("CancelledBy")
            .TextMatrix(r, .ColIndex("Amount")) = "" & rs("Amount")
            If rs("RetainageInvoice") Then
                .TextMatrix(r, .ColIndex("HoldbackBilled")) = "" & rs("Holdback") * -1
                .TextMatrix(r, .ColIndex("Holdback")) = "0"
            Else
                .TextMatrix(r, .ColIndex("Holdback")) = "" & rs("Holdback")
            End If
            
            .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = IIf(.TextMatrix(r, .ColIndex("PostedDate")) = "", vbWindowText, vbGrayText)
            .Cell(flexcpFontStrikethru, r, 0, r, .Cols - 1) = .TextMatrix(r, .ColIndex("Cancelled")) = "True"
            
            rs.MoveNext
        Wend
        .Editable = flexEDNone
    End With

End Sub







Private Sub gAssemblies_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Row < 0 Then Exit Sub
    
    With gAssemblies
        If .Cell(flexcpText, Row, 3) = msSome Then .Cell(flexcpText, Row, 3) = msNone
        Set .Cell(flexcpPicture, Row, 0) = MultiStateIcon(.Cell(flexcpValue, Row, 2), .Cell(flexcpValue, Row, 3), .Cell(flexcpValue, Row, .ColIndex("IsChangeRequest")), .Cell(flexcpText, Row, .ColIndex("POIconType")))
    End With
    Call ResetChildrenChecked(Row)
    Call ResetParentChecked(Row)

    
End Sub

Private Sub gAssemblies_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
    Dim r As Long
    Dim levels As Long
    Dim Level As Long
    
    With gAssemblies
        If State = flexOutlineExpanded Then
            .Row = Row
            Call LoadAssemblies(False)
        End If
    End With
End Sub

Private Sub gAssemblies_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub gAssemblies_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error Resume Next
    gAssemblies.Row = gAssemblies.MouseRow
End Sub

Private Sub gAssemblies_DblClick()
    Call gAssemblies_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub gAssemblies_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gAssemblies
        If .Row < 0 Then Exit Sub
        Select Case True
        
            Case KeyCode = vbKeyF And Shift = vbCtrlMask
                Call FFind.ShowForm(gAssemblies)
            
            Case KeyCode = vbKeySpace
                'toggle selected value
                i = .Row
                .Cell(flexcpText, i, 3) = IIf(.Cell(flexcpValue, i, 3) = msAll, msNone, msAll)
                Set .Cell(flexcpPicture, i, 0) = MultiStateIcon(.Cell(flexcpValue, i, 2), .Cell(flexcpValue, i, 3), .Cell(flexcpValue, i, .ColIndex("IsChangeRequest")), .Cell(flexcpText, i, .ColIndex("POIconType")))
                Call gAssemblies_AfterEdit(i, 0)
            
            Case KeyCode = vbKeyReturn
                .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                .Cell(flexcpFontBold, .Row, 0) = True
                Call LoadItems
                KeyCode = 0
                
            Case KeyCode = vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
                
            Case KeyCode = vbKeyRight
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

Private Sub gAssemblies_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As Long
    Dim i As Long
    
    Dim s As String
    Dim rs As Recordset
    Dim checked As Boolean
    
    With gAssemblies
        If Button = vbRightButton And Shift <> 0 Then
        
            s = ""
            s = s & "SELECT SUM(BudgetPretax + BudgetJCTax) Budgeted" & vbCrLf
            s = s & "      ,SUM(CASE WHEN POGenBatch=0 THEN 0 ELSE POPretax+POJCTax END) Commited" & vbCrLf
            s = s & "  FROM EstimatedItems" & vbCrLf
            s = s & "WHERE " & Mid(gAssemblies.RowData(gAssemblies.Row), 6)
            Set rs = HFApp.SqlExec(s, dbHomefront)
            If Not rs.EOF Then
                .Cell(flexcpText, .Row, 4) = format("" & rs("budgeted"), "0.00")
                .Cell(flexcpText, .Row, 5) = format("" & rs("commited"), "0.00")
            End If
            
        End If
        
    
        .SetFocus
        r = .MouseRow
        If r < 1 Then Exit Sub
        i = .RowOutlineLevel(r)
        If (X >= (14 + i * 13) * Screen.TwipsPerPixelX) And (X <= (24 + i * 13) * Screen.TwipsPerPixelX) Then
            'they clicked on the checkbox
            .Cell(flexcpText, r, 3) = IIf(.Cell(flexcpValue, r, 3) = msAll, msNone, msAll)
            Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(.Cell(flexcpValue, r, 2), .Cell(flexcpValue, r, 3), .Cell(flexcpValue, r, .ColIndex("IsChangeRequest")), .Cell(flexcpText, r, .ColIndex("POIconType")))
            Call gAssemblies_AfterEdit(r, 0)
        Else
            'show menu depending on node type
            If Button = vbRightButton Then
                Select Case .Cell(flexcpData, .Row, 1)
                
                    Case "RFP"
                        If Parse(.Cell(flexcpText, .Row, 1), 1, Chr(2)) <> "" Then Call PopupMenu(FMain.mnuEstimateRFPs)
                    
                    Case "POVendor", "BudgetVendor"
                        If Parse(.Cell(flexcpText, .Row, 1), 1, Chr(2)) <> "" Then Call PopupMenu(FMain.mnuEstimateItemsVendors)
                        
                    Case "PONumber", "PONumberVendor"
                        If HFApp.Options.ValueByName("HidePOChangeVendor") = "True" Then FMain.mnuEstimateItemsPOSub(mcPO_CHANGE).Visible = False
                        FMain.mnuEstimateItemsPOSub(mcPO_CHANGE).Enabled = Not ReadOnly
                        FMain.mnuEstimateItemsPOSub(mcPO_CANCEL).Enabled = Not ReadOnly
                        
                        
                        If Parse(.Cell(flexcpText, .Row, 1), 1, Chr(2)) <> "" Then Call PopupMenu(FMain.mnuEstimateItemsPO)

                    Case "Customer_no"
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADDCO).Enabled = HFApp.Options(CanAddContractItemsFromPurchasing) And Not ReadOnly
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADDCR).Enabled = mUseChangeRequests And Not ReadOnly
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADD).Enabled = False
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_DELETE).Enabled = False
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_RENAME).Enabled = False
                        
                        checked = Val(.Cell(flexcpText, .Row, .ColIndex("BudgetsLocked"))) <> 0
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_FINALIZE).checked = checked
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_FINALIZE).Enabled = mMode = fceBudget And Not ReadOnly And (Not checked Or HFApp.UserPermission("UnFinalizeBudgets"))
                        
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ATTACHQUOTE).Enabled = False
                        Call PopupMenu(FMain.mnuEstimateItemsAssembly)
                    
                    
                    Case "CORSort", "ChangeOrder"
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADDCO).Enabled = False
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADDCR).Enabled = False
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADD).Enabled = HFApp.Options(CanAddContractItemsFromPurchasing) And Not ReadOnly
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_DELETE).Enabled = False
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_RENAME).Enabled = HFApp.Options(CanAddContractItemsFromPurchasing) And Not ReadOnly
                        
                        checked = Val(.Cell(flexcpText, .Row, .ColIndex("BudgetsLocked"))) <> 0
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_FINALIZE).checked = checked
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_FINALIZE).Enabled = mMode = fceBudget And Val(.Cell(flexcpText, .Row, .ColIndex("IsChangeRequest"))) = 0 And Not ReadOnly And (Not checked Or HFApp.UserPermission("UnFinalizeBudgets"))
                        
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ATTACHQUOTE).Enabled = False
                        Call PopupMenu(FMain.mnuEstimateItemsAssembly)
                        
                        
                    Case "EstAssemblyID"
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADDCO).Enabled = False
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADDCR).Enabled = False
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ADD).Enabled = False
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_DELETE).Enabled = HFApp.Options(CanAddContractItemsFromPurchasing) And Not ReadOnly
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_RENAME).Enabled = True
                        
                        checked = Val(.Cell(flexcpText, .Row, .ColIndex("BudgetsLocked"))) <> 0
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_FINALIZE).checked = checked
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_FINALIZE).Enabled = mMode = fceBudget And Val(.Cell(flexcpText, .Row, .ColIndex("IsChangeRequest"))) = 0 And Not ReadOnly And (Not checked Or HFApp.UserPermission("UnFinalizeBudgets"))
                        
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_ATTACHQUOTE).Enabled = Not ReadOnly
                        FMain.mnuEstimateItemsAssemblySub(mcASMBlY_COPYFROMJOB).Enabled = Not ReadOnly
                        Call PopupMenu(FMain.mnuEstimateItemsAssembly)
                        
                        
                        
                End Select
            End If
        End If
    End With
End Sub

Private Sub gAssemblies_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim s As String
    Dim i As Long
    If InIde And gAssemblies.MouseRow = 0 Then
        s = ""
        s = s & "View Definition:" & vbCrLf
        s = s & "------------------------------------" & vbCrLf
        For i = 1 To Parse(mViews(mViewIndex).KeyFlds)
            s = s & String(4 * (i - 1), " ") & Parse(mViews(mViewIndex).KeyFlds, i) & vbCrLf
        Next
        s = s & vbCrLf & vbCrLf
        s = s & "Last Query:" & vbCrLf
        s = s & "------------------------------------" & vbCrLf
s = ""
        s = s & mViewQuery
        MsgBox s, vbInformation, "View Definition"

    End If
End Sub

Private Sub gAssemblies_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim i As Long
    
    With gAssemblies
        i = .MouseRow
        If i = -1 Then
            .ToolTipText = ""
        Else
            Select Case .MouseCol
                Case 0: .ToolTipText = PrettyName(.Cell(flexcpData, .MouseRow, 0))
                Case 4: .ToolTipText = "Budgeted"
                Case 5: .ToolTipText = "Changes"
                Case 6: .ToolTipText = "Commited"
            End Select
        End If
    End With
End Sub

Private Sub gAssemblies_RowColChange()
On Error Resume Next
    With gAssemblies
        .Col = 0
        Call .ShowCell(.Row, 0)
    End With
End Sub

Private Function SaveContract() As Boolean
    Dim i As Long
    Dim s As String
    
    With gContract
    For i = .Rows - 1 To 1 Step -1
        If .RowHidden(i) Then
            
            s = ""
            s = s & "update estimateassemblies" & vbCrLf
            s = s & "set changeorder=''" & vbCrLf
            s = s & "   ,changeapprovedby=''" & vbCrLf
            s = s & "   ,changeapproveddate=null" & vbCrLf
            s = s & "   ,changerequeststatus = case when isnull(changeorder,'')='' then changerequeststatus else 'Pending' end" & vbCrLf
            s = s & "where estassemblyid in(select distinct estassemblyid from estimateditems where billingitemid=" & .TextMatrix(i, .ColIndex("BillingItemID")) & ")"
            Call HFApp.SqlExec(s)
        
            s = ""
            s = s & "delete from billingitems where billingitemid=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BillingItemID"))) & vbCrLf
            s = s & "update estimateitems set billingitemid=0 where billingitemid=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BillingItemID"))) & vbCrLf
            Call HFApp.SqlExec(s)
            Call .RemoveItem(i)
        End If
    Next
        
    For i = 1 To .Rows - 1
        
        If .RowData(i) = "New" Then
            s = ""
            s = s & "INSERT INTO billingitems(job,Description,Price,HoldBackRate,TaxGroup,RevenueAccount,SortOrder)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, mJob) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Description"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("HoldBackRate"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("RevenueAccount"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, i) & ")"
            Call HFApp.SqlExec(s)
            .TextMatrix(i, .ColIndex("BillingItemID")) = HFApp.SqlIdentity("BillingItems")
            .RowData(i) = ""
        Else
            s = ""
            s = s & "update billingitems" & vbCrLf
            s = s & "   set description=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Description"))) & vbCrLf
            s = s & "      ,price=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
            s = s & "      ,HoldBackRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("HoldBackRate"))) & vbCrLf
            s = s & "      ,TaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "      ,RevenueAccount=" & DbQuote(Str, .TextMatrix(i, .ColIndex("RevenueAccount"))) & vbCrLf
            s = s & "      ,sortorder=" & DbQuote(Num, i) & vbCrLf
            s = s & " where billingitemid=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BillingItemID"))) & vbCrLf
            Call HFApp.SqlExec(s)
            .RowData(i) = ""
        End If
    
    Next
    End With
    SaveContract = True
    
End Function


Private Sub lblMoveLine_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Top As Single
    Dim NewLine As Long
    Dim g As VSFlexGrid
    
    Select Case Index
        Case 0:  Set g = gContract
    End Select
    If g.Editable = flexEDNone Then
        lblMoveLine(Index).Visible = False
        Exit Sub
    End If
    
    
    If Button <> 0 Then
        If g.Row <> g.RowSel Then g.Row = g.Row
        Top = lblMoveLine(Index).Top + Y
        NewLine = Top / g.RowHeight(0)
        
        If NewLine < 1 Then NewLine = 1
        
        
        If NewLine > g.Rows - 1 Then NewLine = g.Rows - 1
        Top = NewLine * g.RowHeight(0)
        lblMoveLine(Index).Move 0, Top, g.Width, 30
    End If



End Sub

Private Sub lblMoveLine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Top As Single
    Dim NewLine As Long
    Dim g As VSFlexGrid
    Dim r As Long
    
    
    lblMoveLine(Index).Visible = False
    Select Case Index
        Case 0:  Set g = gContract
    End Select
    If g.Editable = flexEDNone Then Exit Sub
    
    
    Top = lblMoveLine(Index).Top + 30
    NewLine = GetRowFromPoint(g, Top)
    If g.Row < NewLine Then NewLine = NewLine - 1
    
    If g.Row = NewLine Then
        lblMoveLine(0).Move g.CellLeft, g.CellTop, 60, 60
        lblMoveLine(Index).Visible = True
        Exit Sub
    End If
    
    g.RowPosition(g.Row) = NewLine
    g.Row = NewLine
    Dirty = True
End Sub

Private Function GetRowFromPoint(g As VSFlexGrid, Y As Single) As Long
On Error GoTo done
    Dim r As Long
    With g
        r = 0
        While .RowPos(r) < Y
            r = r + 1
        Wend
    End With

done:
r = r - 1
GetRowFromPoint = r
End Function

Private Sub gContract_RowColChange()
On Error Resume Next
    With gContract
        If .Row = .Rows - 1 Then
            lblMoveLine(0).Visible = False
        Else
            lblMoveLine(0).Move .CellLeft, .CellTop, 60, 60
            lblMoveLine(0).Visible = True
        End If
    End With
End Sub

Public Sub mnuDeleteSub_Click(Index As Integer)
    Dim i As Long
    
    With gContract
        If vbYes <> MsgBox("Are you sure you want to delete " & IIf(.Row = .RowSel, "this item?", "these items?"), vbYesNo + vbQuestion, "Confirm Delete") Then Exit Sub
        For i = .Row To .RowSel
            If i > 0 And i < .Rows - 1 Then
               If Val(.TextMatrix(i, .ColIndex("Billed"))) = 0 Then
                    .RowHidden(i) = True
                    Dirty = True
               Else
                    MsgBox "Unable to delete contract items which have already been invoiced previously", vbCritical, "Warning"
               End If
            End If
        Next
    End With
End Sub




Public Sub GroupGrid()
On Error GoTo eh
    
    Dim i As Long
    Dim OldCol As Long
    Dim Row As Long
    Dim Col As Long
    Dim currow As Long, RowPosVal As Long
    
Const ITEMCOLOR = vbHighlight
Const GROUPCOLOR = vbHighlight
    
    
    
    
    If mSplitting Then Exit Sub
    
    With gItems
        'first ungroup everything
        'expand all groups. the clear method doesnt reshow rows hidden because they are collapsed.
        For i = .Rows - 1 To 1 Step -1
            If .IsSubtotal(i) Then .GetNode(i).Expanded = True
        Next
        .SubTotal flexSTClear
        
        
        
        
        'now reset grouping
        OldCol = .Col
    
        RowPosVal = .RowPos(.Row)
        '.Redraw = flexRDNone
        
        Row = .Row
        Col = .Col
        
        .ColPosition(.ColIndex("WarningMessages")) = 0
        .ColWidth(.ColIndex("WarningMessages")) = 300
        .ColPosition(.ColIndex("Selected")) = 1
     
        'ensure itemdesc column if grouped is right most. formatting gets weird if it is left of another group.
        If .ColData(.ColIndex("ItemDesc")) = "GROUPED" Then
            .ColPosition(.ColIndex("ItemDesc")) = .Cols - 1
        End If
     
        GroupedColumns = 0
        For i = 0 To .Cols - 1
            If .ColData(i) = "GROUPED" Then
                .ColPosition(i) = GroupedColumns + 2
                GroupedColumns = GroupedColumns + 1
            End If
        Next
        
        If GroupedColumns <> 0 Then
            
            'sort by grouped columns
            mFunnyFlag = True
            .Col = 2
            .ColSel = 2 + GroupedColumns
            .Sort = flexSortGenericAscending
            mFunnyFlag = False
            
            
            .OutlineCol = 2
            .SubtotalPosition = flexSTAbove
            
            For i = 2 To GroupedColumns + 1
            If .ColData(i) = "GROUPED" Then
                If .ColKey(i) = "ItemDesc" Then
                
                
                    .SubTotal flexSTSum, i, .ColIndex("BudgetPretax"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetJCTax"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetNJCTax"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POPretax"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POJCTax"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("PONJCTax"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POTax"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetTax"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POTotal"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetTotal"), "$(#,###.00)", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("TakeoffQty"), "0.00000000000", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetQty"), "0.00000000000", , ITEMCOLOR, , "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POQty"), "0.00000000000", , ITEMCOLOR, , "%s", 0
                    'special handling
                    'temp row count column. subtotal adds header even if only one row. we only want where duplicates exist
                    .SubTotal flexSTCount, i, .ColIndex("JobDesc"), , , ITEMCOLOR, , "%s", 0

                Else
                    .SubTotal flexSTSum, i, .ColIndex("BudgetPretax"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetJCTax"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetNJCTax"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POPretax"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POJCTax"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("PONJCTax"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POTax"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetTax"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("POTotal"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
                    .SubTotal flexSTSum, i, .ColIndex("BudgetTotal"), "$(#,###.00)", , GROUPCOLOR, True, "%s", 0
               End If
            End If
            Next
            
            'now we need to subtract any "deleted rows" from the autocalculated subtotals.
            For i = 1 To .Rows - 1
                If Not .IsSubtotal(i) Then
                    If ((IsIn(mMode, fceQuote, fceBudget) And .TextMatrix(i, .ColIndex("BudgetDeleted")) = "True") Or ((Not IsIn(mMode, fceQuote, fceBudget)) And .TextMatrix(i, .ColIndex("PODeleted")) = "True")) Then
                        'remove row if its been deleted
                        
                        .TextMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("BudgetPretax")) = .ValueMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("BudgetPretax")) - .ValueMatrix(i, .ColIndex("BudgetPretax"))
                        .TextMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("BudgetJCTax")) = .ValueMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("BudgetJCTax")) - .ValueMatrix(i, .ColIndex("BudgetJCTax"))
                        .TextMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("BudgetNJCTax")) = .ValueMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("BudgetNJCTax")) - .ValueMatrix(i, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("POPretax")) = .ValueMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("POPretax")) - .ValueMatrix(i, .ColIndex("POPretax"))
                        .TextMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("POJCTax")) = .ValueMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("POJCTax")) - .ValueMatrix(i, .ColIndex("POJCTax"))
                        .TextMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("PONJCTax")) = .ValueMatrix(.GetNodeRow(i, flexNTParent), .ColIndex("PONJCTax")) - .ValueMatrix(i, .ColIndex("PONJCTax"))
                    End If
                End If
            Next
            

            'special handling
            'add extra titles and formatting to ItemDesc group header columns
            For i = .Rows - 1 To 1 Step -1
            If .IsSubtotal(i) And .ColKey(.RowOutlineLevel(i)) = "ItemDesc" Then
                If .ValueMatrix(i, .ColIndex("JobDesc")) < 2 Then
                    '.RemoveItem i  <-- dont remove. that makes the previous group inherit this groups items
                    .RowHidden(i) = True
                    .TextMatrix(i, .ColIndex("EstItemID")) = "HIDDEN"
                Else
                    .GetNode(i).Expanded = False
                    .TextMatrix(i, .ColIndex("ItemDesc")) = .TextMatrix(i + 1, .ColIndex("ItemDesc")) & " (" & .TextMatrix(i, .ColIndex("JobDesc")) & ")"
                    
                    If .ValueMatrix(i, .ColIndex("BudgetQty")) <> 0 Then
                        .TextMatrix(i, .ColIndex("BudgetRate")) = .ValueMatrix(i, .ColIndex("BudgetPretax")) / .ValueMatrix(i, .ColIndex("BudgetQty"))
                    End If
                    If .ValueMatrix(i, .ColIndex("POQty")) Then
                        .TextMatrix(i, .ColIndex("PORate")) = .ValueMatrix(i, .ColIndex("POPretax")) / .ValueMatrix(i, .ColIndex("POQty"))
                    End If
                    
                    'reformat these... default is 2 decimal places whether they exist or not.. yucky
                    .TextMatrix(i, .ColIndex("TakeoffQty")) = .ValueMatrix(i, .ColIndex("TakeoffQty"))
                    .TextMatrix(i, .ColIndex("BudgetQty")) = .ValueMatrix(i, .ColIndex("BudgetQty"))
                    .TextMatrix(i, .ColIndex("POQty")) = .ValueMatrix(i, .ColIndex("POQty"))
                    
                End If
                .TextMatrix(i, .ColIndex("JobDesc")) = ""
            End If
            Next
            
            
        End If
    


        .Col = OldCol
        '.Redraw = True
        
    End With
    
    ShowCancelledItems
    Call ColorizeItems(0)
    
Exit Sub
eh: Call errHandler(SRCFILE & "GroupGrid")
End Sub

Private Sub ClearGroups()
    Dim i As Long
    With gItems
    For i = 0 To .Cols - 1
        .ColData(i) = ""
    Next
    End With
End Sub


Private Function ViewIndex(Name As String) As Long
    Dim i As Long
    For i = 0 To UBound(mViews)
        If mViews(i).Name = Name Then
            ViewIndex = i
            Exit Function
        End If
    Next
End Function

Private Sub gInvoices_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With gInvoices
    If Button = vbRightButton Then
        If .MouseRow = 0 Then
            Call FMain.ShowColumnMenu(gInvoices)
        Else
            If .Row = -1 Then
                FMain.mnuInvoicesSub(mcINV_PREVIEW).Enabled = False
                FMain.mnuInvoicesSub(mcINV_PRINT).Enabled = False
                FMain.mnuInvoicesSub(mcINV_NEW).Enabled = True
                FMain.mnuInvoicesSub(mcINV_EDIT).Enabled = False
                FMain.mnuInvoicesSub(mcINV_DELETE).Enabled = False
                FMain.mnuInvoicesSub(mcINV_VOID).Enabled = False
                FMain.mnuInvoicesSub(mcINV_POST).Enabled = False
                PopupMenu FMain.mnuInvoices
            Else
                FMain.mnuInvoicesSub(mcINV_PREVIEW).Enabled = True
                FMain.mnuInvoicesSub(mcINV_PRINT).Enabled = True
                FMain.mnuInvoicesSub(mcINV_NEW).Enabled = True
                FMain.mnuInvoicesSub(mcINV_EDIT).Enabled = True
                FMain.mnuInvoicesSub(mcINV_DELETE).Enabled = .TextMatrix(.Row, .ColIndex("PostedDate")) = "" And .TextMatrix(.Row, .ColIndex("Cancelled")) = "False"
                FMain.mnuInvoicesSub(mcINV_VOID).Enabled = .TextMatrix(.Row, .ColIndex("Cancelled")) = "False"
                FMain.mnuInvoicesSub(mcINV_POST).Enabled = IsIn(HFApp.Options(AccountingSystem), asSimply, asQuickBooks)
                PopupMenu FMain.mnuInvoices
            End If
        End If
    End If
    End With
End Sub

Public Sub mnuInvoicesSub_Click(Index As Integer)
On Error GoTo eh
    'Dim f As New FRptViewer
    Dim Invoice As Long
    Dim RptFile As String
    Dim PrinterName As String
    Dim i As Long
    
    With gInvoices
        RptFile = HFApp.SystemFolder & "System\Reports\Estimating\ARInvoice.rpt"
        If .Row > 0 Then Invoice = .TextMatrix(.Row, .ColIndex("Invoice"))
        Select Case Index
            'Case mcINV_PREVIEW:    Call f.ShowReport(RptFile, True, False, "Invoice", Invoice)
            'Case mcINV_PRINT:      Call f.ShowReport(RptFile, False, False, "Invoice", Invoice)
            Case mcINV_PRINT, mcINV_PREVIEW
                Dim c As New ZybUtil.Crystal
                Call c.LoadODBCReport(RptFile, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
                On Error Resume Next
                Call c.ParameterValue("DivisionID", HFApp.DivisionID)
                Call c.ParameterValue("Invoice", Invoice)
                On Error GoTo eh
                If Index = mcINV_PREVIEW Then
                    Call c.PrintPreview("Print Preview")
                Else
                    If Not VBPrintDlg(i, eprAll, True, , , True, , , , , , , Me.hwnd, , PrinterName) Then Exit Sub
                    Call c.PrintReport(PrinterName)
                End If
            
            Case mcINV_NEW:        If FCreateInvoice.Edit(mJob) Then Call LoadContract
            Case mcINV_EDIT:       If FCreateInvoice.Edit(mJob, Invoice) Then Call LoadContract
            Case mcINV_DELETE:     Call HFApp.SqlExec("exec Purch_DeleteARInvoice " & DbQuote(Str, Invoice), dbHomefront):    Call LoadContract
            Case mcINV_VOID:       Call HFApp.SqlExec("exec Purch_VoidARInvoice " & DbQuote(Str, HFApp.LoginID) & "," & DbQuote(Str, Invoice), dbHomefront):      Call LoadContract
            Case mcINV_POST:       Call PostInvoice(Invoice)
        End Select
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "mnuInvoicesSub_Click")
End Sub

Private Sub PostInvoice(Invoice As Long)
On Error GoTo eh

    Select Case HFApp.Options(AccountingSystem)
        Case asSimply:     Call PostInvoicesToSimply(DbQuote(Num, Invoice))
        Case asQuickBooks: Call PostInvoicesToQB(DbQuote(Num, Invoice), False)
    End Select

Exit Sub
eh: Call errHandler(SRCFILE & "PostInvoice")
End Sub


Public Sub PostInvoicesToQB(InvoiceCSV As String, Repost As Boolean)
On Error GoTo eh

    Dim s As String
    Dim Invoice As String
    Dim rs As Recordset


    s = ""
    s = s & "select m.ARcustomer Customer" & vbCrLf
    s = s & "      ,m.invoice InvoiceNumber" & vbCrLf
    s = s & "      ,m.invoicedate" & vbCrLf
    s = s & "      ,m.description InvoiceDesc" & vbCrLf
    s = s & "      ,d.description ItemDescription" & vbCrLf
    s = s & "      ,o.OptionValue BillingItemListID" & vbCrLf
    s = s & "      ,1 qty" & vbCrLf
    s = s & "      ,'' uom" & vbCrLf
    s = s & "      ,round(d.amtThisPeriod-isnull(d.HoldbackAmount,0),2) Price" & vbCrLf
    s = s & "      ,round(d.amtThisPeriod-isnull(d.HoldbackAmount,0),2) Pretax" & vbCrLf
    s = s & "      ,d.taxgroup" & vbCrLf
    s = s & "      ,d.tax" & vbCrLf
    s = s & "      ,d.revenueaccount" & vbCrLf
    s = s & "      ,0 discount" & vbCrLf
    s = s & "      ,j.externaljobid JobExternalID" & vbCrLf
    s = s & "      ,g.ExternalID TaxGroupExternalID"
    s = s & "      ,c.BillAddr1 Addr1, c.BillAddr2 Addr2 , c.BillCity City, c.BillProvince State, c.BillPostalCode Postal"
    s = s & "  from arinvoices m" & vbCrLf
    s = s & "       left outer join tbljobs j on(m.job=j.job_no and j.DivisionID = m.DivisionID )" & vbCrLf
    s = s & "       left outer join arinvoiceitems d on(m.invoice=d.invoice)" & vbCrLf
    s = s & "       left outer join taxgroups g on (d.DivisionID = g.DivisionID and d.TaxGroup = g.TaxGroup)" & vbCrLf
    s = s & "       left outer join ARcustomers c on(m.DivisionID = c.DivisionID and m.ARCustomer=c.ARCustomer)" & vbCrLf
    s = s & "       left outer join appoptions o on o.DivisionID = " & HFApp.DivisionID & " and uid = '' and optionname = 'QBBillingItemListID'" & vbCrLf
    s = s & " where m.DivisionID = " & HFApp.DivisionID & " and isnull(m.invoice,'')<>''" & vbCrLf
    s = s & "   and isnull(d.amtthisperiod,0)+isnull(d.tax,0)<>0" & vbCrLf
    s = s & "   and m.invoice IN(" & InvoiceCSV & ")" & vbCrLf
    s = s & " and m.PostedDate is null "
    s = s & "ORDER BY m.invoice,d.invoiceitem"
    Set rs = HFApp.SqlExec(s, dbHomefront)

    s = ""
    s = s & HFApp.XmlQBStart()
    While Not rs.EOF
        If Invoice <> "" & rs("InvoiceNumber") Then
            If Invoice <> "" Then
                s = s & "</InvoiceAdd>" & vbCrLf
                s = s & "</InvoiceAddRq>" & vbCrLf
            End If
            Invoice = "" & rs("InvoiceNumber")
            s = s & "<InvoiceAddRq>" & vbCrLf
            s = s & "<InvoiceAdd>" & vbCrLf
            s = s & "<CustomerRef>" & vbCrLf
            
            If "" & HFApp.Options.ValueByName("UseQBARByJob") = "True" Then
                s = s & HFApp.XmlQBAdd(c, 41, "ListID", "" & rs("JobExternalID"))
            Else
                s = s & HFApp.XmlQBAdd(c, 41, "ListID", "" & rs("Customer"))
            End If

            s = s & "</CustomerRef>"
            s = s & HFApp.XmlQBAdd(d, 0, "TxnDate", "" & rs("invoicedate"))
            s = s & HFApp.XmlQBAdd(c, 50, "RefNumber", Invoice)
'            s = s & "<BillAddress>" & vbCrLf
'                s = s & AddQBXml(c, 41, "Addr1", "" & rs("Addr1"))
'                s = s & AddQBXml(c, 41, "Addr2", "" & rs("Addr2"))
'                s = s & AddQBXml(c, 31, "City", "" & rs("city"))
'                s = s & AddQBXml(c, 21, "State", "" & rs("state"))
'                s = s & AddQBXml(c, 13, "PostalCode", "" & rs("postal"))
'            s = s & "</BillAddress>" & vbCrLf
            s = s & HFApp.XmlQBAdd(c, 4000, "Memo", "" & rs("invoicedesc"))
            s = s & HFApp.XmlQBAdd(m, 1, "IsToBePrinted", "0")
        End If
        s = s & "<InvoiceLineAdd>" & vbCrLf
        s = s & "<ItemRef>"
        s = s & HFApp.XmlQBAdd(c, 41, "ListID", "" & rs("BillingItemListID"))
        s = s & "</ItemRef>"
        s = s & HFApp.XmlQBAdd(c, 4000, "Desc", "" & rs("itemdescription"))
        s = s & HFApp.XmlQBAdd(n, 7.5, "Quantity", "" & rs("qty"))

        s = s & HFApp.XmlQBAdd(n, 8.5, "Rate", "" & rs("price"))
        
'        If "" & rs("category") <> "" Then
'            s = s & "<ClassRef>"
'            s = s & AddQBXml(c, 41, "ListID", "" & rs("category"))
'            s = s & "</ClassRef>"
'        End If
        If "" & rs("TaxGroupExternalID") <> "" Then
            s = s & "<SalesTaxCodeRef>" & vbCrLf
            s = s & HFApp.XmlQBAdd(c, 41, "ListID", "" & rs("TaxGroupExternalID"))
        s = s & "</SalesTaxCodeRef>" & vbCrLf
        End If
        s = s & "</InvoiceLineAdd>" & vbCrLf
        rs.MoveNext
    Wend
    s = s & "</InvoiceAdd>" & vbCrLf
    s = s & "</InvoiceAddRq>" & vbCrLf
    s = s & HFApp.XmlQBEnd()
    Call HFApp.XmlQBSubmit(s)
    
    s = ""
    s = s & "UPDATE arinvoices" & vbCrLf
    s = s & "SET PostedDate=getdate()" & vbCrLf
    s = s & "   ,PostedBy=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Invoice IN(" & InvoiceCSV & ")" & vbCrLf
    Call HFApp.SqlExec(s, dbHomefront)


Exit Sub
eh: Call errHandler(SRCFILE & "PostInvoicesToQB", s)

End Sub

Private Sub PostInvoicesToSimply(InvoicesCSV As String)
On Error GoTo eh

    Dim X As SimplyWrapper.ARInvoice
    Dim s As String
    Dim Invoice As String
    Dim rs As Recordset

    s = ""
    s = s & "select m.ARcustomer customer" & vbCrLf
    s = s & "      ,m.invoice" & vbCrLf
    s = s & "      ,m.invoicedate" & vbCrLf
    s = s & "      ,m.description" & vbCrLf
    s = s & "      ,d.description ItemDesc" & vbCrLf
    s = s & "      ,1 qty" & vbCrLf
    s = s & "      ,'' uom" & vbCrLf
    s = s & "      ,isnull(d.amtThisPeriod,0)-isnull(d.HoldbackAmount,0) rate" & vbCrLf
    s = s & "      ,d.taxgroup" & vbCrLf
    s = s & "      ,d.tax" & vbCrLf
    s = s & "      ,d.revenueaccount" & vbCrLf
    s = s & "      ,0 discount" & vbCrLf
    s = s & "      ,j.externaljobid project" & vbCrLf
    s = s & "  from arinvoices m" & vbCrLf
    s = s & "       left outer join tbljobs j on(m.job=j.job_no and j.DivisionID = m.DivisionID)" & vbCrLf
    s = s & "       left outer join arinvoiceitems d on(m.invoice=d.invoice)" & vbCrLf
    s = s & " where isnull(m.invoice,'')<>''" & vbCrLf
    s = s & "   and isnull(d.amtThisPeriod,0)-isnull(d.HoldbackAmount,0)<>0" & vbCrLf
    s = s & "   and m.invoice IN(" & InvoicesCSV & ")" & vbCrLf
    s = s & " and m.PostedDate is null "
    s = s & "ORDER BY m.invoice,d.invoiceitem"
    Set rs = HFApp.SqlExec(s, dbHomefront)

    Invoice = Chr(1)
    If rs.EOF Then Exit Sub

    s = "Set X = New SimplyWrapper.ARInvoice"
    Set X = New SimplyWrapper.ARInvoice

    s = "Call X.OpenDB"
    If Not X.OpenDB(HFApp.Options(SimplyDataFile), HFApp.Options.ValueByName("SimplyUID"), HFApp.Options.ValueByName("SimplyPWD")) Then Exit Sub
    
    While Not rs.EOF
        If Invoice <> "" & rs("Invoice") Then
            If Invoice <> Chr(1) Then
                s = "Call X.SaveInvoice"
                Call X.SaveInvoice
            End If
            Invoice = "" & rs("Invoice")
            s = "Call X.CreateInvoice"
            Call X.CreateInvoice("" & rs("Customer"), Invoice, rs("invoicedate"), "" & rs("description"))
        End If
        s = "Call X.AddLine"
        Call X.AddLine2("" & rs("itemdesc"), Val("" & rs("qty")), "" & rs("uom"), Val("" & rs("rate")), "" & rs("taxgroup"), "" & rs("revenueaccount"), Val("" & rs("discount")), "" & rs("project"))
        rs.MoveNext
    Wend
    If Invoice <> "" Then
        s = "Call X.SaveInvoice"
        Call X.SaveInvoice
    End If
    s = "Call X.CloseDB"
    Call X.CloseDB


    s = ""
    s = s & "UPDATE arinvoices" & vbCrLf
    s = s & "SET PostedDate=getdate()" & vbCrLf
    s = s & "   ,PostedBy=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "WHERE Invoice IN(" & InvoicesCSV & ")" & vbCrLf
    Call HFApp.SqlExec(s, dbHomefront)


Exit Sub
eh: Call errHandler(SRCFILE & "PostInvoicesToSimply", s)
End Sub


Private Sub LoadCustomDescriptions()
    lblCommunity.Caption = FMain.CD_Community
    Label20.Caption = FMain.CD_Color & "/Location"
End Sub


Private Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property
Private Property Let ReadOnly(RHS As Boolean)
    Dim b As Boolean
    Dim i As Long
    
    mReadOnly = RHS
    
'    frmJob.BackColor = IIf(mReadOnly, &H8080FF, vbButtonFace)
    
    b = HFApp.Options(SalesSystem) = SalesSystems.asNone Or mMode = fceQuote
    txtJob.Enabled = mJob = "" And b
    
    txtDescription.Enabled = b
    cboCommunity.Enabled = b
    cboPhase.Enabled = b
    cboModel.Enabled = b
    txtLot.Enabled = b
    txtBlock.Enabled = b
    txtLotPlan.Enabled = b
    
    For i = 3 To Toolbar.Buttons.Count
        Toolbar.Buttons(i).Enabled = Not mReadOnly
    Next
    
    If Not mReadOnly Then
        Toolbar.Buttons("SendPOs").Enabled = mJob <> ""
        Toolbar.Buttons("SendRFQs").Enabled = mJob <> ""
        Toolbar.Buttons("NewRFQ").Enabled = mJob <> ""
        Toolbar.Buttons("Delete").Enabled = mJob <> "" And mMode = fceQuote
    End If
        
    Toolbar.Buttons("Open").Enabled = True
    Toolbar.Buttons("View").Enabled = True
    Toolbar.Buttons("ViewBudgets").Enabled = True
    Toolbar.Buttons("ViewPOs").Enabled = True
                    
End Property

Private Sub WriteException(ProcedureName As String, Message As String, StackTrace As String)
On Error GoTo eh
    Dim s As String

    s = ""
    s = s & "insert into exceptions(applicationname,methodname,exceptiondate,message,stacktrace,innerexception,source,targetsite)" & vbCrLf
    s = s & "values('PrecisionBuilder'" & vbCrLf
    s = s & "      ," & DbQuote(Str, ProcedureName) & vbCrLf
    s = s & "      ,getdate()" & vbCrLf
    s = s & "      ," & DbQuote(Str, Message) & vbCrLf
    s = s & "      ," & DbQuote(Str, StackTrace) & vbCrLf
    s = s & "      ,''" & vbCrLf
    s = s & "      ,''" & vbCrLf
    s = s & "      ,'')" & vbCrLf
    Call HFApp.SqlExec(s, dbHomefront)

Exit Sub
eh: Call errHandler(SRCFILE & "WriteException")
End Sub


Private Sub cboModel_Click()
    Dirty = True
End Sub

Public Sub mnuSnapShotsSub_Click(Index As Integer)
    Dim s As String
    Dim r As Long
    
    If FMain.mnuSnapShotsSub(Index).checked Then
        If Not HFApp.UserPermission("SnapShotRemove") Then
            MsgBox "You are not authorized to perform this function", vbInformation + vbOKOnly, "Access Denied"
        Else
            
            If MsgBox("Are you sure you want to clear this snap shot? All snap shot data will be removed?", vbYesNo + vbExclamation, App.ProductName) = vbNo Then Exit Sub
        
            ' remove snapshot
            Call SaveData(False)
            
            s = ""
            s = s & "update e set" & vbCrLf
            s = s & " budgetrate" & Index & "=null" & vbCrLf
            s = s & ",budgetqty" & Index & "=null" & vbCrLf
            s = s & "from tbljobs j" & vbCrLf
            s = s & "join estimateitems e on j.divisionid=e.divisionid and j.job_no=e.job" & vbCrLf
            s = s & "where j.DivisionID=" & DbQuote(Str, HFApp.DivisionID) & vbCrLf
            s = s & "and j.Job_No=" & DbQuote(Str, mJob) & vbCrLf
            s = s & vbCrLf
            s = s & "update j set" & vbCrLf
            s = s & " snapshot" & Index & "date=null" & vbCrLf
            s = s & ",snapshot" & Index & "userid=null" & vbCrLf
            s = s & "from tbljobs j" & vbCrLf
            s = s & "where j.DivisionID=" & DbQuote(Str, HFApp.DivisionID) & vbCrLf
            s = s & "and j.Job_No=" & DbQuote(Str, mJob) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront, r)
                            
            With gItems
            For r = 1 To .Rows - 1
                If Not .IsSubtotal(r) Then
                    .TextMatrix(r, .ColIndex("Budget" & Index & "Qty")) = ""
                    .TextMatrix(r, .ColIndex("Budget" & Index & "Rate")) = ""
                End If
            Next
            End With
        End If
    Else
        If Not HFApp.UserPermission("SnapShotAdd") Then
            MsgBox "You are not authorized to perform this function", vbInformation + vbOKOnly, "Access Denied"
        Else
                
            ' add snapshot
            Call SaveData(False)
            
            s = ""
            s = s & "update e set" & vbCrLf
            s = s & " budgetrate" & Index & "=e.budgetrate" & vbCrLf
            s = s & ",budgetqty" & Index & "=e.budgetqty" & vbCrLf
            s = s & "from tbljobs j" & vbCrLf
            s = s & "join estimateitems e on j.divisionid=e.divisionid and j.job_no=e.job" & vbCrLf
            s = s & "where j.snapshot" & Index & "date is null" & vbCrLf
            s = s & "and j.DivisionID=" & DbQuote(Str, HFApp.DivisionID) & vbCrLf
            s = s & "and j.Job_No=" & DbQuote(Str, mJob) & vbCrLf
            s = s & vbCrLf
            s = s & "update j set" & vbCrLf
            s = s & " snapshot" & Index & "date=getdate()" & vbCrLf
            s = s & ",snapshot" & Index & "userid=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "from tbljobs j" & vbCrLf
            s = s & "where j.snapshot" & Index & "date is null" & vbCrLf
            s = s & "and j.DivisionID=" & DbQuote(Str, HFApp.DivisionID) & vbCrLf
            s = s & "and j.Job_No=" & DbQuote(Str, mJob) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront, r)
                            
            With gItems
            For r = 1 To .Rows - 1
                If Not .IsSubtotal(r) Then
                    .TextMatrix(r, .ColIndex("Budget" & Index & "Qty")) = .TextMatrix(r, .ColIndex("BudgetQty"))
                    .TextMatrix(r, .ColIndex("Budget" & Index & "Rate")) = .TextMatrix(r, .ColIndex("BudgetRate"))
                End If
            Next
            End With
            
        End If
    End If
    
End Sub


Private Sub ShowSnapshotsMenu(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    s = "select snapshot1date,snapshot1userid,snapshot2date,snapshot2userid,snapshot3date,snapshot3userid" & vbCrLf
    s = s & "from tbljobs" & vbCrLf
    s = s & "where DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and Job_No=" & DbQuote(Str, mJob) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    If rs.EOF Then Exit Sub
    
    FMain.mnuSnapShotsSub(1).Caption = "Initial Budget"
    FMain.mnuSnapShotsSub(2).Caption = "Confirmed Budget"
    FMain.mnuSnapShotsSub(3).Caption = "Modified Budget"
    For i = 1 To 3
        FMain.mnuSnapShotsSub(i).checked = "" & rs("snapshot" & i & "date") <> ""
        If FMain.mnuSnapShotsSub(i).checked Then
            FMain.mnuSnapShotsSub(i).Caption = FMain.mnuSnapShotsSub(i).Caption & " (" & rs("snapshot" & i & "userid") & "  " & format("" & rs("snapshot" & i & "date"), "medium date") & ")"
        End If
    Next
    
    PopupMenu FMain.mnuSnapShots, , Button.Left, Button.Top + Button.Height
    
    Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonDropDown")
End Sub

Private Property Get MultiStateIcon(Completed As MultiStateEnum, Selected As MultiStateEnum, IsChangeRequest As Boolean, POIconType As String) As IPictureDisp
    If IsChangeRequest Then
        Set MultiStateIcon = FMain.SmallIcons.ListImages("quote").Picture
    ElseIf mViews(mViewIndex).Name = "RFQ's" Then
        Set MultiStateIcon = Nothing
    Else
        If POIconType <> "" Then
            Set MultiStateIcon = CombineImgs(FMain.MultiStateIcons.ListImages("K" & Completed & Selected).Picture, _
                                             FMain.MultiStateIcons.ListImages(POIconType).Picture)
        Else
            Set MultiStateIcon = FMain.MultiStateIcons.ListImages("K" & Completed & Selected).Picture
        End If
    End If
End Property
Private Function CombineImgs(a As IPictureDisp, b As IPictureDisp) As IPictureDisp
    Picture1.Cls
    With a
        .Render Picture1.hDC, 0&, 0&, ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), 0&, .Height, .Width, -.Height, ByVal 0&
    End With
    With b
        .Render Picture1.hDC, ScaleX(.Width, vbHimetric, vbPixels), 0&, ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), 0&, .Height, .Width, -.Height, ByVal 0&
    End With
    Set CombineImgs = Picture1.Image
End Function
