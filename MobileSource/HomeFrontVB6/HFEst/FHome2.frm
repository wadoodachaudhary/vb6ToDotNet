VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FHome2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Home"
   ClientHeight    =   11745
   ClientLeft      =   1230
   ClientTop       =   1830
   ClientWidth     =   19065
   Icon            =   "FHome2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11745
   ScaleWidth      =   19065
   Begin VB.PictureBox WorkflowBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F0E5&
      ForeColor       =   &H80000008&
      Height          =   10845
      Index           =   0
      Left            =   120
      ScaleHeight     =   10815
      ScaleWidth      =   15255
      TabIndex        =   3
      Top             =   1410
      Width           =   15285
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   9990
         Index           =   0
         Left            =   0
         ScaleHeight     =   9990
         ScaleWidth      =   13365
         TabIndex        =   8
         Top             =   0
         Width           =   13365
         Begin VB.Image Task 
            Height          =   1035
            Index           =   14
            Left            =   8310
            MouseIcon       =   "FHome2.frx":000C
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":0316
            Tag             =   "Salespricing Worksheets"
            Top             =   240
            Width           =   1050
         End
         Begin VB.Image Task 
            Height          =   990
            Index           =   12
            Left            =   2790
            MouseIcon       =   "FHome2.frx":3C7C
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":3F86
            Tag             =   "Model & Option Library"
            Top             =   2190
            Width           =   1155
         End
         Begin VB.Image Task 
            Height          =   720
            Index           =   17
            Left            =   5580
            MouseIcon       =   "FHome2.frx":7B98
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":7EA2
            Tag             =   "Communities"
            Top             =   2310
            Width           =   945
         End
         Begin VB.Image Task 
            Height          =   930
            Index           =   9
            Left            =   8490
            MouseIcon       =   "FHome2.frx":A2E4
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":A5EE
            Tag             =   "Vendor Pricelists"
            Top             =   4590
            Width           =   690
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   8
            Left            =   5670
            MouseIcon       =   "FHome2.frx":C818
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":CB22
            Tag             =   "Vendor"
            Top             =   6810
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   960
            Index           =   18
            Left            =   5640
            MouseIcon       =   "FHome2.frx":E84C
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":EB56
            Tag             =   "Default Vendors"
            Top             =   4620
            Width           =   765
         End
         Begin VB.Image Task 
            Height          =   765
            Index           =   11
            Left            =   390
            MouseIcon       =   "FHome2.frx":11298
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":115A2
            Tag             =   "Item DB"
            Top             =   4530
            Width           =   630
         End
         Begin VB.Image Image1 
            Height          =   8115
            Index           =   4
            Left            =   2160
            Picture         =   "FHome2.frx":12F64
            Stretch         =   -1  'True
            Top             =   2070
            Width           =   9630
         End
      End
   End
   Begin VB.PictureBox WorkflowBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F0E5&
      ForeColor       =   &H80000008&
      Height          =   8115
      Index           =   1
      Left            =   120
      ScaleHeight     =   8085
      ScaleWidth      =   11025
      TabIndex        =   4
      Top             =   930
      Width           =   11055
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   7605
         Index           =   1
         Left            =   0
         ScaleHeight     =   7605
         ScaleWidth      =   10785
         TabIndex        =   9
         Top             =   0
         Width           =   10785
         Begin VB.Image Task 
            Height          =   810
            Index           =   28
            Left            =   6720
            MouseIcon       =   "FHome2.frx":111A0E
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":111D18
            Tag             =   "Contract"
            Top             =   4140
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   795
            Index           =   27
            Left            =   3960
            MouseIcon       =   "FHome2.frx":113C92
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":113F9C
            Tag             =   "Estimate"
            Top             =   4230
            Width           =   765
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   26
            Left            =   3900
            MouseIcon       =   "FHome2.frx":11602A
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":116334
            Tag             =   "Proposal"
            Top             =   120
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   885
            Index           =   25
            Left            =   9300
            MouseIcon       =   "FHome2.frx":11805E
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":118368
            Tag             =   "Customer"
            Top             =   4080
            Width           =   915
         End
         Begin VB.Image Task 
            Height          =   870
            Index           =   24
            Left            =   5490
            MouseIcon       =   "FHome2.frx":11AE12
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":11B11C
            Tag             =   "RFI's"
            Top             =   2130
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   1005
            Index           =   23
            Left            =   2430
            MouseIcon       =   "FHome2.frx":11CF46
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":11D250
            Tag             =   "Bids & Tenders"
            Top             =   2130
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   22
            Left            =   390
            MouseIcon       =   "FHome2.frx":11F51E
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":11F828
            Tag             =   "Vendor"
            Top             =   2190
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   6
            Left            =   4050
            MouseIcon       =   "FHome2.frx":121552
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":12185C
            Tag             =   "Inquiry"
            Top             =   6330
            Width           =   570
         End
         Begin VB.Image Image1 
            Height          =   7545
            Index           =   2
            Left            =   750
            Picture         =   "FHome2.frx":122F46
            Stretch         =   -1  'True
            Top             =   1500
            Width           =   10770
         End
      End
   End
   Begin VB.PictureBox CommandBar 
      Align           =   4  'Align Right
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   11745
      Left            =   14760
      ScaleHeight     =   11745
      ScaleWidth      =   4305
      TabIndex        =   18
      Top             =   0
      Width           =   4305
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   90
         TabIndex        =   19
         Top             =   330
         Width           =   4095
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pending Changes:"
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   1545
         End
         Begin VB.Label lblJTDPercent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(42%)"
            Height          =   195
            Left            =   3420
            TabIndex        =   36
            Top             =   2070
            Width           =   390
         End
         Begin VB.Label lblCostToComplete 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "219,388.03"
            Height          =   285
            Left            =   1710
            TabIndex        =   35
            Top             =   2310
            Width           =   1545
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cost to Complete:"
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Top             =   2310
            Width           =   1545
         End
         Begin VB.Line Line1 
            X1              =   1920
            X2              =   3420
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblJTDCost 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "150,847.22"
            Height          =   285
            Left            =   1710
            TabIndex        =   33
            Top             =   2070
            Width           =   1545
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Job to Date Cost:"
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   32
            Top             =   2070
            Width           =   1545
         End
         Begin VB.Label lblBudget 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "370,253.25"
            Height          =   285
            Left            =   1710
            TabIndex        =   31
            Top             =   1710
            Width           =   1545
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Budget:"
            Height          =   405
            Index           =   8
            Left            =   120
            TabIndex        =   30
            Top             =   1710
            Width           =   1545
         End
         Begin VB.Label lblPendingChanges 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   285
            Left            =   2490
            TabIndex        =   29
            Top             =   1440
            Width           =   1545
         End
         Begin VB.Label lblApprovedChanges 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "14,252.00"
            Height          =   285
            Left            =   1710
            TabIndex        =   28
            Top             =   1200
            Width           =   1545
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Approved Changes:"
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   1545
         End
         Begin VB.Label lblEstimate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "352,450.00"
            Height          =   285
            Left            =   1710
            TabIndex        =   26
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Original Budget:"
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label lblContract 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "424,800.00"
            Height          =   285
            Left            =   1710
            TabIndex        =   24
            Top             =   720
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stuart Olsen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   4635
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Current Contract:"
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Attachments"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   2580
            TabIndex        =   21
            Top             =   3180
            Width           =   1275
         End
         Begin VB.Image imgLink 
            Height          =   360
            Index           =   1
            Left            =   2130
            Picture         =   "FHome2.frx":22BBBC
            Top             =   3120
            Width           =   360
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contacts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   810
            TabIndex        =   20
            Top             =   3180
            Width           =   915
         End
         Begin VB.Image imgLink 
            Height          =   360
            Index           =   0
            Left            =   360
            Picture         =   "FHome2.frx":22C2A6
            Top             =   3120
            Width           =   360
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid gJobs 
         Height          =   7125
         Left            =   90
         TabIndex        =   38
         Top             =   4530
         Width           =   4110
         _cx             =   7250
         _cy             =   12568
         Appearance      =   0
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
         BackColorAlternate=   15986917
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FHome2.frx":22C990
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
         ExplorerBar     =   7
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
      Begin MSComctlLib.ImageList Icons 
         Left            =   420
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FHome2.frx":22C9DF
               Key             =   "Crystal"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Jobs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   39
         Top             =   4290
         Width           =   1320
      End
   End
   Begin VB.PictureBox WorkflowBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F0E5&
      ForeColor       =   &H80000008&
      Height          =   7995
      Index           =   2
      Left            =   120
      ScaleHeight     =   7965
      ScaleWidth      =   11295
      TabIndex        =   5
      Top             =   1185
      Width           =   11325
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   7875
         Index           =   2
         Left            =   0
         ScaleHeight     =   7875
         ScaleWidth      =   9435
         TabIndex        =   11
         Top             =   0
         Width           =   9435
         Begin VB.Image Task 
            Height          =   795
            Index           =   47
            Left            =   1620
            MouseIcon       =   "FHome2.frx":22CF79
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":22D283
            Tag             =   "Estimate"
            Top             =   2730
            Width           =   765
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   31
            Left            =   450
            MouseIcon       =   "FHome2.frx":22F311
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":22F61B
            Tag             =   "Proposal"
            Top             =   330
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   1005
            Index           =   29
            Left            =   6540
            MouseIcon       =   "FHome2.frx":231345
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":23164F
            Tag             =   "Bids & Tenders"
            Top             =   4980
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   870
            Index           =   20
            Left            =   8430
            MouseIcon       =   "FHome2.frx":23391D
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":233C27
            Tag             =   "RFI's"
            Top             =   2820
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   960
            Index           =   19
            Left            =   4680
            MouseIcon       =   "FHome2.frx":235A51
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":235D5B
            Tag             =   "Contracts & CO's"
            Top             =   330
            Width           =   795
         End
         Begin VB.Image Task 
            Height          =   885
            Index           =   16
            Left            =   2520
            MouseIcon       =   "FHome2.frx":23859D
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":2388A7
            Tag             =   "Customer"
            Top             =   270
            Width           =   915
         End
         Begin VB.Image Task 
            Height          =   960
            Index           =   15
            Left            =   8370
            MouseIcon       =   "FHome2.frx":23B351
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":23B65B
            Tag             =   "Custom Requests"
            Top             =   300
            Width           =   690
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   10
            Left            =   4830
            MouseIcon       =   "FHome2.frx":23D99D
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":23DCA7
            Tag             =   "Vendor"
            Top             =   6840
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   975
            Index           =   7
            Left            =   2580
            MouseIcon       =   "FHome2.frx":23F9D1
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":23FCDB
            Tag             =   "Purchase Orders"
            Top             =   5010
            Width           =   885
         End
         Begin VB.Image Task 
            Height          =   990
            Index           =   30
            Left            =   630
            MouseIcon       =   "FHome2.frx":242AD1
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":242DDB
            Tag             =   "Field PO's"
            Top             =   4890
            Width           =   825
         End
         Begin VB.Image Task 
            Height          =   600
            Index           =   1
            Left            =   4860
            MouseIcon       =   "FHome2.frx":24596D
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":245C77
            Tag             =   "Job"
            Top             =   2880
            Width           =   450
         End
         Begin VB.Image Image1 
            Height          =   7830
            Index           =   0
            Left            =   810
            Picture         =   "FHome2.frx":246B19
            Stretch         =   -1  'True
            Top             =   1290
            Width           =   9450
         End
      End
   End
   Begin VB.PictureBox WorkflowBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F0E5&
      ForeColor       =   &H80000008&
      Height          =   7995
      Index           =   3
      Left            =   120
      ScaleHeight     =   7965
      ScaleWidth      =   10695
      TabIndex        =   6
      Top             =   690
      Width           =   10725
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   7785
         Index           =   3
         Left            =   0
         ScaleHeight     =   7785
         ScaleWidth      =   9225
         TabIndex        =   10
         Top             =   0
         Width           =   9225
         Begin VB.Image Task 
            Height          =   795
            Index           =   5
            Left            =   1470
            MouseIcon       =   "FHome2.frx":337D43
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":33804D
            Tag             =   "Estimate"
            Top             =   2640
            Width           =   765
         End
         Begin VB.Image Task 
            Height          =   945
            Index           =   46
            Left            =   8250
            MouseIcon       =   "FHome2.frx":33A0DB
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":33A3E5
            Tag             =   "Change Orders"
            Top             =   210
            Width           =   750
         End
         Begin VB.Image Task 
            Height          =   765
            Index           =   45
            Left            =   4560
            MouseIcon       =   "FHome2.frx":33C98F
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":33CC99
            Tag             =   "Contract"
            Top             =   180
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   855
            Index           =   44
            Left            =   390
            MouseIcon       =   "FHome2.frx":33EA57
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":33ED61
            Tag             =   "Field PO's"
            Top             =   4860
            Width           =   825
         End
         Begin VB.Image Task 
            Height          =   975
            Index           =   43
            Left            =   2250
            MouseIcon       =   "FHome2.frx":34130B
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":341615
            Tag             =   "Purchase Orders"
            Top             =   4890
            Width           =   885
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   42
            Left            =   4590
            MouseIcon       =   "FHome2.frx":34440B
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":344715
            Tag             =   "Vendor"
            Top             =   6750
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   1005
            Index           =   41
            Left            =   6270
            MouseIcon       =   "FHome2.frx":34643F
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":346749
            Tag             =   "Bids & Tenders"
            Top             =   4860
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   870
            Index           =   40
            Left            =   8250
            MouseIcon       =   "FHome2.frx":348A17
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":348D21
            Tag             =   "RFI's"
            Top             =   2580
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   885
            Index           =   2
            Left            =   1350
            MouseIcon       =   "FHome2.frx":34AB4B
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":34AE55
            Tag             =   "Customer"
            Top             =   90
            Width           =   915
         End
         Begin VB.Image Task 
            Height          =   600
            Index           =   0
            Left            =   4650
            MouseIcon       =   "FHome2.frx":34D8FF
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":34DC09
            Tag             =   "Job"
            Top             =   2760
            Width           =   450
         End
         Begin VB.Image Image1 
            Height          =   7725
            Index           =   1
            Left            =   1710
            Picture         =   "FHome2.frx":34EAAB
            Stretch         =   -1  'True
            Top             =   1830
            Width           =   9165
         End
      End
   End
   Begin VB.PictureBox WorkflowBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F0E5&
      ForeColor       =   &H80000008&
      Height          =   8655
      Index           =   4
      Left            =   120
      ScaleHeight     =   8625
      ScaleWidth      =   11955
      TabIndex        =   7
      Top             =   1650
      Width           =   11985
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   8175
         Index           =   4
         Left            =   0
         ScaleHeight     =   8175
         ScaleWidth      =   9975
         TabIndex        =   12
         Top             =   0
         Width           =   9975
         Begin VB.Image Task 
            Height          =   600
            Index           =   32
            Left            =   4500
            MouseIcon       =   "FHome2.frx":435871
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":435B7B
            Tag             =   "Job"
            Top             =   4590
            Width           =   450
         End
         Begin VB.Image Task 
            Height          =   1020
            Index           =   39
            Left            =   8730
            MouseIcon       =   "FHome2.frx":436A1D
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":436D27
            Tag             =   "Cost Plus Billing"
            Top             =   4530
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   975
            Index           =   38
            Left            =   6510
            MouseIcon       =   "FHome2.frx":4394B9
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":4397C3
            Tag             =   "Progress Billing"
            Top             =   6750
            Width           =   750
         End
         Begin VB.Image Task 
            Height          =   1020
            Index           =   37
            Left            =   2310
            MouseIcon       =   "FHome2.frx":43BE9D
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":43C1A7
            Tag             =   "Time & Equipment"
            Top             =   4530
            Width           =   855
         End
         Begin VB.Image Task 
            Height          =   975
            Index           =   36
            Left            =   6540
            MouseIcon       =   "FHome2.frx":43EF99
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":43F2A3
            Tag             =   "Journal Entries"
            Top             =   2250
            Width           =   705
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   35
            Left            =   2340
            MouseIcon       =   "FHome2.frx":441775
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":441A7F
            Tag             =   "Invoices"
            Top             =   2310
            Width           =   750
         End
         Begin VB.Image Task 
            Height          =   1110
            Index           =   34
            Left            =   2250
            MouseIcon       =   "FHome2.frx":443871
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":443B7B
            Tag             =   "Invoice Approval"
            Top             =   210
            Width           =   855
         End
         Begin VB.Image Task 
            Height          =   885
            Index           =   33
            Left            =   8670
            MouseIcon       =   "FHome2.frx":446D75
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":44707F
            Tag             =   "Customer"
            Top             =   6690
            Width           =   915
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   21
            Left            =   300
            MouseIcon       =   "FHome2.frx":449B29
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":449E33
            Tag             =   "Vendor"
            Top             =   2310
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   4
            Left            =   4470
            MouseIcon       =   "FHome2.frx":44BB5D
            MousePointer    =   99  'Custom
            Picture         =   "FHome2.frx":44BE67
            Tag             =   "Inquiry"
            Top             =   6780
            Width           =   570
         End
         Begin VB.Image Image1 
            Height          =   8145
            Index           =   3
            Left            =   780
            Picture         =   "FHome2.frx":44D551
            Stretch         =   -1  'True
            Top             =   1350
            Width           =   9915
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "res est"
      Height          =   195
      Index           =   0
      Left            =   11820
      TabIndex        =   17
      Top             =   1410
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "jc acct"
      Height          =   195
      Index           =   4
      Left            =   12120
      TabIndex        =   16
      Top             =   1650
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "comm est"
      Height          =   195
      Index           =   1
      Left            =   11190
      TabIndex        =   15
      Top             =   930
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "comm pm"
      Height          =   195
      Index           =   3
      Left            =   10860
      TabIndex        =   14
      Top             =   690
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "res pm"
      Height          =   195
      Index           =   2
      Left            =   11460
      TabIndex        =   13
      Top             =   1170
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      Caption         =   "JOB COST ACCOUNTING"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   9990
      TabIndex        =   2
      Top             =   60
      Width           =   4905
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      BackColor       =   &H00F3F0E5&
      Caption         =   "PROJECT MANAGEMENT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   1
      Top             =   60
      Width           =   4905
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      Caption         =   "ESTIMATING"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4905
   End
End
Attribute VB_Name = "FHome2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FHome2::"

Private mCommercial As Boolean
Private mTab As Integer
Private mview As Integer





Private Sub Form_Initialize()
    Call Form_Resize
End Sub

Private Sub Form_Load()
On Error Resume Next
    mTab = Val(IniGet(AppIni, "FHome2", "Tab"))
    Call IniGetGrid(Me, gJobs)
    Call LoadCurJob
    Call LoadJobs
    Call ConfigForm
End Sub


Private Sub gJobs_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
On Error Resume Next
    Dim s As String
    If OldRowSel <> NewRowSel Then
    With gJobs
        s = .TextMatrix(.Row, .ColIndex("Job"))
        If s <> "" Then
            FMain.CurrentJob = s
            Call LoadCurJob
        End If
    End With
    End If
End Sub
Private Sub SetValues(Description As String, Contract As Double, Estimate As Double, ApprovedChanges As Double, PendingChanges As Double, JTDCost As Double)
On Error Resume Next
'Exit Sub
    Dim Budget As Double
    Dim CostToComplete As Double
    
    Budget = Estimate + ApprovedChanges '+ PendingChanges
    CostToComplete = Budget - JTDCost
    
    lblDescription = Description
    lblContract = format(Contract, "#,##0.00")
    lblEstimate = format(Estimate, "#,##0.00")
    lblApprovedChanges = format(ApprovedChanges, "#,##0.00")
    lblPendingChanges = format(PendingChanges, "#,##0.00")
    lblBudget = format(Budget, "#,##0.00")
    lblJTDCost = format(JTDCost, "#,##0.00")
    lblCostToComplete = format(CostToComplete, "#,##0.00")
    
    If (Budget) = 0 Then
        lblJTDPercent = ""
    Else
        lblJTDPercent = "(" & Round(Val(JTDCost / Budget * 100), 1) & "%)"
    End If
    
End Sub
Private Sub gJobs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = vbRightButton Then
        Call FMain.ShowColumnMenu(gJobs)
    End If
End Sub

Private Sub lblLink_Click(Index As Integer)
    Call RunLink(Index)
End Sub
Private Sub RunLink(Index As Integer)
    If Index = 0 Then
        If FMain.CurrentJob <> "" Then
            Call FContacts.ShowForm(FMain.CurrentJob)
        End If
    Else
        If FMain.CurrentJob <> "" Then
            Call HFApp.RunTask("EditAttachments|J~" & FMain.CurrentJob & "|Job Documents")
        End If
    End If
End Sub
Private Sub lblTab_Click(Index As Integer)
On Error Resume Next
    Dim i As Long
    mTab = Index
    Select Case Index
        Case 0:  mview = IIf(mCommercial, 1, 0)
        Case 1:  mview = IIf(mCommercial, 3, 2)
        Case 2:  mview = 4
    End Select

    For i = 0 To 2
        lblTab(i).BackColor = IIf(i = mTab, &HF3F0E5, vbButtonFace)
        lblTab(i).FontBold = i = mTab
    Next
    For i = 0 To 4
        WorkflowBG(i).Visible = i = mview
        Workflow(i).BorderStyle = 0
        Workflow(i).BackColor = &HF3F0E5
    Next
    
    'Call Form_Resize
    
End Sub

Public Sub ConfigForm()
On Error Resume Next
    mCommercial = HFApp.Options.ValueByName("BuilderType") = "Commercial"
    Call lblTab_Click(mTab)
End Sub

Private Sub Form_Resize()
On Error GoTo eh
    Const margin = 60
    Dim i As Long


    'form is minimized
    If Me.ScaleHeight < 1 Then Exit Sub

    CommandBar.ZOrder 0
    If Me.Height < 8730 Then Me.Height = 8730
    If Me.Width < 15390 Then Me.Width = 15390

    On Error Resume Next
    gJobs.Height = Me.ScaleHeight - gJobs.Top - 90
    On Error GoTo eh


    'tabs
    lblTab(0).Move margin, margin, (Me.ScaleWidth - CommandBar.Width - 4 * margin) / 3
    lblTab(1).Move lblTab(0).Left + lblTab(0).Width + margin, margin, lblTab(0).Width
    lblTab(2).Move lblTab(1).Left + lblTab(1).Width + margin, margin, lblTab(0).Width


On Error GoTo eh
    
    'workflowBG
    WorkflowBG(0).Move margin, lblTab(0).Height + 2 * margin, Me.ScaleWidth - CommandBar.Width - 2 * margin, Me.ScaleHeight - WorkflowBG(0).Top - margin
    For i = 1 To 4
        WorkflowBG(i).Move WorkflowBG(0).Left, WorkflowBG(0).Top, WorkflowBG(0).Width, WorkflowBG(0).Height
    Next

    'workflow
    For i = 0 To 4
        Workflow(i).BorderStyle = 0
        WorkflowBG(i).BorderStyle = 0
        Workflow(i).Move (WorkflowBG(i).Width - Workflow(i).Width) / 2, (WorkflowBG(i).Height - Workflow(i).Height) / 2
        Image1(i).Move 0, 0
    Next
Exit Sub
eh: Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call IniPut(AppIni, "FHome2", "Tab", mTab)
    Call SetErrorMode(SEM_NOERRORS)
    Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
    Set FHome2 = Nothing
End Sub

Private Sub Task_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo eh
    If Task(Index).Tag = "" Or Button <> vbLeftButton Then Exit Sub
    Select Case Task(Index).Tag
        Case "Bids & Tenders":            Call FMain.RunTask("Bids", Shift = vbCtrlMask)
        Case "Change Orders":             Call FMain.RunTask("Contracts", Shift = vbCtrlMask)
        Case "Contract":                  Call FMain.RunTask("Contracts", Shift = vbCtrlMask)
        Case "Contracts & CO's":          Call FMain.RunTask("Inbox", Shift = vbCtrlMask)
        Case "Progress Billing":          Call FMain.RunTask("Billing", Shift = vbCtrlMask)
        Case "Custom Requests":           Call FInboxCustomQuote.Show(vbModal, FMain)
        Case "Customer":                  Call HFApp.RunTask("EditCustomer|" & FMain.CurrentCustomer)
        Case "Inquiry":
        
        Case "Cost Plus Billing":         Call Shell(PathAppend(App.Path, "hfworkticket.exe"))
        Case "Invoice Approval":          Call Shell(PathAppend(App.Path, "hfpayables.exe"))
        Case "Invoices":                  Call Shell(PathAppend(App.Path, "hfpayables.exe"))
        
        Case "Proposal":                  Call FMain.RunTask("Prepare Job Quote", Shift = vbCtrlMask)
        Case "RFI's":                     Call FMain.RunTask("RFI's", Shift = vbCtrlMask)
        Case "Estimate":                  Call FMain.RunTask("issue budgets", Shift = vbCtrlMask)
        Case "Time & Equipment":          Call Shell(PathAppend(App.Path, "hfworkticket.exe"))
        Case "Journal Entries":           Call HFApp.RunTask("JournalEntries")
        Case "Communities":               Call FMain.RunTask("Community Setup", Shift = vbCtrlMask)
        Case "Job":                       Call FMain.RunTask("Job setup", Shift = vbCtrlMask)
        Case "Default Vendors":           Call FMain.RunTask("default vendors", Shift = vbCtrlMask)
        Case "Field PO's":                Call FMain.RunTask("field po's", Shift = vbCtrlMask)
        Case "Item DB":                   Call FMain.RunTask("edit item database", Shift = vbCtrlMask)
        Case "Model & Option Library":    Call FMain.RunTask("edit models and options", Shift = vbCtrlMask)
        Case "Purchase Orders":           Call FMain.RunTask("issue po's", Shift = vbCtrlMask)
        Case "Salespricing Worksheets":   Call FMain.RunTask("open an estimating worksheet", Shift = vbCtrlMask)
        Case "Vendor":                    Call FMain.RunTask("vendor setup", Shift = vbCtrlMask)
        Case "Vendor Pricelists":         Call FMain.RunTask("edit item prices", Shift = vbCtrlMask)
        Case Else
            
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "Task_MouseDown")
End Sub


Public Sub LoadCurJob()
On Error Resume Next

Dim t As Single
Dim d As String
t = Timer()
    
    Dim s As String
    Dim rs As Recordset
    
    Dim Contract        As Double
    Dim Estimate        As Double
    Dim ApprovedChanges As Double
    Dim PendingChanges  As Double
    Dim JTDCost         As Double
    
    s = ""
    s = s & "SELECT SUM(price)" & vbCrLf
    s = s & "  FROM billingitems" & vbCrLf
    s = s & " WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "   AND Job=" & DbQuote(Str, FMain.CurrentJob) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    Contract = Val("" & rs(0))
d = d & "billing: " & Timer - t & vbCrLf
t = Timer
    
    s = ""
    s = s & "SELECT SUM(CASE WHEN IsChange=0 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) Budgeted" & vbCrLf
    s = s & "      ,SUM(CASE WHEN IsChange=1 and ischangerequest=0 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) ApprovedChanges" & vbCrLf
    s = s & "      ,SUM(CASE WHEN IsChange=1 and ischangerequest=1 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) PendingChanges" & vbCrLf
    s = s & "  FROM Estimateditems" & vbCrLf
    s = s & " WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "   AND Job_no=" & DbQuote(Str, FMain.CurrentJob) & vbCrLf

    Set rs = HFApp.SqlExec(s)
    Estimate = Val("" & rs(0))
    ApprovedChanges = Val("" & rs(1))
    PendingChanges = Val("" & rs(2))
    
d = d & "estim: " & Timer - t & vbCrLf
t = Timer
    
    
    s = ""
    s = s & "select sum(pretax+jctax)" & vbCrLf
    s = s & "  from JCTransactions" & vbCrLf
    s = s & " WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "   AND Job=" & DbQuote(Str, FMain.CurrentJob) & vbCrLf
    s = s & "   and TransactionType<>'Estimate' and TransactionType<>'Purchaseorder'"
    Set rs = HFApp.SqlExec(s)
    JTDCost = Val("" & rs(0))

d = d & "jctran: " & Timer - t & vbCrLf
t = Timer
    
    Call SetValues(HFApp.FormatJob(FMain.CurrentJob) & "   " & FMain.CurrentJobDesc, Contract, Estimate, ApprovedChanges, PendingChanges, JTDCost)
    
d = d & "final: " & Timer - t & vbCrLf
t = Timer

'MsgBox d
    
End Sub

Public Sub LoadJobs()
    Dim i As Long
    Dim rs As Recordset
    Dim s As String


    With gJobs
    
        If Not .DataSource Is Nothing Then Call IniPutGrid(Me, gJobs)
        Set .DataSource = Nothing
        s = "select wj.* from workflowjobs wj"
        s = s & " join tbljobs j on (wj.DivisionID = j.DivisionID and wj.job = j.job_no)" & vbCrLf
        s = s & " left outer join DivisionCommunities d on (d.Community = j.Community)" & vbCrLf
        s = s & " where (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
        s = s & " and j.DivisionID = " & HFApp.DivisionID
        s = s & " order by wj.job"
        Set rs = HFApp.SqlExec(s)
        .Rows = 1
        .FixedRows = 1
        .DataMode = flexDMFree
        .AutoSearch = flexSearchFromCursor
        .Editable = flexEDNone
    
        Set .DataSource = rs

        On Error Resume Next
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
        Next

        .Row = 1
        
        If .Row > 0 Then
        FMain.CurrentJob = .TextMatrix(.Row, .ColIndex("Job"))
        End If
        
    End With
    Call IniGetGrid(Me, gJobs)
    gJobs.ColHidden(gJobs.ColIndex("DivisionID")) = True
    
End Sub


