VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Begin VB.Form FHome2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Home"
   ClientHeight    =   11745
   ClientLeft      =   675
   ClientTop       =   420
   ClientWidth     =   19065
   Icon            =   "FHome3.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11745
   ScaleWidth      =   19065
   Begin VB.PictureBox WorkflowBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F0E5&
      ForeColor       =   &H80000008&
      Height          =   9225
      Index           =   0
      Left            =   120
      ScaleHeight     =   9195
      ScaleWidth      =   11655
      TabIndex        =   4
      Top             =   1410
      Width           =   11685
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   8145
         Index           =   0
         Left            =   0
         ScaleHeight     =   8145
         ScaleWidth      =   9645
         TabIndex        =   9
         Top             =   0
         Width           =   9645
         Begin VB.Image Task 
            Height          =   1035
            Index           =   14
            Left            =   8310
            MouseIcon       =   "FHome3.frx":000C
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":0316
            Tag             =   "Salespricing Worksheets"
            Top             =   240
            Width           =   1050
         End
         Begin VB.Image Task 
            Height          =   990
            Index           =   12
            Left            =   2790
            MouseIcon       =   "FHome3.frx":3C7C
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":3F86
            Tag             =   "Model & Option Library"
            Top             =   2190
            Width           =   1155
         End
         Begin VB.Image Task 
            Height          =   720
            Index           =   17
            Left            =   5580
            MouseIcon       =   "FHome3.frx":7B98
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":7EA2
            Tag             =   "Communities"
            Top             =   2310
            Width           =   945
         End
         Begin VB.Image Task 
            Height          =   930
            Index           =   9
            Left            =   8490
            MouseIcon       =   "FHome3.frx":A2E4
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":A5EE
            Tag             =   "Vendor Pricelists"
            Top             =   4590
            Width           =   690
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   8
            Left            =   5670
            MouseIcon       =   "FHome3.frx":C818
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":CB22
            Tag             =   "Vendor"
            Top             =   6810
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   960
            Index           =   18
            Left            =   5640
            MouseIcon       =   "FHome3.frx":E84C
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":EB56
            Tag             =   "Default Vendors"
            Top             =   4620
            Width           =   765
         End
         Begin VB.Image Task 
            Height          =   1035
            Index           =   13
            Left            =   420
            MouseIcon       =   "FHome3.frx":11298
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":115A2
            Tag             =   "Standard Substitutions"
            Top             =   6720
            Width           =   1020
         End
         Begin VB.Image Task 
            Height          =   765
            Index           =   11
            Left            =   390
            MouseIcon       =   "FHome3.frx":14CE0
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":14FEA
            Tag             =   "Item DB"
            Top             =   4530
            Width           =   630
         End
         Begin VB.Image Image1 
            Height          =   8115
            Index           =   4
            Left            =   2160
            Picture         =   "FHome3.frx":169AC
            Stretch         =   -1  'True
            Top             =   2070
            Width           =   9630
         End
      End
   End
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl CommandBar 
      Align           =   4  'Align Right
      Height          =   11745
      Left            =   13980
      TabIndex        =   3
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   20717
      BackColorEnd    =   16777215
      BackColorStart  =   16777215
      Begin VSFlex8Ctl.VSFlexGrid gJobs 
         Height          =   4815
         Left            =   720
         TabIndex        =   14
         Top             =   3510
         Width           =   4230
         _cx             =   7461
         _cy             =   8493
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
         FormatString    =   $"FHome3.frx":115456
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
      Begin HFEst.CustomerPanel CustomerPanel 
         Height          =   2355
         Left            =   720
         TabIndex        =   15
         Top             =   870
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   4154
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageList Icons 
         Left            =   0
         Top             =   0
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
               Picture         =   "FHome3.frx":1154A5
               Key             =   "Crystal"
            EndProperty
         EndProperty
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
      TabIndex        =   8
      Top             =   1650
      Width           =   11985
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   8175
         Index           =   4
         Left            =   0
         ScaleHeight     =   8175
         ScaleWidth      =   9975
         TabIndex        =   13
         Top             =   0
         Width           =   9975
         Begin VB.Image Task 
            Height          =   600
            Index           =   32
            Left            =   4500
            MouseIcon       =   "FHome3.frx":115A3F
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":115D49
            Tag             =   "Job"
            Top             =   4590
            Width           =   450
         End
         Begin VB.Image Task 
            Height          =   1020
            Index           =   39
            Left            =   8730
            MouseIcon       =   "FHome3.frx":116BEB
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":116EF5
            Tag             =   "Cost Plus Billing"
            Top             =   4530
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   975
            Index           =   38
            Left            =   6510
            MouseIcon       =   "FHome3.frx":119687
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":119991
            Tag             =   "Progress Billing"
            Top             =   6750
            Width           =   750
         End
         Begin VB.Image Task 
            Height          =   1020
            Index           =   37
            Left            =   2310
            MouseIcon       =   "FHome3.frx":11C06B
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":11C375
            Tag             =   "Time & Equipment"
            Top             =   4530
            Width           =   855
         End
         Begin VB.Image Task 
            Height          =   975
            Index           =   36
            Left            =   6540
            MouseIcon       =   "FHome3.frx":11F167
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":11F471
            Tag             =   "Journal Entries"
            Top             =   2250
            Width           =   705
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   35
            Left            =   2340
            MouseIcon       =   "FHome3.frx":121943
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":121C4D
            Tag             =   "Invoices"
            Top             =   2310
            Width           =   750
         End
         Begin VB.Image Task 
            Height          =   1110
            Index           =   34
            Left            =   2250
            MouseIcon       =   "FHome3.frx":123A3F
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":123D49
            Tag             =   "Invoice Approval"
            Top             =   210
            Width           =   855
         End
         Begin VB.Image Task 
            Height          =   885
            Index           =   33
            Left            =   8670
            MouseIcon       =   "FHome3.frx":126F43
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":12724D
            Tag             =   "Customer"
            Top             =   6690
            Width           =   915
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   21
            Left            =   300
            MouseIcon       =   "FHome3.frx":129CF7
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":12A001
            Tag             =   "Vendor"
            Top             =   2310
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   4
            Left            =   4470
            MouseIcon       =   "FHome3.frx":12BD2B
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":12C035
            Tag             =   "Inquiry"
            Top             =   6780
            Width           =   570
         End
         Begin VB.Image Image1 
            Height          =   8145
            Index           =   3
            Left            =   780
            Picture         =   "FHome3.frx":12D71F
            Stretch         =   -1  'True
            Top             =   1350
            Width           =   9915
         End
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
      TabIndex        =   6
      Top             =   1170
      Width           =   11325
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   7875
         Index           =   2
         Left            =   0
         ScaleHeight     =   7875
         ScaleWidth      =   9435
         TabIndex        =   12
         Top             =   0
         Width           =   9435
         Begin VB.Image Task 
            Height          =   795
            Index           =   47
            Left            =   1620
            MouseIcon       =   "FHome3.frx":2347A1
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":234AAB
            Tag             =   "Estimate"
            Top             =   2730
            Width           =   765
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   31
            Left            =   450
            MouseIcon       =   "FHome3.frx":236B39
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":236E43
            Tag             =   "Proposal"
            Top             =   330
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   1005
            Index           =   29
            Left            =   6540
            MouseIcon       =   "FHome3.frx":238B6D
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":238E77
            Tag             =   "Bids & Tenders"
            Top             =   4980
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   870
            Index           =   20
            Left            =   8430
            MouseIcon       =   "FHome3.frx":23B145
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":23B44F
            Tag             =   "RFI's"
            Top             =   2820
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   960
            Index           =   19
            Left            =   4680
            MouseIcon       =   "FHome3.frx":23D279
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":23D583
            Tag             =   "Contracts & CO's"
            Top             =   330
            Width           =   795
         End
         Begin VB.Image Task 
            Height          =   885
            Index           =   16
            Left            =   2520
            MouseIcon       =   "FHome3.frx":23FDC5
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":2400CF
            Tag             =   "Customer"
            Top             =   270
            Width           =   915
         End
         Begin VB.Image Task 
            Height          =   960
            Index           =   15
            Left            =   8370
            MouseIcon       =   "FHome3.frx":242B79
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":242E83
            Tag             =   "Custom Requests"
            Top             =   300
            Width           =   690
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   10
            Left            =   4830
            MouseIcon       =   "FHome3.frx":2451C5
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":2454CF
            Tag             =   "Vendor"
            Top             =   6840
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   975
            Index           =   7
            Left            =   2580
            MouseIcon       =   "FHome3.frx":2471F9
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":247503
            Tag             =   "Purchase Orders"
            Top             =   5010
            Width           =   885
         End
         Begin VB.Image Task 
            Height          =   855
            Index           =   30
            Left            =   630
            MouseIcon       =   "FHome3.frx":24A2F9
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":24A603
            Tag             =   "Field PO's"
            Top             =   4890
            Width           =   825
         End
         Begin VB.Image Task 
            Height          =   600
            Index           =   1
            Left            =   4860
            MouseIcon       =   "FHome3.frx":24CBAD
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":24CEB7
            Tag             =   "Job"
            Top             =   2880
            Width           =   450
         End
         Begin VB.Image Image1 
            Height          =   7830
            Index           =   0
            Left            =   810
            Picture         =   "FHome3.frx":24DD59
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
      Height          =   7875
      Index           =   1
      Left            =   120
      ScaleHeight     =   7845
      ScaleWidth      =   11025
      TabIndex        =   5
      Top             =   930
      Width           =   11055
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   7605
         Index           =   1
         Left            =   0
         ScaleHeight     =   7605
         ScaleWidth      =   10785
         TabIndex        =   10
         Top             =   0
         Width           =   10785
         Begin VB.Image Task 
            Height          =   810
            Index           =   28
            Left            =   6720
            MouseIcon       =   "FHome3.frx":33EF83
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":33F28D
            Tag             =   "Contract"
            Top             =   4140
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   795
            Index           =   27
            Left            =   3960
            MouseIcon       =   "FHome3.frx":341207
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":341511
            Tag             =   "Estimate"
            Top             =   4230
            Width           =   765
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   26
            Left            =   3900
            MouseIcon       =   "FHome3.frx":34359F
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":3438A9
            Tag             =   "Proposal"
            Top             =   120
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   885
            Index           =   25
            Left            =   9300
            MouseIcon       =   "FHome3.frx":3455D3
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":3458DD
            Tag             =   "Customer"
            Top             =   4080
            Width           =   915
         End
         Begin VB.Image Task 
            Height          =   870
            Index           =   24
            Left            =   5490
            MouseIcon       =   "FHome3.frx":348387
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":348691
            Tag             =   "RFI's"
            Top             =   2130
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   1005
            Index           =   23
            Left            =   2430
            MouseIcon       =   "FHome3.frx":34A4BB
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":34A7C5
            Tag             =   "Bids & Tenders"
            Top             =   2130
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   22
            Left            =   390
            MouseIcon       =   "FHome3.frx":34CA93
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":34CD9D
            Tag             =   "Vendor"
            Top             =   2190
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   6
            Left            =   4050
            MouseIcon       =   "FHome3.frx":34EAC7
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":34EDD1
            Tag             =   "Inquiry"
            Top             =   6330
            Width           =   570
         End
         Begin VB.Image Image1 
            Height          =   7545
            Index           =   2
            Left            =   750
            Picture         =   "FHome3.frx":3504BB
            Stretch         =   -1  'True
            Top             =   1500
            Width           =   10770
         End
      End
   End
   Begin VB.PictureBox WorkflowBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F0E5&
      ForeColor       =   &H80000008&
      Height          =   7935
      Index           =   3
      Left            =   120
      ScaleHeight     =   7905
      ScaleWidth      =   10695
      TabIndex        =   7
      Top             =   690
      Width           =   10725
      Begin VB.PictureBox Workflow 
         BorderStyle     =   0  'None
         Height          =   7785
         Index           =   3
         Left            =   0
         ScaleHeight     =   7785
         ScaleWidth      =   9225
         TabIndex        =   11
         Top             =   0
         Width           =   9225
         Begin VB.Image Task 
            Height          =   945
            Index           =   46
            Left            =   8250
            MouseIcon       =   "FHome3.frx":459131
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":45943B
            Tag             =   "Change Orders"
            Top             =   210
            Width           =   750
         End
         Begin VB.Image Task 
            Height          =   765
            Index           =   45
            Left            =   4560
            MouseIcon       =   "FHome3.frx":45B9E5
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":45BCEF
            Tag             =   "Contract"
            Top             =   180
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   855
            Index           =   44
            Left            =   390
            MouseIcon       =   "FHome3.frx":45DAAD
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":45DDB7
            Tag             =   "Field PO's"
            Top             =   4860
            Width           =   825
         End
         Begin VB.Image Task 
            Height          =   975
            Index           =   43
            Left            =   2250
            MouseIcon       =   "FHome3.frx":460361
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":46066B
            Tag             =   "Purchase Orders"
            Top             =   4890
            Width           =   885
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   42
            Left            =   4590
            MouseIcon       =   "FHome3.frx":463461
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":46376B
            Tag             =   "Vendor"
            Top             =   6750
            Width           =   735
         End
         Begin VB.Image Task 
            Height          =   1005
            Index           =   41
            Left            =   6270
            MouseIcon       =   "FHome3.frx":465495
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":46579F
            Tag             =   "Bids & Tenders"
            Top             =   4860
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   870
            Index           =   40
            Left            =   8250
            MouseIcon       =   "FHome3.frx":467A6D
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":467D77
            Tag             =   "RFI's"
            Top             =   2580
            Width           =   660
         End
         Begin VB.Image Task 
            Height          =   885
            Index           =   2
            Left            =   1350
            MouseIcon       =   "FHome3.frx":469BA1
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":469EAB
            Tag             =   "Customer"
            Top             =   90
            Width           =   915
         End
         Begin VB.Image Task 
            Height          =   600
            Index           =   0
            Left            =   4650
            MouseIcon       =   "FHome3.frx":46C955
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":46CC5F
            Tag             =   "Job"
            Top             =   2760
            Width           =   450
         End
         Begin VB.Image Task 
            Height          =   750
            Index           =   3
            Left            =   1470
            MouseIcon       =   "FHome3.frx":46DB01
            MousePointer    =   99  'Custom
            Picture         =   "FHome3.frx":46DE0B
            Tag             =   "Inquiry"
            Top             =   2640
            Width           =   570
         End
         Begin VB.Image Image1 
            Height          =   7725
            Index           =   1
            Left            =   1710
            Picture         =   "FHome3.frx":46F4F5
            Stretch         =   -1  'True
            Top             =   1830
            Width           =   9165
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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

Private mCommercial As Boolean
Private mTab As Integer
Private mview As Integer



Private Sub CommandBar_BarClick(Bar As vbalExplorerBarLib6.cExplorerBar)
    Call Form_Resize
End Sub

Private Sub CommandBar_SettingChange()
    Call Form_Resize
End Sub

Private Sub CommandBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
Dim f As Form
    Select Case True
        Case itm.tag = "RPT"
            Set f = New FRptViewer
            Call f.ShowReport(itm.Key, True, False, "Job", FMain.CurrentJob)
    
    End Select
End Sub

Private Sub Form_Load()
    mCommercial = HFApp.Options.ValueByName("BuilderType") = "Commercial"
    mTab = Val(IniGet(AppIni, "FHome2", "Tab"))
    Call IniGetGrid(Me, gJobs)
    Call ConfigBar
    Call ConfigForm
    Call lblTab_Click(mTab)
End Sub


Private Sub gJobs_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
On Error Resume Next
    Dim s As String
    If OldRowSel <> NewRowSel Then
    With gJobs
        s = .TextMatrix(.Row, .ColIndex("Job"))
        If s <> "" Then FMain.CurrentJob = s
    End With
    End If
End Sub

Private Sub gJobs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call FMain.ShowColumnMenu(gJobs)
    End If
End Sub

Private Sub lblTab_Click(Index As Integer)
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
    
    Call LoadReports
    
    Call Form_Resize
    
End Sub

Public Sub ConfigForm()
'    'hide sales stuff if not installed
'    Task(1).Visible = HFApp.Options(SalesSystem) <> SalesSystems.asNone
'    Task(27).Visible = Not Task(1).Visible
'
'    Task(24).Visible = HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440
'    Flow(14).Visible = HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440
'
'    'hide accounting stuff if not installed
'    Task(4).Visible = HFApp.Options(AccountingSystem) <> AccountingSystems.asNone
'    Task(28).Visible = Not Task(4).Visible
'
'    Task(6).Visible = HFApp.Options(AccountingSystem) <> AccountingSystems.asNone
'    Task(29).Visible = Not Task(6).Visible
'
'
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long


    'form is minimized
    If Me.ScaleHeight < 1 Then Exit Sub

    CommandBar.ZOrder 0
    If Me.Height < 8730 Then Me.Height = 8730
    If Me.Width < 15390 Then Me.Width = 15390

    On Error Resume Next
    CommandBar.Redraw = False
    gJobs.Height = Me.ScaleHeight - gJobs.Top - 360
    CommandBar.Redraw = True
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

eh: Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutGrid(Me, gJobs)
    Call IniPut(AppIni, "FHome2", "Tab", mTab)
End Sub

Private Sub Task_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Task(Index).tag = "" Or Button <> vbLeftButton Then Exit Sub
    Select Case Task(Index).tag
        Case "Bids & Tenders":            Call FMain.RunTask("Bids", Shift = vbCtrlMask)
        Case "Change Orders":             Call FMain.RunTask("Contracts", Shift = vbCtrlMask)
        Case "Contract":                  Call FMain.RunTask("Contracts", Shift = vbCtrlMask)
        Case "Contracts & CO's":          Call FMain.RunTask("Inbox", Shift = vbCtrlMask)
        Case "Cost Plus Billing":         Call FMain.RunTask("Billing", Shift = vbCtrlMask)
        Case "Progress Billing":          Call FMain.RunTask("Billing", Shift = vbCtrlMask)
        Case "Custom Requests":           Call FInboxCustomQuote.Show(vbModal, FMain)
        Case "Customer":                  Call HFApp.RunTask("EditCustomer|" & FMain.CurrentCustomer)
        Case "Inquiry":
        
        Case "Invoice Approval":          Call Shell(PathAppend(App.Path, "hfpayables.exe"))
        Case "Invoices":                  Call Shell(PathAppend(App.Path, "hfpayables.exe"))
        
        Case "Proposal":                  Call FMain.RunTask("Prepare Job Quote", Shift = vbCtrlMask)
        Case "RFI's":
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
        Case "Standard Substitutions":    Call FMain.RunTask("community standards", Shift = vbCtrlMask)
        Case "Vendor":                    Call FMain.RunTask("vendor setup", Shift = vbCtrlMask)
        Case "Vendor Pricelists":         Call FMain.RunTask("edit item prices", Shift = vbCtrlMask)
        Case Else
            Stop
    End Select
End Sub

Private Sub ConfigBar()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem


    With CommandBar

        .ImageList = Icons
        .UseExplorerStyle = False
        .UseExplorerTransitionStyle = True


        Set cBar = .Bars.Add(, "CURJOB", "4 Cougarstone Way")
        cBar.IsSpecial = True
        Set cItem = cBar.items.Add(, "CustomerPanel", , , eItemControlPlaceHolder)
        cItem.SpacingAfter = 6
        cItem.Control = CustomerPanel
        Call LoadCurJob

'        Set cBar = .Bars.Add(, "TASKS", "Incomplete Tasks")
'        cBar.IsSpecial = True
'        Call LoadTasks


        Set cBar = .Bars.Add(, "RPTS")
        cBar.IsSpecial = True
'        cBar.State = eBarCollapsed
        Call LoadReports


        Set cBar = .Bars.Add(, "GRD", "Open Projects")
        cBar.IsSpecial = True
        cBar.items.Add(, "grid", , , eItemControlPlaceHolder).Control = gJobs
        Call LoadJobs

    End With
End Sub

Public Sub LoadCurJob()
    Dim s As String
    Dim rs As Recordset
    Dim cBar As cExplorerBar
    
    Dim Contract        As Double
    Dim Estimate        As Double
    Dim ApprovedChanges As Double
    Dim PendingChanges  As Double
    Dim JTDCost         As Double
    
    s = ""
    s = s & "SELECT SUM(price)" & vbCrLf
    s = s & "  FROM billingitems" & vbCrLf
    s = s & " WHERE Job=" & DbQuote(Str, FMain.CurrentJob) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    Contract = Val("" & rs(0))
    
    s = ""
    s = s & "SELECT SUM(CASE WHEN IsChange=0 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) Budgeted" & vbCrLf
    s = s & "      ,SUM(CASE WHEN IsChange=1 and ischangerequest=0 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) ApprovedChanges" & vbCrLf
    s = s & "      ,SUM(CASE WHEN IsChange=1 and ischangerequest=1 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) PendingChanges" & vbCrLf
    s = s & "  FROM Estimateditems" & vbCrLf
    s = s & " WHERE Job_No=" & DbQuote(Str, FMain.CurrentJob) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    Estimate = Val("" & rs(0))
    ApprovedChanges = Val("" & rs(1))
    PendingChanges = Val("" & rs(2))
    
    
    
    s = ""
    s = s & "select sum(pretax+jctax)" & vbCrLf
    s = s & "  from JCTransactions" & vbCrLf
    s = s & " where job=" & DbQuote(Str, FMain.CurrentJob) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    JTDCost = Val("" & rs(0))

    
    CommandBar.Redraw = False
    Set cBar = CommandBar.Bars("CURJOB")
    cBar.Title = HFApp.FormatJob(FMain.CurrentJob) & "      " & FMain.CurrentJobDesc
    Call CustomerPanel.SetValues(FMain.CurrentCustomerDesc, Contract, Estimate, ApprovedChanges, PendingChanges, JTDCost)
    CommandBar.Redraw = True
    
End Sub

Public Sub LoadJobs()
    Dim i As Long
    Dim rs As Recordset


    With gJobs
    
        If Not .DataSource Is Nothing Then Call IniPutGrid(Me, gJobs)
        Set .DataSource = Nothing
        Set rs = HFApp.SqlExec("select * from workflowjobs")
        .Rows = 1
        .FixedRows = 1
        .DataMode = flexDMFree
        .AutoSearch = flexSearchFromCursor
        .Editable = flexEDNone
    
        Set .DataSource = rs
    
        On Error Resume Next
        For i = 0 To .cols - 1
            .ColKey(i) = .TextMatrix(0, i)
        Next
        
        .Row = 1
        
        FMain.CurrentJob = .TextMatrix(.Row, .ColIndex("Job"))
        
    End With
    Call IniGetGrid(Me, gJobs)
    
End Sub

Private Sub LoadReports()
    Dim i As Long
    Dim cBar As cExplorerBar
    Set cBar = CommandBar.Bars("RPTS")
    Dim sFolder As String
    Dim sFile As String


    'do not use cbar.items.clear, it doesnt work
    For i = cBar.items.Count To 1 Step -1
        Call cBar.items.Remove(i)
    Next


    Select Case mTab
        Case 0
            cBar.Title = "Estimating Reports"
            sFolder = IIf(mCommercial, "Commercial", "Residential") & " Estimating Reports"
        Case 1
            cBar.Title = "Project Management Reports"
            sFolder = IIf(mCommercial, "Commercial", "Residential") & " PM Reports"
        Case 2
            cBar.Title = "Accounting Reports"
            sFolder = "JC Reports"
    End Select
    
    sFolder = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\workflows", sFolder)
    sFile = Dir(sFolder & "\*.rpt", vbDirectory, True)
    While sFile <> ""
        If sFile <> "." And sFile <> ".." Then
            cBar.items.Add(, PathAppend(sFolder, sFile), FileName(sFile), ImageIndex(Icons, "Crystal"), eItemLink).tag = "RPT"
        End If
        sFile = Dir
    Wend

End Sub

Private Sub LoadTasks()
    Dim i As Long
    Dim cBar As cExplorerBar
    Set cBar = CommandBar.Bars("TASKS")

    'do not use cbar.items.clear, it doesnt work
    For i = cBar.items.Count To 1 Step -1
        Call cBar.items.Remove(i)
    Next

    cBar.items.Add(, "RFISxx", "RFI's", , eItemText).Bold = True
    cBar.items.Add(, "RFIS1", "   Blueprint Clarification", , eItemLink).tag = ""
    cBar.items.Add(, "RFIS2", "   Completion Date (Sunshine)", , eItemLink).tag = ""
    cBar.items.Add(, "RFIS3", "   Blueprint Clarificationasdfasdf", , eItemLink).SpacingAfter = 3

    cBar.items.Add(, "BIDS", "Bids && Tenders", , eItemText).Bold = True
    cBar.items.Add(, "B1", "   Blueprint Clarification", , eItemLink).tag = ""
    cBar.items.Add(, "B2", "   Completion Date (Sunshine)", , eItemLink).tag = ""
    cBar.items.Add(, "B3", "   Blueprint Clarificationasdfasdf", , eItemLink).tag = ""

End Sub


