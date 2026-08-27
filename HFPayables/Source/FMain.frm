VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form FMain 
   Caption         =   "Payables Desk - Enter Invoices"
   ClientHeight    =   10665
   ClientLeft      =   5805
   ClientTop       =   3660
   ClientWidth     =   13605
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   13605
   Begin VB.PictureBox picWarningBanner 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   13605
      TabIndex        =   81
      Top             =   0
      Width           =   13605
      Begin VB.Label lblWarningBanner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FMain.frx":08CA
         Height          =   195
         Left            =   420
         TabIndex        =   78
         Top             =   90
         Width           =   11550
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   105
         Picture         =   "FMain.frx":0978
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.Frame FrameEnter 
      Caption         =   "Enter"
      Height          =   6465
      Left            =   3630
      OLEDropMode     =   1  'Manual
      TabIndex        =   56
      Top             =   90
      Width           =   14685
      Begin HFPayables.VBCombo cboDepartment 
         Height          =   240
         Left            =   9000
         TabIndex        =   11
         Top             =   912
         Width           =   2388
         _ExtentX        =   4207
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.TextBox txtVendor 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   840
         ScrollBars      =   1  'Horizontal
         TabIndex        =   0
         Top             =   900
         Width           =   2955
      End
      Begin VB.TextBox txtCompany 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2955
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2145
         Width           =   2955
      End
      Begin VB.TextBox lblInvoice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   150
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   66
         Text            =   "Invoice"
         Top             =   1155
         Width           =   645
      End
      Begin VB.TextBox lblVendor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   165
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Text            =   "Vendor"
         Top             =   915
         Width           =   630
      End
      Begin TDBNumber6Ctl.TDBNumber numAmount 
         Height          =   225
         Left            =   4980
         TabIndex        =   4
         Top             =   900
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calculator      =   "FMain.frx":0F02
         Caption         =   "FMain.frx":0F22
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":0F87
         Keys            =   "FMain.frx":0FA5
         Spin            =   "FMain.frx":0FEF
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,##0.00;-###,###,##0.00;0.00;0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   -1
         ForeColor       =   -2147483640
         Format          =   "###,###,##0.00;-###,###,##0.00"
         HighlightText   =   -1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.99
         MinValue        =   -999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1952776197
         MinValueVT      =   1146880005
      End
      Begin MSComctlLib.Toolbar Toolbar 
         Height          =   570
         Left            =   60
         TabIndex        =   61
         Top             =   180
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   1005
         ButtonWidth     =   2064
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&New"
               Key             =   "new"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&PO"
               Key             =   "quickpay"
               Object.ToolTipText     =   "Add a PO to the invoice"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "A&dd"
               Key             =   "add"
               Object.ToolTipText     =   "Add a PO to the invoice"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Var &Rem"
               Key             =   "var"
               Object.ToolTipText     =   "Apply amount remaining as variance"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "P&ost"
               Key             =   "post"
               Object.ToolTipText     =   "Post the invoice to accounting"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Approve"
               Key             =   "approved"
               Object.ToolTipText     =   "Save the invoice into the approved file"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "P&ending"
               Key             =   "pending"
               Object.ToolTipText     =   "Save the invoice for approval"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Hold"
               Key             =   "hold"
               Object.ToolTipText     =   "Save the invoice into the hold file"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cancel"
               Key             =   "cancel"
               Object.ToolTipText     =   "Cancel the invoice"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Doc&uments"
               Key             =   "documents"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "E&mail"
               Key             =   "email"
            EndProperty
         EndProperty
         OLEDropMode     =   1
         Begin MSComctlLib.ImageList ToolbarImageList 
            Left            =   12900
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   14
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1017
                  Key             =   "new"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1E69
                  Key             =   "post"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2743
                  Key             =   "quickpay"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":301D
                  Key             =   "email"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":38F7
                  Key             =   "scan"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":41D1
                  Key             =   "add"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":4AAB
                  Key             =   "var"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":58FD
                  Key             =   "documents"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":61D7
                  Key             =   "documentsfull"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":6AB1
                  Key             =   "cancel"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":738B
                  Key             =   "hold"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":81DD
                  Key             =   "split"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":8AB7
                  Key             =   "pending"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":9909
                  Key             =   "approved"
               EndProperty
            EndProperty
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid gEnterTail 
         Height          =   1065
         Left            =   90
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   5310
         Width           =   9975
         _cx             =   1981957307
         _cy             =   1981941591
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   -2147483635
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   0
         GridLineWidth   =   0
         Rows            =   4
         Cols            =   8
         FixedRows       =   4
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   220
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FMain.frx":A1E3
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   1
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
      Begin VSFlex8Ctl.VSFlexGrid gEnterDist 
         Height          =   1875
         Left            =   60
         TabIndex        =   21
         Top             =   3300
         Width           =   12195
         _cx             =   1981961223
         _cy             =   1981943019
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
         BackColorAlternate=   -2147483624
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   -2147483635
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   36
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FMain.frx":A32E
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   3
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
      Begin TDBDate6Ctl.TDBDate dteReceived 
         Height          =   225
         Left            =   7020
         TabIndex        =   7
         Top             =   900
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FMain.frx":A88A
         Caption         =   "FMain.frx":A97F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":A9E4
         Keys            =   "FMain.frx":AA02
         Spin            =   "FMain.frx":AA60
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
         Value           =   3.66819414244724E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dteInvoice 
         Height          =   225
         Left            =   840
         TabIndex        =   2
         Top             =   1380
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FMain.frx":AA88
         Caption         =   "FMain.frx":AB7D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":ABE2
         Keys            =   "FMain.frx":AC00
         Spin            =   "FMain.frx":AC56
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
         Value           =   1.10537306944062E-317
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dtePayment 
         Height          =   225
         Left            =   7020
         TabIndex        =   8
         Top             =   1140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FMain.frx":AC7E
         Caption         =   "FMain.frx":AD73
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":ADD8
         Keys            =   "FMain.frx":ADF6
         Spin            =   "FMain.frx":AE54
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
         Value           =   2.69366467562102E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dteAccounting 
         Height          =   225
         Left            =   7020
         TabIndex        =   10
         Top             =   1620
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FMain.frx":AE7C
         Caption         =   "FMain.frx":AF71
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":AFD6
         Keys            =   "FMain.frx":AFF4
         Spin            =   "FMain.frx":B052
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
      Begin TDBNumber6Ctl.TDBNumber numTax 
         Height          =   225
         Left            =   4980
         TabIndex        =   5
         Top             =   1140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calculator      =   "FMain.frx":B07A
         Caption         =   "FMain.frx":B09A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":B0FF
         Keys            =   "FMain.frx":B11D
         Spin            =   "FMain.frx":B167
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,##0.00;-###,###,##0.00;0.00;0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   -1
         ForeColor       =   -2147483640
         Format          =   "###,###,##0.00;-###,###,##0.00"
         HighlightText   =   -1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.99
         MinValue        =   -999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1952776197
         MinValueVT      =   1146880005
      End
      Begin TDBNumber6Ctl.TDBNumber numDiscount 
         Height          =   225
         Left            =   4980
         TabIndex        =   6
         Top             =   1380
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calculator      =   "FMain.frx":B18F
         Caption         =   "FMain.frx":B1AF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":B214
         Keys            =   "FMain.frx":B232
         Spin            =   "FMain.frx":B27C
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,##0.00;-###,###,##0.00;0.00;0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   -1
         ForeColor       =   -2147483640
         Format          =   "###,###,##0.00;-###,###,##0.00"
         HighlightText   =   -1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.99
         MinValue        =   -999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1952776197
         MinValueVT      =   1146880005
      End
      Begin VB.TextBox txtInvoice 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   840
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1140
         Width           =   2295
      End
      Begin VB.TextBox txtJob 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   840
         TabIndex        =   3
         Top             =   1620
         Width           =   1935
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5070
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1920
         Width           =   2535
      End
      Begin TDBDate6Ctl.TDBDate dteDiscount 
         Height          =   225
         Left            =   7020
         TabIndex        =   9
         Top             =   1380
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calendar        =   "FMain.frx":B2A4
         Caption         =   "FMain.frx":B399
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":B3FE
         Keys            =   "FMain.frx":B41C
         Spin            =   "FMain.frx":B47A
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
         Value           =   9.4965543544696E-318
         CenturyMode     =   0
      End
      Begin VB.TextBox txtInvoiceCode1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5070
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   1635
      End
      Begin VB.TextBox txtInvoiceCode2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5070
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   2640
         Width           =   1635
      End
      Begin HFPayables.VBCombo cboApprover 
         Height          =   240
         Left            =   9000
         TabIndex        =   12
         Top             =   1176
         Width           =   2388
         _ExtentX        =   4207
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   9000
         MultiLine       =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Hold Reason"
         Top             =   1425
         Width           =   5295
      End
      Begin TDBNumber6Ctl.TDBNumber numWrapInsuranceAmount 
         Height          =   225
         Left            =   5070
         TabIndex        =   18
         Top             =   2160
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   397
         Calculator      =   "FMain.frx":B4A2
         Caption         =   "FMain.frx":B4C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FMain.frx":B527
         Keys            =   "FMain.frx":B545
         Spin            =   "FMain.frx":B58F
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,##0.00;-###,###,##0.00;0.00;0.00"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   -1
         ForeColor       =   -2147483640
         Format          =   "###,###,##0.00;-###,###,##0.00"
         HighlightText   =   -1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.99
         MinValue        =   -999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1952776197
         MinValueVT      =   1146880005
      End
      Begin VB.Label lblWrapInsuranceRate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6225
         TabIndex        =   80
         Top             =   2175
         Width           =   45
      End
      Begin VB.Label lblWrapInsuranceAmount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Misc Deduction"
         Height          =   195
         Left            =   3960
         TabIndex        =   79
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label lblHoldbackInvoice 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "This is a holdback invoice."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4170
         TabIndex        =   75
         Top             =   3090
         Width           =   2175
      End
      Begin VB.Label lblApprovalHistory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "history..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   11670
         TabIndex        =   77
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   195
         Index           =   0
         Left            =   8205
         TabIndex        =   76
         Top             =   1425
         Width           =   750
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   195
         Index           =   13
         Left            =   8115
         TabIndex        =   74
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Job"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   495
         TabIndex        =   67
         Top             =   1620
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Date"
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   38
         Top             =   1380
         Width           =   630
      End
      Begin VB.Image imgEditJob 
         Height          =   240
         Left            =   2790
         Picture         =   "FMain.frx":B5B7
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgEditVendor 
         Height          =   240
         Left            =   3900
         Picture         =   "FMain.frx":BB41
         Top             =   885
         Width           =   240
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approver"
         Height          =   195
         Index           =   12
         Left            =   8280
         TabIndex        =   73
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label lblJobDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "job not found"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3090
         TabIndex        =   72
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblInvoiceCode2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Misc Code 2"
         Height          =   195
         Left            =   4170
         TabIndex        =   71
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblInvoiceCode1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Misc Code 1"
         Height          =   195
         Left            =   4170
         TabIndex        =   70
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Variance exceeds your override limit. The invoice has an un-coded variance category. Exceeds committed unit values."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   8100
         TabIndex        =   63
         Top             =   1950
         Width           =   4035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   11
         Left            =   4230
         TabIndex        =   54
         Top             =   1920
         Width           =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   9.99999e5
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   39
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         Height          =   195
         Index           =   9
         Left            =   6345
         TabIndex        =   43
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Received"
         Height          =   195
         Index           =   7
         Left            =   6315
         TabIndex        =   42
         Top             =   900
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Index           =   6
         Left            =   6360
         TabIndex        =   44
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting"
         Height          =   195
         Index           =   10
         Left            =   6180
         TabIndex        =   45
         Top             =   1620
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax"
         Height          =   195
         Index           =   2
         Left            =   4665
         TabIndex        =   40
         Top             =   1140
         Width           =   270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Index           =   5
         Left            =   4320
         TabIndex        =   41
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label lblTerms 
         BackColor       =   &H80000005&
         Height          =   465
         Left            =   840
         TabIndex        =   16
         Top             =   2760
         UseMnemonic     =   0   'False
         Width           =   2955
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   60
         Top             =   2730
         Width           =   435
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   59
         Top             =   2145
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   58
         Top             =   1920
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   9.99999e5
         Y1              =   795
         Y2              =   795
      End
   End
   Begin VB.Frame FrameView 
      Caption         =   "View"
      Height          =   6015
      Left            =   8985
      TabIndex        =   57
      Top             =   4710
      Visible         =   0   'False
      Width           =   7635
      Begin HFPayables.Slider Slider 
         Height          =   45
         Left            =   120
         Top             =   3570
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   79
         Orientation     =   1
         Max             =   18105
      End
      Begin VB.Frame FrameSearch 
         BorderStyle     =   0  'None
         Caption         =   "4"
         Height          =   1815
         Left            =   90
         TabIndex        =   32
         Top             =   30
         Width           =   7125
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            Enabled         =   0   'False
            Height          =   315
            Left            =   5580
            TabIndex        =   34
            Top             =   1380
            Width           =   1155
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            Height          =   315
            Left            =   5580
            TabIndex        =   33
            Top             =   1020
            Width           =   1155
         End
         Begin TDBDate6Ctl.TDBDate dteSearchInvoiceDate 
            Height          =   225
            Left            =   4380
            TabIndex        =   28
            Top             =   720
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   397
            Calendar        =   "FMain.frx":C0CB
            Caption         =   "FMain.frx":C1C0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FMain.frx":C225
            Keys            =   "FMain.frx":C243
            Spin            =   "FMain.frx":C2AF
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
            ErrorBeep       =   0
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
            Value           =   38476
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate dteSearchReceivedDate 
            Height          =   225
            Left            =   4380
            TabIndex        =   29
            Top             =   960
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   397
            Calendar        =   "FMain.frx":C2D7
            Caption         =   "FMain.frx":C3CC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FMain.frx":C431
            Keys            =   "FMain.frx":C44F
            Spin            =   "FMain.frx":C4BB
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
            ErrorBeep       =   0
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
            Value           =   38476
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate dteSearchPaymentDate 
            Height          =   225
            Left            =   4380
            TabIndex        =   30
            Top             =   1200
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   397
            Calendar        =   "FMain.frx":C4E3
            Caption         =   "FMain.frx":C5D8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FMain.frx":C63D
            Keys            =   "FMain.frx":C65B
            Spin            =   "FMain.frx":C6C7
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
            ErrorBeep       =   0
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
            Value           =   38476
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate dteSearchAccountingDate 
            Height          =   225
            Left            =   4380
            TabIndex        =   31
            Top             =   1440
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   397
            Calendar        =   "FMain.frx":C6EF
            Caption         =   "FMain.frx":C7E4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FMain.frx":C849
            Keys            =   "FMain.frx":C867
            Spin            =   "FMain.frx":C8D3
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
            ErrorBeep       =   0
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
            Value           =   38476
            CenturyMode     =   0
         End
         Begin VB.TextBox txtSearchVendor 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1500
            TabIndex        =   23
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtSearchInvoice 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1500
            TabIndex        =   24
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtSearchJob 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1500
            TabIndex        =   25
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtSearchCommitment 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1500
            TabIndex        =   26
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txtSearchUser 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1500
            TabIndex        =   22
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtSearchStatus 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4380
            TabIndex        =   27
            Top             =   480
            Width           =   975
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "FMain.frx":C8FB
            Top             =   120
            Width           =   480
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000F&
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   2.30016e6
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search for invoices using the criteria below."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   660
            TabIndex        =   68
            Top             =   60
            Width           =   3675
         End
         Begin VB.Label lblSearchStatus 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
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
            Left            =   3825
            TabIndex        =   64
            Top             =   480
            Width           =   450
         End
         Begin VB.Label lblSearchUser 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User"
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
            Left            =   1065
            TabIndex        =   62
            Top             =   480
            Width           =   330
         End
         Begin VB.Label lblSearchVendor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor"
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
            Left            =   885
            TabIndex        =   46
            Top             =   720
            Width           =   510
         End
         Begin VB.Label lblSearchInvoice 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice"
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
            Left            =   870
            TabIndex        =   47
            Top             =   960
            Width           =   525
         End
         Begin VB.Label lblSearchJob 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job"
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
            Left            =   1140
            TabIndex        =   48
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label lblSearchCommitment 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Commitment"
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
            Left            =   540
            TabIndex        =   49
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   3660
            TabIndex        =   50
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Received"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   3570
            TabIndex        =   51
            Top             =   960
            Width           =   690
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   3660
            TabIndex        =   52
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accounting"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   3450
            TabIndex        =   53
            Top             =   1440
            Width           =   810
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid gViewInv 
         Height          =   1395
         Left            =   120
         TabIndex        =   35
         Top             =   2100
         Width           =   7095
         _cx             =   1981952227
         _cy             =   1981942173
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
         BackColorAlternate=   -2147483624
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   -2147483635
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   30
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FMain.frx":D1C5
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   3
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
      Begin VSFlex8Ctl.VSFlexGrid gViewDist 
         Height          =   1935
         Left            =   120
         TabIndex        =   36
         Top             =   3660
         Width           =   7095
         _cx             =   1981952227
         _cy             =   1981943125
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
         BackColorAlternate=   -2147483624
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   28
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FMain.frx":D6C2
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
         ExplorerBar     =   3
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
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
      Begin VB.Image imgIconWeb 
         Height          =   240
         Left            =   7320
         Picture         =   "FMain.frx":DB43
         Top             =   4920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgIconDesk 
         Height          =   240
         Left            =   7320
         Picture         =   "FMain.frx":E0CD
         Top             =   4620
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgFlagError 
         Height          =   240
         Left            =   7320
         Picture         =   "FMain.frx":E657
         Top             =   4140
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgFlagHead 
         Height          =   240
         Left            =   7320
         Picture         =   "FMain.frx":EBE1
         Top             =   4380
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl ControlPanel 
      Align           =   3  'Align Left
      CausesValidation=   0   'False
      Height          =   10275
      Left            =   0
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   390
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   18124
      BackColorEnd    =   -1
      BackColorStart  =   -1
      Begin VB.PictureBox picRes 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   1620
         Picture         =   "FMain.frx":F16B
         ScaleHeight     =   1260
         ScaleWidth      =   1200
         TabIndex        =   69
         Top             =   1320
         Visible         =   0   'False
         Width           =   1200
      End
      Begin vbalIml6.vbalImageList IconsSmall 
         Left            =   60
         Top             =   4860
         _ExtentX        =   953
         _ExtentY        =   953
         ColourDepth     =   8
         Size            =   19516
         Images          =   "FMain.frx":1406D
         Version         =   131072
         KeyCount        =   17
         Keys            =   "ÿENTERINVOICESÿÿÿÿPOSTPENDINGÿVIEWARCHIVEÿREPORTÿAPPOPTIONSÿAPPABOUTÿÿÿÿExcelÿÿBACKCHARGESÿPREVIEW"
      End
      Begin vbalIml6.vbalImageList IconsLarge 
         Left            =   600
         Top             =   4860
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   32
         IconSizeY       =   32
         ColourDepth     =   32
         Size            =   4412
         Images          =   "FMain.frx":18CC9
         Version         =   131072
         KeyCount        =   1
         Keys            =   ""
      End
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   68
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":19E25
               Key             =   "SaveAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1A6FF
               Key             =   "snapshots"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":310C1
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3199B
               Key             =   "AssemblyCosts"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":32275
               Key             =   "MassChange"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":32B4F
               Key             =   "FieldPOs"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":33429
               Key             =   "NewRFQ"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":33D03
               Key             =   "CreateJob"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":345DD
               Key             =   "takeoffsettings"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":34EB7
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":35791
               Key             =   "DecreaseDecimals"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3606B
               Key             =   "IncreaseDecimals"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":36945
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3721F
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":37AF9
               Key             =   "UpdatePrices"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":383D3
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":38CAD
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":39587
               Key             =   "LookupPublished"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":39E61
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3A73B
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3B015
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3B8EF
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3C1C9
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3CAA3
               Key             =   "SendPOs"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3D37D
               Key             =   "SendRFQs"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3DC57
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3E531
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3EE0B
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3F6E5
               Key             =   "TakeoffPlanSwift"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3FFBF
               Key             =   "TakeoffPipeline"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":40899
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":41173
               Key             =   "TakeoffItemChart"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":41A4D
               Key             =   "New"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":42327
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":42C01
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":434DB
               Key             =   "Import"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":43DB5
               Key             =   "Export"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4468F
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":44F69
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":45843
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4611D
               Key             =   "AddPricelist"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":469F7
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":472D1
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":47BAB
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":48485
               Key             =   "PricelistExport"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":48D5F
               Key             =   "PricelistImport"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":49639
               Key             =   "NewPricelist"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":49F13
               Key             =   "View"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4A7ED
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4B0C7
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4B9A1
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4C27B
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4CB55
               Key             =   ""
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4D42F
               Key             =   "OptionWiz"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4F7B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5008B
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":50965
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5123F
               Key             =   "Models"
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":51B19
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":523F3
               Key             =   "Generate"
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":52CCD
               Key             =   "MasterBuilder"
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":535A7
               Key             =   "Timberline"
            EndProperty
            BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":53E81
               Key             =   "EEEstimating"
            EndProperty
            BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5475B
               Key             =   "Quickbooks"
            EndProperty
            BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":55035
               Key             =   "Sage50"
            EndProperty
            BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5590F
               Key             =   "Approve"
            EndProperty
            BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":561E9
               Key             =   "Decline"
            EndProperty
            BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":56AC3
               Key             =   "BuildPro"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "<mnuEdit>"
      Begin VB.Menu mnuEditSub 
         Caption         =   "Variance Remaining"
         Index           =   0
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "PO Inquiry"
         Index           =   1
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "Delete"
         Index           =   3
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuControlPanel 
      Caption         =   "<mnuControlPanel>"
      Begin VB.Menu mnuControlPanelSub 
         Caption         =   "Manage"
         Index           =   0
      End
   End
   Begin VB.Menu mnuPost 
      Caption         =   "<mnuPost>"
      Begin VB.Menu mnuPostSub 
         Caption         =   "(no pending invoices)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuViewInv 
      Caption         =   "<mnuViewInv>"
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Open"
         Index           =   10
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Post"
         Index           =   20
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Re-number..."
         Index           =   30
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Create Lien Voucher..."
         Index           =   40
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Lien Release Received"
         Index           =   45
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Set Draw Number..."
         Index           =   46
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "-"
         Index           =   50
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Documents"
         Index           =   60
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "-"
         Index           =   70
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Status"
         Index           =   80
         Begin VB.Menu mnuStatusSub 
            Caption         =   "Hold"
            Index           =   20
         End
         Begin VB.Menu mnuStatusSub 
            Caption         =   "Pending"
            Index           =   30
         End
         Begin VB.Menu mnuStatusSub 
            Caption         =   "Approved"
            Index           =   40
         End
         Begin VB.Menu mnuStatusSub 
            Caption         =   "Posted"
            Index           =   50
         End
         Begin VB.Menu mnuStatusSub 
            Caption         =   "-"
            Index           =   60
         End
         Begin VB.Menu mnuStatusSub 
            Caption         =   "Lien Hold"
            Index           =   70
         End
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "-"
         Index           =   90
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Select All"
         Index           =   100
      End
      Begin VB.Menu mnuViewInvSub 
         Caption         =   "Delete"
         Index           =   110
      End
   End
   Begin VB.Menu mnuGrid 
      Caption         =   "<mnuGrid>"
      Begin VB.Menu mnuGridSub 
         Caption         =   "Sort Ascending"
         Index           =   0
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "Sort Descending"
         Index           =   1
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "Remove this column"
         Index           =   3
      End
      Begin VB.Menu mnuColumns 
         Caption         =   "Insert a column"
         Begin VB.Menu mnuColumnsSub 
            Caption         =   "(none available)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuGridSub1 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuGridSub1 
         Caption         =   "Print..."
         Index           =   5
      End
      Begin VB.Menu mnuGridSub1 
         Caption         =   "Save As..."
         Index           =   6
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FMain::"


Private ValidationDEBUG As String

Private Enum IconIndexes
    WALLETBATCHES
    ENTERINVOICES
    VIEWHELD
    VIEWPENDING
    VIEWAPPROVED
    POSTPENDING
    INVOICESEARCH
    Report
    appoptions
    APPABOUT
    VIEWREJECTED
    VIEWPRINTFILE
    POCOMPLETIONS
    EXCEL
    BACKCHARGES = 15
    JOURNAL
End Enum

Public MouseForm  As Form
Public MouseCtrl  As Control
Public MouseCol   As Long


'enter frame stuff
    Private mCodeMode As Boolean
    ' stupid flag -- see chkAdvance_Click() and txtInvoice_Change()
    Private mHBInvoiceSuffix As String
    Private mTotalHBAmount As Double
    Private mTotalHBTax As Double
    
    Private mInvoiceID     As Long
    Private mStatus        As String  'status of invoice when opened
    Private mIsHBInvoice   As Boolean 'the opened invoice either is or has a holdback invoice associated with it. No new hold backs allowed
    Private mDirty         As Boolean
    Private mLoading       As Boolean
    Private mReadOnly      As Boolean
    Private mVendorType    As String
    Private mWrapInsuranceRate As Double
    Private mAPInvoicesDefaultToBillable As Boolean
    Private mCancelEdit    As Boolean
    Private mImages        As Collection
    

    'vendor defaults
    Private mVDefPhase        As String
    Private mVDefCategory     As String
    Private mVDefDebitAccount As String
    Private mVDefTaxGroup     As String
    
    'vendor payment terms
    Private mDiscType    As String   'InAGivenNumberOfDays, OnADayOfTheMonth, DayOfMonthAfterEOM, NumberOfDaysAfterEOM
    Private mNetType     As String
    Private mDiscDays    As Long
    Private mNetDays     As Long
    Private mDiscPercent As Double
    

'view frame stuff
    Private mStopSearch    As Boolean
    'grid menu
    
    Private Const mcGRID_ASC = 0
    Private Const mcGRID_DESC = 1
    Private Const mcGRID_HIDE = 3
    Private Const mcGRID_PRINT = 5
    Private Const mcGRID_SAVEAS = 6
    'view invoices menu
    Private Const mcVIEWINV_OPEN = 10
    Private Const mcVIEWINV_POST = 20
    Private Const mcVIEWINV_RENUMBER = 30
    Private Const mcVIEWINV_CREATEVOUCHER = 40
    Private Const mcVIEWINV_LIENRELEASERECEIVED = 45
    
    Private Const mcVIEWINV_DOCUMENTS = 60
    
    Private Const mcVIEWINV_STATUS = 80
    
    Private Const mcVIEWINV_SELECTALL = 100
    Private Const mcVIEWINV_DELETE = 110

'menu constants
Private Const mcEDIT_VARREM = 0
Private Const mcEDIT_INQUIRY = 1
Private Const mcEDIT_DELETE = 3

Private mV99 As String

Private m_ADOConnection As ADODB.Connection
 



Private Sub POInquiry()
On Error GoTo eh
    Dim po As String
    Dim r As Long
    
    With gEnterDist
    If .Row > 0 And .Row < .Rows - 1 Then
        'sitting on a PO so use it
        po = .TextMatrix(.Row, .ColIndex("commitment"))
        If po <> "" Then Call HFApp.RunTask("POInquiry|" & po)
    Else
        'find first PO
        For r = 1 To .Rows - 2
            po = .TextMatrix(r, .ColIndex("commitment"))
            If po <> "" Then
                Call HFApp.RunTask("POInquiry|" & po)
                Exit Sub
            End If
        Next
    End If
    End With
Exit Sub
eh: Call ErrHandler(SRCFILE & "POInquiry")
End Sub






Private Sub cboApprover_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub cboDepartment_Click()
    Call GetApprover(, GetComboBoxListID(cboDepartment))
    Dirty = True
End Sub

Private Sub cboDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub dteAccounting_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub dteDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub dteInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub dtePayment_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub dteReceived_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub gEnterDist_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
On Error Resume Next
    With gEnterDist
        Call .Select(.Row, .Col, .RowSel, .Col)
    End With
End Sub

Private Sub gEnterDist_BeforeSort(ByVal Col As Long, Order As Integer)
    gEnterDist.RemoveItem gEnterDist.Rows - 1
End Sub
Private Sub gEnterDist_AfterSort(ByVal Col As Long, Order As Integer)
    gEnterDist.AddItem ""
End Sub

Private Sub gEnterDist_GotFocus()
    Call GridGotFocus(gEnterDist)
End Sub

Private Sub gEnterDist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With gEnterDist
        If Button <> vbRightButton Or .MouseRow < 0 Or .MouseRow > .Rows - 2 Or .MouseCol < 0 Then Exit Sub
        
        If .MouseRow = 0 Then
            If gEnterDist.MouseRow = 0 Then Call ShowColumnMenu(gEnterDist, , gEnterDist.TextMatrix(0, gEnterDist.MouseCol) <> "")
        Else
            If ReadOnly Then Exit Sub
            Call .Select(.MouseRow, .MouseCol)
            mnuEditSub(mcEDIT_INQUIRY).Enabled = .TextMatrix(.Row, .ColIndex("commitment")) <> ""
            PopupMenu mnuEdit
        End If
    
    End With
End Sub

Private Sub gViewInv_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    Dim r           As Long
    Dim i           As Long
        
    With gViewInv
        For r = Min(Row, .RowSel) To Max(Row, .RowSel)
            If .Cell(flexcpChecked, r, .ColIndex("RetainageInvoice")) = flexChecked Then
                If .TextMatrix(r, .ColIndex("InvoiceDate")) <> .TextMatrix(Row, .ColIndex("InvoiceDate")) Then
                    .TextMatrix(r, .ColIndex("InvoiceDate")) = .TextMatrix(Row, .ColIndex("InvoiceDate"))
                End If
                Call HFApp.SqlExec("Update Invoices set InvoiceDate = " & DbQuote(Date, .TextMatrix(r, .ColIndex("InvoiceDate"))) & " Where DivisionID =" & HFApp.DivisionID & " and InvoiceID = " & DbQuote(num, .ValueMatrix(r, .ColIndex("InvoiceID"))))
            End If
        Next
    End With
    
Exit Sub
End Sub

Private Sub gViewInv_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With gViewInv
    .ComboList = ""
    If .ColKey(Col) = "InvoiceDate" And .Cell(flexcpChecked, Row, .ColIndex("RetainageInvoice")) = flexChecked Then
        .ComboList = "|..."
    Else
        Cancel = True
    End If
End With
End Sub

Private Sub gViewInv_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With gViewInv
    If .ColKey(Col) = "InvoiceDate" Then
        Call DCalendar.Popup(gViewInv, .RowPos(.Row) + .RowHeight(.Row), .colPos(.Col))
        Call gViewInv_AfterEdit(Row, Col)
    End If
End With
End Sub

Private Sub imgEditJob_Click()
    Call HFApp.EditJob(txtJob.Text)
End Sub

Private Sub imgEditVendor_Click()
    If App.EditVendors Then
        Call HFApp.EditVendor(txtVendor.tag)
    Else
        MsgBox "Your security settings do not allow you to edit vendors." & vbCrLf & "Contact an administrator to change your permissions.", vbInformation, App.ProductName
    End If
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 And Shift <> 0 And ValidationDEBUG <> "" Then MsgBox ValidationDEBUG, vbInformation, App.ProductName
End Sub

Private Sub lblApprovalHistory_Click()
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select l.dstmp,l.ustmp,l.action,d.name department,l.assignedappr,l.approvercomments" & vbCrLf
    s = s & "from invoices i" & vbCrLf
    s = s & "join approvallog l on i.docid=l.docid" & vbCrLf
    s = s & "left outer join departments d on l.assigneddept=d.deptid" & vbCrLf
    s = s & "where i.invoiceid=" & DbQuote(num, mInvoiceID) & vbCrLf
    s = s & "order by 1" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    
    s = ""
    While Not rs.EOF
        s = s & Format("" & rs(0), App.Options(TimberlineDateFormat)) & vbTab & rs(1) & vbTab & rs(2) & vbTab & rs(3) & vbTab & rs(4) & vbTab & rs(5) & vbCrLf
        rs.MoveNext
    Wend
    
    Call FComments.Edit(s, Nothing, False, "History")
    
    
End Sub

Private Sub mnuControlPanelSub_Click(Index As Integer)
On Error GoTo ExitSub
    Dim s As String
    
    Select Case mnuControlPanel.tag
        Case "HELP":    s = PathAppend(HFApp.SystemFolder, "Help\Payables")
        Case "TOOLS":   s = PathAppend(HFApp.SystemFolder, "Payables\Tools")
        Case "REPORTS": s = PathAppend(HFApp.SystemFolder, "Payables\Reports")
    End Select
    
    If s <> "" Then
        Call CreatePath("", s)
        Call ShellFile(Me.hwnd, s)
    End If
ExitSub:
End Sub




Private Sub gEnterDist_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Cancel = mCancelEdit
End Sub

Private Sub gEnterDist_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next
    Dim Cancel As Boolean
    Select Case KeyCode
        Case vbKeyReturn
            Call gEnterDist.Select(Row, Col + 1)
            
        Case vbKeyF4
            If Shift = 0 Then
                If gEnterDist.ComboList <> "" Then
                    mLoading = True 'so that afteredit wont validate partial entry
                    Call gEnterDist_CellButtonClick(gEnterDist.Row, gEnterDist.Col)
                    mLoading = False
                    Call gEnterDist_AfterEdit(Row, Col)
                End If
            End If
            
    End Select
End Sub

Private Sub gViewInv_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo eh
    Dim s As String
    Dim r As Long
    With gViewInv
    Select Case True
        Case gViewInv.MouseRow = 0 And Button = vbRightButton
            Cancel = True
            Call ShowColumnMenu(gViewInv, , gViewInv.TextMatrix(0, gViewInv.MouseCol) <> "")
            
        Case IsBetween(gViewInv.MouseRow, 1, gViewInv.Rows - 1) And Button = vbRightButton
            If gViewInv.SelectedRows < 1 Then
                gViewInv.Row = gViewInv.MouseRow
            Else
                Cancel = True
            End If
            
            
            mnuViewInvSub(mcVIEWINV_LIENRELEASERECEIVED).Enabled = False
            For r = 0 To .SelectedRows - 1
                If .TextMatrix(.SelectedRow(r), .ColIndex("LienVoucher")) <> "" Then
                    mnuViewInvSub(mcVIEWINV_LIENRELEASERECEIVED).Enabled = True
                End If
            Next

            mnuViewInvSub(mcVIEWINV_POST).Enabled = App.PostInvoices
            mnuViewInvSub(mcVIEWINV_DELETE).Enabled = App.EditInvoices
            PopupMenu mnuViewInv, , , , mnuViewInvSub(mcVIEWINV_OPEN)
            
    End Select
    End With
        
Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewInv_BeforeMouseDown")
End Sub

Private Sub gViewInv_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo eh
    Dim s  As String
    Dim r  As Long
    Dim rs As Recordset
    
    If OldRow = NewRow And Not (OldRow = 1 And OldCol = 0) Then Exit Sub

    With gViewDist
    
        Screen.MousePointer = vbHourglass
    
        s = ""
        s = s & "select Commitment,CommitmentDesc,CommitmentItem,Job,Extra,ExtraDesc,JobDesc,Phase,PhaseDesc,Category,CategoryDesc,DebitAccount,DebitAccountDesc,TaxGroup,TaxGroupDesc,TaxRate,Description,EquipmentCostCode,EquipmentCostCodeDesc,Equipment,EquipmentDesc,InvoicedQuantity,InvoicedUnitPrice,PreTax,Tax,Retainage,JointPayee,Errors" & vbCrLf
        s = s & "  from invoiceitems" & vbCrLf
        s = s & " where DivisionID =" & HFApp.DivisionID & " and Invoiceid=" & DbQuote(num, gViewInv.ValueMatrix(NewRow, gViewInv.ColIndex("InvoiceID"))) & vbCrLf
        s = s & "order by itemid" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        
        .Rows = 1
        gViewDist.Rows = 1
        mStopSearch = False
        While Not (rs.EOF Or mStopSearch)
            .AddItem ""
            r = .Rows - 1
            
            .TextMatrix(r, .ColIndex("Flag")) = "" & rs("Errors") 'varcharmax col read it first
            If Trim(.TextMatrix(r, .ColIndex("Flag"))) <> "" Then .Cell(flexcpPicture, r, .ColIndex("Flag")) = imgFlagError.Picture
            
            .TextMatrix(r, .ColIndex("Commitment")) = "" & rs("Commitment")
            .TextMatrix(r, .ColIndex("CommitmentDesc")) = "" & rs("CommitmentDesc")
            .TextMatrix(r, .ColIndex("CommitmentItem")) = IIf(Val("" & rs("CommitmentItem")) = 0, "", "" & rs("CommitmentItem"))
            .TextMatrix(r, .ColIndex("Job")) = "" & rs("Job")
            .TextMatrix(r, .ColIndex("Extra")) = "" & rs("Extra")
            .TextMatrix(r, .ColIndex("ExtraDesc")) = "" & rs("ExtraDesc")
            .TextMatrix(r, .ColIndex("JobDesc")) = "" & rs("JobDesc")
            .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(r, .ColIndex("PhaseDesc")) = "" & rs("PhaseDesc")
            .TextMatrix(r, .ColIndex("Category")) = "" & rs("Category")
            .TextMatrix(r, .ColIndex("CategoryDesc")) = "" & rs("CategoryDesc")
            .TextMatrix(r, .ColIndex("DebitAccount")) = "" & rs("DebitAccount")
            .TextMatrix(r, .ColIndex("DebitAccountDesc")) = "" & rs("DebitAccountDesc")
            .TextMatrix(r, .ColIndex("TaxGroup")) = "" & rs("TaxGroup")
            .TextMatrix(r, .ColIndex("TaxGroupDesc")) = "" & rs("TaxGroupDesc")
            .TextMatrix(r, .ColIndex("TaxRate")) = "" & rs("TaxRate")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("JointPayee")) = "" & rs("JointPayee")
            
            .TextMatrix(r, .ColIndex("EQCostCode")) = "" & rs("EquipmentCostCode")
            .TextMatrix(r, .ColIndex("EQCostCodeDesc")) = "" & rs("EquipmentCostCodeDesc")
            .TextMatrix(r, .ColIndex("Equipment")) = "" & rs("Equipment")
            .TextMatrix(r, .ColIndex("EquipmentDesc")) = "" & rs("EquipmentDesc")
            
            .TextMatrix(r, .ColIndex("InvoicedQuantity")) = "" & rs("InvoicedQuantity")
            .TextMatrix(r, .ColIndex("InvoicedUnitPrice")) = "" & rs("InvoicedUnitPrice")

            .TextMatrix(r, .ColIndex("PreTax")) = "" & rs("PreTax")
            .TextMatrix(r, .ColIndex("Tax")) = "" & rs("Tax")
            .TextMatrix(r, .ColIndex("Retainage")) = "" & rs("Retainage")
            

            .Cell(flexcpForeColor, r, .ColIndex("PreTax")) = IIf(.ValueMatrix(r, .ColIndex("PreTax")) < -0.001, vbRed, vbWindowText)
            .Cell(flexcpForeColor, r, .ColIndex("Tax")) = IIf(.ValueMatrix(r, .ColIndex("Tax")) < -0.001, vbRed, vbWindowText)
            .Cell(flexcpForeColor, r, .ColIndex("Retainage")) = IIf(.ValueMatrix(r, .ColIndex("Retainage")) < -0.001, vbRed, vbWindowText)

            
            rs.MoveNext
        Wend

    End With

    Screen.MousePointer = vbDefault
    

    Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewInv_BeforeRowColChange")
End Sub

Private Sub gViewInv_DblClick()
    Call mnuViewInvSub_Click(mcVIEWINV_OPEN)
    Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewInv_BeforeRowColChange")
End Sub


Private Sub gViewDist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo eh
    If Button = vbRightButton And gViewDist.MouseRow = 0 Then Call ShowColumnMenu(gViewDist, , gViewDist.TextMatrix(0, gViewDist.MouseCol) <> "")
    Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewDist_MouseDown")
End Sub


Private Sub ShowPostMenu()
    Dim i  As Long
    Dim rs As Recordset
    Dim s As String
    
    Screen.MousePointer = vbHourglass
    mnuPostSub(0).Visible = True
    For i = mnuPostSub.UBound To 1 Step -1
        Unload mnuPostSub(i)
    Next
    
    i = 0
    
    
    s = ""
    s = s & "select isnull(i.ustmp,'unknown') Ustmp,isnull(i.ustmp,'unknown') + '''s invoices (' +  cast(count(*) as VARCHAR(10))+ ' approved)' " & vbCrLf
    s = s & "from invoices i" & vbCrLf
    s = s & "join tblvendors v on i.divisionid=v.divisionid and i.vendor=v.vendor_id and isnull(v.istbd,0)=0" & vbCrLf
    s = s & "where i.DivisionID=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & "and status in('Approved','SiteApproved') group by ustmp" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        i = i + 1
        Load mnuPostSub(i)
        mnuPostSub(i).Caption = rs(1)
        mnuPostSub(i).tag = rs(0)
        mnuPostSub(i).Visible = True
        mnuPostSub(i).Enabled = True
        mnuPostSub(0).Visible = False
        rs.MoveNext
    Wend
    Screen.MousePointer = vbNormal
    
    Call PopupMenu(mnuPost)
    Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewInv_BeforeRowColChange")
End Sub


Private Sub gViewInv_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:   Call mnuViewInvSub_Click(mcVIEWINV_OPEN)
        Case vbKeyDelete:   Call mnuViewInvSub_Click(mcVIEWINV_DELETE)
        Case vbKeyF2:       Call mnuViewInvSub_Click(mcVIEWINV_RENUMBER)
    End Select
End Sub


Private Sub gViewInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    With gViewInv
        If .ColIndex("Flag") = .MouseCol Or .ColIndex("Source") = .MouseCol Then
            .ToolTipText = Replace(.Cell(flexcpTextDisplay, .MouseRow, .MouseCol), vbCrLf, " - ")
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub gViewDist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    With gViewDist
        If .ColIndex("Flag") = .MouseCol Then
            .ToolTipText = Replace(.Cell(flexcpTextDisplay, .MouseRow, .MouseCol), vbCrLf, " - ")
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub lblSearchStatus_Click()
    If FPickList.Choose(HFApp.Databases(dbHomeFront), "Status", "SELECT distinct Status FROM Invoices where DivisionID =" & HFApp.DivisionID, txtSearchStatus.Text) Then
        txtSearchStatus.Text = FPickList.SelectedItem(1)
    End If
    Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewInv_BeforeRowColChange")
End Sub

Private Sub mnuEditSub_Click(Index As Integer)
    Select Case Index
        Case mcEDIT_INQUIRY:   Call POInquiry
        Case mcEDIT_VARREM:    Call VarianceRemaining
        Case mcEDIT_DELETE:    Call gEnterDist_KeyDown(vbKeyDelete, vbShiftMask)
    End Select
End Sub

Private Sub mnuGridSub1_Click(Index As Integer)
    Call mnuGridSub_Click(Index)
End Sub

Private Sub mnuPostSub_Click(Index As Integer)
    Call PostInvoices(DbQuote(Str, mnuPostSub(Index).tag))
Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewInv_BeforeRowColChange")
End Sub


Private Sub ShowColumnMenu(Grid As VSFlexGrid, Optional Sortable As Boolean = True, Optional Hideable As Boolean = True)
On Error GoTo eh
    Dim i As Long
    Dim j As Long
    
    'save this stuff for menu click
    Set MouseGrid = Grid
    MouseCol = Grid.MouseCol
    
    'set these
    mnuGridSub(mcGRID_ASC).Enabled = Sortable
    mnuGridSub(mcGRID_DESC).Enabled = Sortable
    mnuGridSub(mcGRID_HIDE).Enabled = MouseCol >= 0 And Hideable
    

    'load Grid column names
    mnuColumnsSub(0).Visible = True
    For i = mnuColumnsSub.UBound To 1 Step -1
        Unload mnuColumnsSub(i)
    Next
    For i = 0 To MouseGrid.Cols - 1
        If MouseGrid.ColHidden(i) And MouseGrid.TextMatrix(0, i) <> "" Then
            j = j + 1
            Load mnuColumnsSub(j)
            mnuColumnsSub(j).tag = MouseGrid.ColKey(i)
            mnuColumnsSub(j).Caption = MouseGrid.TextMatrix(0, i)
            mnuColumnsSub(j).Visible = True
            mnuColumnsSub(j).Enabled = True
        End If
    Next
    
    
    
    If j = 0 Then mnuColumnsSub(0).Caption = "(none available)"
    mnuColumnsSub(0).Visible = j = 0
    mnuColumnsSub(0).Enabled = False
    
    'show menu
    PopupMenu mnuGrid

    Exit Sub
eh: Call ErrHandler(SRCFILE & "ShowColumnMenu")
End Sub

Private Sub lblSearchUser_Click()
    If FPickList.Choose(HFApp.Databases(dbHomeFront), "UserID", "SELECT distinct UStmp UserID FROM Invoices where DivisionID =" & HFApp.DivisionID, txtSearchUser.Text) Then
        txtSearchUser.Text = FPickList.SelectedItem(1)
    End If
    Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewInv_BeforeRowColChange")
End Sub


Private Sub mnuStatusSub_Click(Index As Integer)
    Dim InvoiceIDs As String
    Dim s As String
    Dim Approver As String
    Dim Comment  As String
    Dim r As Long
    
    
    'get list of selected invoice ids
    With gViewInv
        s = ""
        For r = 0 To .SelectedRows - 1
            InvoiceIDs = InvoiceIDs & "," & .TextMatrix(.SelectedRow(r), .ColIndex("InvoiceID"))
        Next
    End With
    InvoiceIDs = Mid(InvoiceIDs, 2)
    
    
    'validate complete and correct
    If Not PrePostValidate(" and i.invoiceid in(" & InvoiceIDs & ")") Then Exit Sub
    
    
    
    'check permissions and limits
    s = ""
    If InvoicesPosted(InvoiceIDs) And Not App.EditPostedInvoices Then
        s = "Your security settings do not allow you to edit posted invoices." & vbCrLf & "Contact an administrator to change your permissions."
        MsgBox s, vbInformation, App.ProductName
        Exit Sub
    End If
    Select Case mnuStatusSub(Index).Caption
    
        Case "Posted"
            If Not InvoiceApproved(InvoiceIDs) Then s = "Invoice has not been approved for posting"
            If Not App.PostInvoices Then s = "Your security settings do not allow you to post invoices." & vbCrLf & "Contact an administrator to change your permissions."
        
        Case "Approved"
            If ExceedsLimitDB(InvoiceIDs) Then s = "Invoice exceeds your approval limit." & vbCrLf & "Contact an administrator to change your limits."
            If Not App.ApproveInvoices Then s = "Your security settings do not allow you to approve invoices." & vbCrLf & "Contact an administrator to change your permissions."
            
 '       Case "Hold"
 '           If Not App.EditInvoices Then s = "Your security settings do not allow you to edit invoices." & vbCrLf & "Contact an administrator to change your permissions."
            
        Case Else
            If Not (App.PostInvoices Or App.EditInvoices) Then
                s = "Your security settings do not allow you to edit invoices." & vbCrLf & "Contact an administrator to change your permissions."
            Else
                s = "Mark" & IIf(Parse(InvoiceIDs) = 1, " this invoice ", " these " & Parse(InvoiceIDs) & " invoices ") & mnuStatusSub(Index).Caption & "?"
                If MsgBox(s, vbQuestion + vbYesNo, App.ProductName) <> vbYes Then
                    Exit Sub
                Else
                    s = ""
                End If
            End If
            
    End Select
    If s <> "" Then
        MsgBox s, vbInformation, App.ProductName
        Exit Sub
    End If
    
    
    If InvoiceIDs <> "" Then
        s = ""
        s = s & "update invoices" & vbCrLf
        s = s & "set status = " & DbQuote(Str, mnuStatusSub(Index).Caption, , , 15) & vbCrLf
        'approved
        If Index = 3 Then
            s = s & "   ,Approver=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        End If
        s = s & "   ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "   ,DStmp=getdate()" & vbCrLf
        s = s & "   ,TStmp=getdate()" & vbCrLf
        s = s & "where DivisionID =" & HFApp.DivisionID & vbCrLf
        s = s & " and invoiceid in(" & InvoiceIDs & ")"
        If Not App.EditPostedInvoices Then
            s = s & " and status not in('Exported','Posted')"
        End If
        
        Call HFApp.SqlExec(s)
    
        With gViewInv
            For r = 0 To .SelectedRows - 1
                .TextMatrix(.SelectedRow(r), .ColIndex("Status")) = mnuStatusSub(Index).Caption
            Next
        End With
    
        Call WriteApprovalLog(InvoiceIDs, mnuStatusSub(Index).Caption)
        
    End If

End Sub


Private Sub mnuViewInvSub_Click(Index As Integer)
On Error GoTo eh
    Dim r As Long
    Dim s As String
    Dim newNumber As String
    Dim oldNumber As String
    
    Select Case Index
        
        Case mcVIEWINV_RENUMBER
            With gViewInv
            If IsBetween(.Row, 1, .Rows - 1) Then
                
                If .TextMatrix(.Row, .ColIndex("Status")) = "Posted" Then Exit Sub
            
                newNumber = InputBox(vbCrLf & vbCrLf & "Enter a new invoice number for " & .TextMatrix(.Row, .ColIndex("VendorName")) & " invoice " & .TextMatrix(.Row, .ColIndex("Invoice")), "Renumber Invoice", .TextMatrix(.Row, .ColIndex("Invoice")))
                If newNumber = "" Then Exit Sub
                oldNumber = .TextMatrix(.Row, .ColIndex("Invoice"))
                
                s = ""
                s = s & "update invoices set invoice=" & DbQuote(Str, newNumber) & " where invoiceid=" & DbQuote(num, .TextMatrix(.Row, .ColIndex("InvoiceID"))) & vbCrLf
                s = s & "update invoiceitems set invoice=" & DbQuote(Str, newNumber) & " where invoiceid=" & DbQuote(num, .TextMatrix(.Row, .ColIndex("InvoiceID"))) & vbCrLf
'                s = s & "Update dms_documents set objectid=" & DbQuote(Str, InvoiceObjectID(.TextMatrix(.Row, .ColIndex("VendorName")), newNumber)) & " where objectid=" & DbQuote(Str, InvoiceObjectID(.TextMatrix(.Row, .ColIndex("VendorName")), oldNumber))
                Call HFApp.SqlExec(s)
        
                .TextMatrix(.Row, .ColIndex("Invoice")) = newNumber
                
            End If
            End With
            
        Case mcVIEWINV_DOCUMENTS
            With gViewInv
                If IsBetween(.Row, 1, .Rows - 1) Then
                    s = InvoiceObjectID(.TextMatrix(.Row, .ColIndex("Vendor")), .TextMatrix(.Row, .ColIndex("Invoice")))
                    Call HFApp.RunTask("EditAttachments|" & s & "|Invoice " & .TextMatrix(.Row, .ColIndex("Invoice")))
                End If
            End With
        
    
    
        Case mcVIEWINV_SELECTALL
            With gViewInv
                On Error Resume Next
                Call .Select(1, 0, .Rows - 1, .Cols - 1)
            End With
            
        Case mcVIEWINV_OPEN
            With gViewInv
                If IsBetween(.Row, 1, .Rows - 1) Then
                    FrameEnter.Visible = True
                    FrameView.Visible = False
                    Call Form_Resize
                    
                    Call OpenInvoice(.TextMatrix(.Row, .ColIndex("Vendor")), .TextMatrix(.Row, .ColIndex("Invoice")), True)
                End If
            End With
            
            
            
        Case mcVIEWINV_CREATEVOUCHER
            With gViewInv
                s = ""
                For r = 0 To .SelectedRows - 1
                    s = s & "," & .TextMatrix(.SelectedRow(r), .ColIndex("InvoiceID"))
                Next
                s = Mid(s, 2)
                If CreateLienVouchers(s) Then Call FindInvoices("")
            End With
            
            
        Case mcVIEWINV_LIENRELEASERECEIVED
            With gViewInv
                s = ""
                For r = 0 To .SelectedRows - 1
                    If .TextMatrix(.SelectedRow(r), .ColIndex("LienVoucher")) <> "" Then
                        s = s & "," & .TextMatrix(.SelectedRow(r), .ColIndex("LienVoucher"))
                    End If
                Next
                s = Mid(s, 2)
                If LienReleaseReceived(s) Then Call FindInvoices("")
            End With
                
            
            
        Case mcVIEWINV_DELETE
            With gViewInv
                s = ""
                For r = 0 To .SelectedRows - 1
                    s = s & "," & .TextMatrix(.SelectedRow(r), .ColIndex("InvoiceID"))
                Next
            End With
            s = Mid(s, 2)
            If s <> "" Then
                 If Parse(s) = 1 Then
                    If vbOK = MsgBox("Are you sure you want to permanently delete this invoice?", vbQuestion + vbOKCancel, App.ProductName) Then
                        Call HFApp.SqlExec("delete from Invoices where DivisionID =" & HFApp.DivisionID & " and invoiceid in(" & s & ")")
                        Call HFApp.SqlExec("delete from InvoiceItems where DivisionID =" & HFApp.DivisionID & " and invoiceid in(" & s & ")")
                        
                        With gViewInv
                            For r = .Rows - 1 To 1 Step -1
                                If .IsSelected(r) Then .RemoveItem (r)
                            Next
                        End With
                    End If
                Else
                    If vbOK = MsgBox("Are you sure you want to permanently delete these " & Parse(s) & " invoices?", vbQuestion + vbOKCancel, App.ProductName) Then
                        Call HFApp.SqlExec("delete from Invoices where DivisionID =" & HFApp.DivisionID & " and invoiceid in(" & s & ")")
                        Call HFApp.SqlExec("delete from InvoiceItems where DivisionID =" & HFApp.DivisionID & " and invoiceid in(" & s & ")")
                        
                        With gViewInv
                            For r = .Rows - 1 To 1 Step -1
                                If .IsSelected(r) Then .RemoveItem (r)
                            Next
                        End With
                    End If
                End If
            End If
            
        Case mcVIEWINV_POST
            If App.PostInvoices Then
                With gViewInv
                    s = ""
                    For r = 0 To .SelectedRows - 1
                        s = s & "," & .TextMatrix(.SelectedRow(r), .ColIndex("InvoiceID"))
                    Next
                End With
                s = Mid(s, 2)
                If s <> "" Then Call PostInvoices(, s)
                Call FindInvoices("")
            End If
    End Select
    Exit Sub
eh: If InStr(1, err.Description, "duplicate", vbTextCompare) Then
        MsgBox "Unable to re-number this invoice. That number has already been entered.", vbExclamation, App.ProductName
    Else
        Call ErrHandler(SRCFILE & "gViewInv_Click")
    End If
End Sub



Private Sub lblVendor_Click()
    
    Dim c As Boolean
    Dim s As String
    
    SetCtrlFocus txtVendor
    If Not txtVendor.Enabled Then Exit Sub
    
    If HFApp.Options(AccountingSystem) = asTimberline Then
        s = "SELECT Vendor_id Vendor, Vendor_Name Name FROM tblVendors WHERE DivisionID =" & HFApp.DivisionID & " and isnull(istbd,0)=0 and isnull(inactive,0)=0 order by 1"
    Else
        s = "SELECT Vendor_Name Name, Vendor_id Vendor FROM tblVendors WHERE DivisionID =" & HFApp.DivisionID & " and isnull(istbd,0)=0 and isnull(inactive,0)=0 order by 1"
    End If
    If FPickList.Choose(HFApp.Databases(dbHomeFront), "Vendor", s, , , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Vendor", "")) Then
        s = FPickList.SelectedItem("Vendor")
        txtVendor.Text = s
        On Error Resume Next
        Call txtVendor_Validate(c)
        txtInvoice.SetFocus
    End If
End Sub

Private Sub numAmount_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub numDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub numTax_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub txtComments_Change()
    Dirty = True
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub txtInvoice_GotFocus()
    SelectAll txtInvoice
End Sub

Private Sub txtInvoiceCode1_Change()
On Error Resume Next
    Dirty = True
End Sub
Private Sub txtInvoiceCode2_Change()
On Error Resume Next
    Dirty = True
End Sub

Private Sub txtInvoiceCode2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub

Private Sub txtSearchStatus_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblSearchStatus_Click
    End Select
End Sub

Private Sub txtSearchUser_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblSearchUser_Click
    End Select
End Sub


Private Sub lblSearchVendor_Click()
On Error Resume Next
    Dim s As String
    If HFApp.Options(AccountingSystem) = asTimberline Then
        s = "SELECT Vendor_id Vendor, Vendor_Name Name FROM tblVendors WHERE isnull(istbd,0)=0 and DivisionID =" & HFApp.DivisionID & " order by vendor_name"
    Else
        s = "SELECT Vendor_Name Name, Vendor_id Vendor FROM tblVendors WHERE isnull(istbd,0)=0 and DivisionID =" & HFApp.DivisionID & " order by vendor_name"
    End If
    If FPickList.Choose(HFApp.Databases(dbHomeFront), "Vendor", s) Then
        txtSearchVendor.Text = FPickList.SelectedItem("Vendor")
    End If
End Sub
Private Sub txtSearchVendor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblSearchVendor_Click
    End Select
End Sub

Private Sub lblSearchJob_Click()
On Error Resume Next
    If FPickList.Choose(HFApp.Databases(AccountingDB), App.Options(Caption_Job), SelectJob, txtSearchJob.Text) Then
        txtSearchJob.Text = FPickList.SelectedItem(1)
    End If
End Sub
Private Sub txtSearchJob_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblSearchJob_Click
    End Select
End Sub


Private Sub lblSearchInvoice_Click()
On Error Resume Next
    If FPickList.Choose(HFApp.Databases(dbHomeFront), "Invoice", "SELECT Invoice,InvoiceDate,Description,Status FROM Invoices WHERE Vendor = " & DbQuote(Str, txtSearchVendor.Text), txtSearchInvoice.Text) Then
        txtSearchInvoice.Text = FPickList.SelectedItem(1)
    End If
End Sub
Private Sub txtSearchInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblSearchInvoice_Click
    End Select
End Sub

Private Sub lblSearchCommitment_Click()
On Error Resume Next
    If FPickList.Choose(HFApp.Databases(AccountingDB), "Commitment", SelectPO, txtSearchCommitment.Text) Then
        txtSearchCommitment.Text = FPickList.SelectedItem(1)
    End If
End Sub
Private Sub txtSearchCommitment_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblSearchCommitment_Click
    End Select
End Sub





Private Sub mnuColumnsSub_Click(Index As Integer)
On Error Resume Next
    MouseGrid.ColHidden(MouseGrid.ColIndex(mnuColumnsSub(Index).tag)) = False
    MouseGrid.ColPosition(MouseGrid.ColIndex(mnuColumnsSub(Index).tag)) = IIf(MouseCol < 0, MouseGrid.Cols - 1, MouseCol)
End Sub
Private Sub mnuGridSub_Click(Index As Integer)
On Error Resume Next
    Dim i As Long
    Dim s As String
    Select Case Index
        Case mcGRID_ASC
            MouseGrid.Col = MouseCol
            MouseGrid.Sort = flexSortGenericAscending
            
        Case mcGRID_DESC
            MouseGrid.Col = MouseCol
            MouseGrid.Sort = flexSortGenericDescending
            
        Case mcGRID_HIDE
            For i = 0 To MouseGrid.Cols - 1
                If MouseGrid.ColHidden(i) = False Then
                    MouseGrid.ColHidden(MouseCol) = True
                    Exit Sub
                End If
            Next
            
        Case mcGRID_SAVEAS
            If VBGetSaveFileName(s, , , "Excel (*.xls)|*.xls|Text (*.txt)|*.txt|Comma Separated (*.csv)|*.csv", , , , "txt", FMain.hwnd) Then
                Select Case UCase(FileExt(s))
                    Case "XLS":  Call MouseGrid.SaveGrid(s, flexFileExcel, True)
                    Case "CSV":  Call MouseGrid.SaveGrid(s, flexFileCommaText, True)
                    Case Else:   Call MouseGrid.SaveGrid(s, flexFileTabText, True)
                End Select
            End If
            
        Case mcGRID_PRINT
            MouseGrid.TopRow = MouseGrid.FixedRows
            MouseGrid.LeftCol = MouseGrid.FixedCols
            Call MouseGrid.PrintGrid("", True, , 720, 720)

        
    End Select
End Sub

Private Property Get Dirty() As Boolean
    Dirty = mDirty
End Property

Private Property Let Dirty(RHS As Boolean)
    Dim s As String
    Dim InvOpen As Boolean
    Dim InvValid As Boolean
    Dim DistValid As Boolean
    Dim InvCompl As Boolean
    Dim POsCompl As Boolean
    Dim OvrLimit As Boolean
    Dim StatusLocked As Boolean
    
    If mReadOnly Then RHS = False
    mDirty = RHS
    

    'set warning label
    s = ""
    If ExceedsLimit Then s = s & "Invoice exceeds your approval limit." & vbCrLf
    If HasV99 Then s = s & "The invoice has an un-coded variance." & vbCrLf
    If ExceedsCommittedUnitValues Then s = s & "Exceeds committed unit values." & vbCrLf
    If InvoiceRemaining <> 0 Then s = s & Format(InvoiceRemaining, "$#,##0.00") & " remaining." & vbCrLf
    lblWarning.Caption = s
    
    
    'enable toolbar buttons
    InvOpen = txtInvoice.Text <> ""
    InvValid = ValidateInvoice("test", App.Options.Value(SaveIncompleteInvoices))
    InvCompl = gEnterDist.Rows > 2 And ValidateInvoice("test", False) 'inv balances or no rows entered and all coding correct
    StatusLocked = IsIn(mStatus, "Exported", "Posted") And Not App.EditPostedInvoices
    POsCompl = ArePOsCompleted() Or Not IsPOInvoice
    If Not POsCompl Then ValidationDEBUG = " " & vbBullet & " POs are not approved for payment" & vbCrLf & ValidationDEBUG
    'ignore if warn or off
    If Not POsCompl And App.Options(SchedulingUsage) < RestrictIfIncomplete Then POsCompl = True
    
    
    OvrLimit = ExceedsLimit()
    If OvrLimit Then ValidationDEBUG = " " & vbBullet & " Exceeds your limit" & vbCrLf & ValidationDEBUG
    
    With Toolbar
        
        .Buttons("hold").Enabled = InvOpen And InvValid And Not StatusLocked
        .Buttons("pending").Enabled = InvOpen And InvValid And Not StatusLocked
        .Buttons("approved").Enabled = POsCompl And InvCompl And InvOpen And InvValid And App.ApproveInvoices And Not OvrLimit And Not StatusLocked
        
        
        If IsPOInvoice Then
            'this is the dumbest feature. With autoapprovepo turned on, you cant put a po payment on pending
            If App.Options(AutoApprovePOInvoices) Then
                .Buttons("pending").Enabled = False
            End If
            If App.Options(SchedulingUsage) = HoldUntilCompleted And Not POsCompl Then
                .Buttons("pending").Enabled = False
                .Buttons("approved").Enabled = False
                .Buttons("post").Enabled = False
            End If
        End If
        
        .Buttons("post").Enabled = .Buttons("approved").Enabled And App.PostInvoices And Not StatusLocked
        
        .Buttons("add").Enabled = InvOpen And Not mReadOnly
        .Buttons("var").Enabled = InvOpen And Not mReadOnly And InvoiceRemaining <> 0 And App.EditInvoices
        .Buttons("quickpay").Enabled = True
        
        .Buttons("documents").Enabled = InvOpen
        .Buttons("email").Enabled = InvOpen
    End With
    
    
    
End Property

Private Function ExceedsCommittedUnitValues() As Boolean
    ExceedsCommittedUnitValues = False
'    Dim r As Long
'    With gEnterDist
'        For r = 1 To .Rows - 2
'            If .TextMatrix(r, .ColIndex("Commitment")) <> "" And _
'                (.ValueMatrix(r, .ColIndex("InvoicedQuantity")) > .ValueMatrix(r, .ColIndex("CommittedQuantity")) Or _
'                 .ValueMatrix(r, .ColIndex("InvoicedUnitPrice")) > .ValueMatrix(r, .ColIndex("CommittedUnitPrice"))) Then
'                ExceedsCommittedUnitValues = True
'                Exit Function
'            End If
'        Next
'    End With
'    ExceedsCommittedUnitValues = False
End Function

Private Function HasV99() As Boolean
    Dim r As Long
    With gEnterDist
        For r = 1 To .Rows - 2
            If mV99 = .TextMatrix(r, .ColIndex("Category")) And Not .RowHidden(r) Then
                HasV99 = True
                Exit Function
            End If
        Next
    End With
    HasV99 = False
End Function

Private Function IsPOInvoice() As Boolean
'   return true if invoice has nothing but po'd lines
    Dim r As Long
    
    Dim p As Boolean ' has po line
    Dim n As Boolean ' has non-po line
    
    With gEnterDist
        p = False
        n = False
        For r = 1 To .Rows - 2
            If Not .RowHidden(r) Then
                If Trim(.TextMatrix(r, .ColIndex("Commitment"))) <> "" Then
                    p = True
                Else
                    n = True
                End If
            End If
        Next
    End With
    
    IsPOInvoice = p And Not n
    
End Function

Private Sub dteAccounting_Change()
On Error Resume Next
    Dirty = True
End Sub


Private Sub dteDiscount_Change()
On Error Resume Next
    Dirty = True
End Sub

Private Sub dteInvoice_Click()
    Call dteInvoice_Change
End Sub



Private Sub dtePayment_Change()
    Dirty = True
End Sub

Private Sub dteReceived_Click()
    Call dteReceived_Change
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Dirty Then
        Select Case MsgBox("This invoice has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbCancel: Cancel = True
            Case vbNo:     Dirty = False
            Case vbYes:    Cancel = Not SaveInvoice()
        End Select
    End If
End Sub



Private Sub Form_Load()
On Error GoTo eh



    Dim i As Long
    Dim s As String
    Dim X As String
    Dim f As New ClsFileInfo
    Dim citem As cExplorerBarItem
    Dim c As Long
    
    lblJobDesc = ""
    
    
    mnuEdit.Visible = False
    mnuControlPanel.Visible = False
    mnuPost.Visible = False
    mnuViewInv.Visible = False
    mnuGrid.Visible = False
    
    Call SetToolbarIcons(Toolbar, ToolbarImageList)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gEnterDist)
    Call IniGetGrid(Me, gViewInv)
    Call IniGetGrid(Me, gViewDist)
    
    Call ReadSettings
    
    Toolbar.Visible = True
    Slider.Visible = True
    
    gViewInv.Editable = flexEDKbdMouse
    Call GetFormatMasks
    
    Call LoadControlPanel
    
    Dim b As Boolean
    b = Not HFApp.Options.ValueByName("IsProductionDatabase")
    picWarningBanner.Visible = b
    If Not HFApp.Options.ValueByName("IsProductionVerified") Then
        lblWarningBanner.Caption = "Cannot determine if this is your live database as the license server was not reachable. Try restarting the application."
        picWarningBanner.BackColor = &H80C0FF
    End If
    
    
    With gViewInv
        .Cell(flexcpPicture, 0, .ColIndex("Flag")) = Me.imgFlagHead
        .Cell(flexcpPictureAlignment, 0, .ColIndex("Flag")) = flexPicAlignCenterCenter
    End With
    With gViewDist
        .Cell(flexcpPicture, 0, .ColIndex("Flag")) = Me.imgFlagHead
        .Cell(flexcpPictureAlignment, 0, .ColIndex("Flag")) = flexPicAlignCenterCenter
    End With
    
    dteAccounting = ""
    Call SetWindowTitle
    Call OpenInvoice("", "")
    
    
    Call LoadDepartments
    Call DoAccountingSystemConfig
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "Form_Load")
End Sub

Private Sub LoadControlPanel()
On Error GoTo eh
    Dim s As String
    Dim X As String
    Dim i As Long
    Dim f As New ClsFileInfo
    Dim citem As cExplorerBarItem
    Dim c As Long
    
    With ControlPanel
    
        .Redraw = False
        .ImageList = IconsSmall.hIml
        .BarTitleImageList = IconsLarge.hIml
        .Bars.Clear
        
        With .Bars.Add(, "TASKS")
            .Title = "Payables Tasks"
            .IsSpecial = True
            .IconIndex = 0
            If App.EditInvoices Then
                .items.Add(, "ENTERINVOICES", "Enter Invoices", ENTERINVOICES).ToolTipText = "Enter and edit payable invoices"
                
                If App.Options(PayablesDisablePayPOsScreen) Then
                    Toolbar.Buttons("add").Visible = False
                    Toolbar.Buttons("quickpay").Visible = False
                Else
                    .items.Add(, "GENERATEINVOICES", "Pay Purchase Orders", ENTERINVOICES).ToolTipText = ""
                End If
            End If
                        
            .items.Add(, "VIEWPENDING", "View Pending Invoices", VIEWPENDING).ToolTipText = ""
            .items.Add(, "VIEWLIENHOLD", "View Lien Hold Invoices", VIEWPENDING).ToolTipText = ""
            .items.Add(, "VIEWLIENVOUCHER", "View Lien Vouchers", VIEWPENDING).ToolTipText = ""
            .items.Add(, "VIEWAPPROVED", "View Approved Invoices", VIEWAPPROVED).ToolTipText = ""
            .items.Add(, "VIEWHELD", "View Held Invoices", VIEWHELD).ToolTipText = ""
            .items.Add(, "VIEWHOLDBACK", "View Holdback Invoices", VIEWAPPROVED).ToolTipText = "View unposted holdback invoices"
            .items.Add(, "INVOICESEARCH", "Search the Archives", INVOICESEARCH).ToolTipText = "Find invoices that match your criteria"
            If App.EditInvoices Then
                If App.Options(EnableBackChargePOs) And HFApp.Options(AccountingSystem) = asTimberline Then
                    .items.Add(, "BACKCHARGES", "Back Chargeable PO's", BACKCHARGES).ToolTipText = "Create credit memos for back chargable PO's"
                End If
            End If
            If App.PostInvoices Then
                .items.Add(, "POSTPENDING", "Post Invoices", POSTPENDING).ToolTipText = "Post pending invoices"
                If HFApp.Options(AccountingSystem) = asIntacct Then
                    .items.Add(, "VIEWREJECTED", "View Rejected Invoices", VIEWREJECTED).ToolTipText = "View invoices that Intacct has rejected"
                End If
                If HFApp.Options(AccountingSystem) = asTimberline Then
                    .items.Add(, "VIEWREJECTED", "View Rejected Invoices", VIEWREJECTED).ToolTipText = "View invoices that Sage300 has rejected"
                    .items.Add(, "VIEWPRINTFILE", "View Sage300 Print File", VIEWPRINTFILE).ToolTipText = "View invoice import report"
                End If
            End If
            
            If App.EditInvoices Then
                .items.Add(, "POCOMPLETIONS", App.Options(Caption_Commitment) & " Completions", POCOMPLETIONS).ToolTipText = "View and update purchase orders"
                .items.Add(, "IMPORTINVOICE", "Import invoice...", EXCEL).ToolTipText = ""
                .items.Add(, "JOURNAL", "Journal Entries", JOURNAL).ToolTipText = ""
            End If
            
            
            
            .items.Add(, "SWITCHDIV", "Change Division", APPABOUT).ToolTipText = ""
            
        End With
                
                
        'load wallet UIs
        s = ""
        s = s & "select a.WalletBankAccount" & vbCrLf
        s = s & "from divisions d" & vbCrLf
        s = s & "join datasources a on d.bookofaccount=a.bookofaccount" & vbCrLf
        s = s & "where d.divisionid=" & HFApp.DivisionID & vbCrLf
        On Error Resume Next
        X = HFApp.SqlExec(s, dbHomeFront)(0)
        On Error GoTo eh
        If X <> "" Then
        With .Bars.Add(, "WALLET")
            .State = eBarCollapsed
            .Title = "Hyphen Wallet"
            i = 0
            
            'if sage300 and micr check writing is enabled
            If HFApp.Options(AccountingSystem) = asTimberline Then
                On Error Resume Next
                s = "" 'careful with this step. don't want it to error out if Sage is not available
                s = "" & HFApp.SqlExec("select using_micr_encoding from apm_master__ap_controls", dbAccounting)(0)
                On Error GoTo eh
                If s = "True" Or InIde() Then
                    .items.Add(, "SELECTINVOICESTOPAY", "Select Invoices To Pay", VIEWPRINTFILE).ToolTipText = ""
                End If
            End If
            
            .items.Add(, "WALLETBATCHES", "Post Payments to Wallet", WALLETBATCHES).ToolTipText = ""
            .items.Add(, "TOOL:-1", "Read AP Payments", POSTPENDING).tag = SyncCmd("ReadAP")
            .items.Add(, "WALLETBATCHSTATUS", "Posting Log", JOURNAL).ToolTipText = "Review status of batches sent to Wallet"
            .items.Add(, "EXPORTWALLETPAYEES", "Export Payees", EXCEL).ToolTipText = "Produce payee export file for Wallet support team"
        End With
        End If


        With .Bars.Add(, "REPORTS")
            .State = eBarCollapsed
            .Title = "Reports"
            i = 0
            
            s = Dir(PathAppend(HFApp.SystemFolder, "Payables\Reports\*.rpt"), , True)
            While s <> ""
                i = i + 1
                .items.Add(, "REPORT:" & i, FileName(s), Report).tag = PathAppend(HFApp.SystemFolder, "Payables\Reports", s)
                s = Dir(, , True)
            Wend
        End With
        
        With .Bars.Add(, "TOOLS")
            .State = eBarCollapsed
            .Title = "Tools"
            
            If App.Admin Then
                .items.Add(, "USERSECURITY", "User Permissions...", appoptions).ToolTipText = ""
                .items.Add(, "APPOPTIONS", "Application Options...", appoptions).ToolTipText = ""
                .items.Add(, "HFOPTIONS", "System Settings...", appoptions).ToolTipText = ""
            End If
            If App.EditVendors Then
                .items.Add(, "HFVENDORS", "Vendor Setup...", appoptions).ToolTipText = ""
            End If
            If App.Admin Then
                .items.Add(, "TOOL:0", "Read Vendors", POSTPENDING).tag = SyncCmd("ReadAllVendors")
                .items.Add(, "TOOL:1", "Read Accounts", POSTPENDING).tag = SyncCmd("ReadAllGLAccounts")
                If IsIn(Val(HFApp.Options(AccountingSystem)), asTimberline, asQuickBooks) Then
                    .items.Add(, "TOOL:2", "Read Cost Codes", POSTPENDING).tag = SyncCmd("ReadAllStandardCostCodes")
                End If
                .items.Add(, "TOOL:3", "Read Tax Settings", POSTPENDING).tag = SyncCmd("ReadTaxGroups")
            End If
            
            i = 4
            s = Dir(PathAppend(HFApp.SystemFolder, "Payables\Tools\*.*"), , True)
            While s <> ""
                i = i + 1
                f.FullPathName = PathAppend(HFApp.SystemFolder, "Payables\Tools", s)
                .items.Add(, "TOOL:" & i, FileName(s), FileExtensionIndex(f.hSmlIList, f.hSmlIcon)).tag = PathAppend(HFApp.SystemFolder, "Payables\Tools", s)
                s = Dir(, , True)
            Wend
        End With
        
        With .Bars.Add(, "HELP")
            .State = eBarCollapsed
            .Title = "Support Center"
            i = 0
            s = Dir(PathAppend(HFApp.SystemFolder, "Help\Payables\*.*"), , True)
            While s <> ""
                i = i + 1
                f.FullPathName = PathAppend(HFApp.SystemFolder, "Help\Payables", s)
                .items.Add(, "HELP:" & i, FileName(s), FileExtensionIndex(f.hSmlIList, f.hSmlIcon)).tag = PathAppend(HFApp.SystemFolder, "Help\Payables", s)
                s = Dir(, , True)
            Wend
            .items.Add , "APPABOUT", "About...", APPABOUT
        End With
        

        .Redraw = True
    End With

Exit Sub
eh: Call ErrHandler(SRCFILE & "LoadControlPanel")
End Sub

Private Sub cmdSearch_Click()
On Error GoTo eh

    Dim i As String
    Dim d As String
    
    
    
    'build where clause
    'i = "DivisionID =" & HFApp.DivisionID
    i = ""
    If txtSearchUser <> "" Then i = i & "   and isnull(i.UStmp,'') LIKE " & DbQuote(Str, "%" & txtSearchUser & "%") & vbCrLf
    If txtSearchVendor <> "" Then i = i & "   and (i.vendor LIKE " & DbQuote(Str, "%" & txtSearchVendor & "%") & " OR i.vendorname LIKE " & DbQuote(Str, "%" & txtSearchVendor & "%") & ")" & vbCrLf
    If txtSearchInvoice <> "" Then i = i & "   and i.Invoice LIKE " & DbQuote(Str, "%" & txtSearchInvoice & "%") & vbCrLf
    If Not dteSearchInvoiceDate.ValueIsNull Then i = i & "   and i.InvoiceDate = " & DbQuote(Date, dteSearchInvoiceDate) & vbCrLf
    If Not dteSearchReceivedDate.ValueIsNull Then i = i & "   and i.ReceivedDate = " & DbQuote(Date, dteSearchReceivedDate) & vbCrLf
    If Not dteSearchPaymentDate.ValueIsNull Then i = i & "   and i.PaymentDate = " & DbQuote(Date, dteSearchPaymentDate) & vbCrLf
    If Not dteSearchAccountingDate.ValueIsNull Then i = i & "   and i.AccountingDate = " & DbQuote(Date, dteSearchAccountingDate) & vbCrLf
    If txtSearchStatus <> "" Then i = i & "   and i.Status LIKE " & DbQuote(Str, "%" & txtSearchStatus & "%") & vbCrLf
                                                   
    If txtSearchJob <> "" Then d = d & "                 and d.Job LIKE " & DbQuote(Str, "%" & txtSearchJob & "%") & vbCrLf
    If txtSearchCommitment <> "" Then d = d & "                 and d.Commitment LIKE " & DbQuote(Str, "%" & txtSearchCommitment & "%") & vbCrLf
    
    If d <> "" Then i = i & "   and exists(select * " & vbCrLf & _
                            "                from invoiceitems d" & vbCrLf & _
                            "               where i.invoiceid = d.invoiceid" & vbCrLf & _
                            d & ")" & vbCrLf

    If i = "" Then i = i & "   and isnull(i.UStmp,'') LIKE " & DbQuote(Str, "%") & vbCrLf
    Call FindInvoices(Mid(i, 8))
    Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdSearch_Click")
End Sub

Private Sub cmdStop_Click()
    mStopSearch = True
End Sub

Private Sub FindInvoices(WhereClause As String)
    Dim s  As String
    Dim r  As Long
    Dim rs As Recordset
    
Static LastClause As String
If WhereClause = "" Then
    WhereClause = LastClause
Else
    LastClause = WhereClause
End If
    
    
    Screen.MousePointer = vbHourglass

    If Trim(WhereClause) = "" Then WhereClause = "1=2"
    s = ""
    s = s & "select d.name DeptName, InvoiceID, Vendor, VendorType, VendorName, Invoice, Job, JobDesc, Status, Pretax, Tax, Discount, DiscountDate, ReceivedDate, InvoiceDate, PaymentDate, AccountingDate, Description, UStmp, DStmp, Errors, Source, PostingDate, InvoiceCode1, InvoiceCode2, Approver, ApproverComments, DateApproved, Advance, BatchNumber, tstmp,  VendorAddr1, VendorAddr2, VendorCity, VendorProv, VendorPostal, isnull(RetainageInvoice,0) RetainageInvoice, Comments,i.LienVoucher" & vbCrLf
    s = s & "from invoices i" & vbCrLf
    s = s & "join tblvendors v on i.divisionid=v.divisionid and i.vendor=v.vendor_id" & vbCrLf
    s = s & "left outer join departments d on i.deptid=d.deptid" & vbCrLf
    s = s & "where i.DivisionID =" & HFApp.DivisionID & " and " & WhereClause
    Set rs = HFApp.SqlExec(s)
    
    With gViewInv
        .Rows = 1
        gViewDist.Rows = 1
        
        mStopSearch = False
        While Not (rs.EOF Or mStopSearch)
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("Flag")) = "" & rs("Errors") 'varcharmax col read first
            If Trim(.TextMatrix(r, .ColIndex("Flag"))) <> "" Then .Cell(flexcpPicture, r, .ColIndex("Flag")) = imgFlagError.Picture

            .TextMatrix(r, .ColIndex("InvoiceID")) = "" & rs("InvoiceID")
            .TextMatrix(r, .ColIndex("Status")) = "" & rs("Status")
            .TextMatrix(r, .ColIndex("Vendor")) = "" & rs("Vendor")
            .TextMatrix(r, .ColIndex("Invoice")) = "" & rs("Invoice")
            .TextMatrix(r, .ColIndex("Department")) = "" & rs("DeptName")
            .TextMatrix(r, .ColIndex("VendorName")) = "" & rs("VendorName")
            .TextMatrix(r, .ColIndex("InvoiceDate")) = "" & rs("InvoiceDate")
            .TextMatrix(r, .ColIndex("Job")) = "" & rs("Job")
            .TextMatrix(r, .ColIndex("JobDesc")) = "" & rs("JobDesc")
            .TextMatrix(r, .ColIndex("Amount")) = Val("" & rs("PreTax")) + Val("" & rs("Tax"))
            .TextMatrix(r, .ColIndex("Tax")) = "" & rs("Tax")
            .TextMatrix(r, .ColIndex("Discount")) = "" & rs("Discount")
            .TextMatrix(r, .ColIndex("ReceivedDate")) = "" & rs("ReceivedDate")
            .TextMatrix(r, .ColIndex("PaymentDate")) = "" & rs("PaymentDate")
            .TextMatrix(r, .ColIndex("DiscountDate")) = "" & rs("DiscountDate")
            .TextMatrix(r, .ColIndex("AccountingDate")) = "" & rs("AccountingDate")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("LienVoucher")) = "" & rs("LienVoucher")
            .TextMatrix(r, .ColIndex("InvoiceCode1")) = "" & rs("InvoiceCode1")
            .TextMatrix(r, .ColIndex("InvoiceCode2")) = "" & rs("InvoiceCode2")
            .TextMatrix(r, .ColIndex("PostingDate")) = "" & rs("PostingDate")
            .TextMatrix(r, .ColIndex("BatchNumber")) = "" & rs("BatchNumber")
            .TextMatrix(r, .ColIndex("Comments")) = "" & rs("Comments")
            
            .TextMatrix(r, .ColIndex("Approver")) = "" & rs("Approver")
            
            
            .TextMatrix(r, .ColIndex("Source")) = "" & rs("Source")
            If Trim("" & rs("Source")) = "Web Invoice" Then
                .Cell(flexcpPicture, r, .ColIndex("Source")) = imgIconWeb.Picture
            Else
                .Cell(flexcpPicture, r, .ColIndex("Source")) = imgIconDesk.Picture
            End If

            .Cell(flexcpChecked, r, .ColIndex("RetainageInvoice")) = IIf(rs("RetainageInvoice"), flexChecked, flexUnchecked)
            .Cell(flexcpForeColor, r, .ColIndex("Amount")) = IIf(.ValueMatrix(r, .ColIndex("Amount")) < -0.001, vbRed, vbWindowText)
            .Cell(flexcpForeColor, r, .ColIndex("Tax")) = IIf(.ValueMatrix(r, .ColIndex("Tax")) < -0.001, vbRed, vbWindowText)
            .Cell(flexcpForeColor, r, .ColIndex("Discount")) = IIf(.ValueMatrix(r, .ColIndex("Discount")) < -0.001, vbRed, vbWindowText)
            
            .TextMatrix(r, .ColIndex("UStmp")) = "" & rs("UStmp")
            .TextMatrix(r, .ColIndex("DStmp")) = "" & rs("DStmp")
            .TextMatrix(r, .ColIndex("TStmp")) = "" & rs("TStmp")
            
            
            cmdStop.Enabled = True
            DoEvents
            
            rs.MoveNext
        Wend

    End With
    cmdStop.Enabled = False
    Screen.MousePointer = vbDefault
End Sub


Private Function FileExtensionIndex(hList As Long, hIcon As Long) As Long
    Dim c As New cVBALSysImageList
    Dim h As Long
    On Error Resume Next
    FileExtensionIndex = IconsSmall.ItemIndex("FILEEXT" & hIcon) - 1
    If err.Number <> 0 Then
        c.Create
        h = c.ItemCopyOfIcon(hIcon)
        Call IconsSmall.AddFromHandle(h, IMAGE_ICON, "FILEEXT" & hIcon)
        FileExtensionIndex = IconsSmall.ItemIndex("FILEEXT" & hIcon) - 1
        Call DestroyIcon(h)
    End If
End Function

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim d As Double
    
    
    'set up form - probably wrong place to do this but it works
    Call ApplySettings
        
    FrameEnter.BorderStyle = 0
    FrameEnter.Move ControlPanel.Width, IIf(Me.picWarningBanner.Visible, picWarningBanner.Height, 0), Me.ScaleWidth - ControlPanel.Width, Me.ScaleHeight - IIf(Me.picWarningBanner.Visible, picWarningBanner.Height, 0)
        Toolbar.Move 0, 0, FrameEnter.Width
        gEnterDist.Move margin, gEnterDist.Top, FrameEnter.Width - 2 * margin, FrameEnter.Height - gEnterDist.Top - gEnterTail.Height - 2 * margin
        With gEnterTail
            .Move margin, gEnterDist.Top + gEnterDist.Height + margin, gEnterDist.Width
            
            Call .AutoSize(0)
            Call .AutoSize(2)
            Call .AutoSize(4)
            .ColWidth(5) = 1200
            .ColWidth(6) = 900
            .ColWidth(7) = 1200
            
            d = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7))
            .ColWidth(1) = Min(3000, d / 2)
            .ColWidth(3) = d - .ColWidth(1)
        End With
        
        
    FrameView.BorderStyle = 0
    FrameView.Move FrameEnter.left, FrameEnter.Top, FrameEnter.Width, FrameEnter.Height
        FrameSearch.Move 0, 0, Me.ScaleWidth
        Slider.Min = IIf(FrameSearch.Visible, FrameSearch.Height + 600, 600)
        Slider.Max = FrameView.Height - 600
        Slider.Move margin, Max(Slider.Min, Min(Slider.Max, Slider.Top)), FrameView.Width - 2 * margin, margin
        If FrameSearch.Visible Then
            gViewInv.Move margin, FrameSearch.Height + margin, FrameView.Width - 2 * margin, Slider.Top - FrameSearch.Height - margin
            gViewDist.Move margin, Slider.Top + Slider.Height, FrameView.Width - 2 * margin, FrameView.Height - gViewInv.Top - gViewInv.Height - 2 * margin
            lblSearchUser.Visible = App.Admin
            txtSearchUser.Visible = App.Admin
        Else
            gViewInv.Move margin, margin, FrameView.Width - 2 * margin, Slider.Top - margin
            gViewDist.Move margin, Slider.Top + Slider.Height, FrameView.Width - 2 * margin, FrameView.Height - gViewInv.Top - gViewInv.Height - 2 * margin
        End If
        
    
End Sub

Private Sub Form_Terminate()
   If (Forms.Count = 0) Then UnloadApp
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gEnterDist)
    Call IniPutGrid(Me, gViewInv)
    Call IniPutGrid(Me, gViewDist)
End Sub

Private Sub numAmount_Change()
On Error Resume Next
    Dim TaxRate As Double
    
    TaxRate = Val("" & HFApp.SqlExec("select grouprate FROM taxgroups where DivisionID =" & HFApp.DivisionID & " and taxgroup=" & DbQuote(Str, Trim(mVDefTaxGroup)), dbHomeFront)(0))
    
    TaxRate = 1 + TaxRate / 100
    numTax = numAmount - (numAmount / TaxRate)
    numDiscount = (numAmount - numTax) * mDiscPercent / 100
    Call RecalcTotals
    Dirty = True
End Sub

Private Sub numDiscount_Change()
    Dirty = True
End Sub

Private Sub numTax_Change()
    numDiscount = (numAmount - numTax) * mDiscPercent / 100
    Call RecalcTotals
    Dirty = True
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub

Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
    
        Case KeyCode = vbKeyEscape:            mStopSearch = True
                                               KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("cancel"))
                                               
        Case Not Toolbar.Visible:              Exit Sub
        
        Case KeyCode = vbKeyF11:                            Call POInquiry
        
        Case KeyCode = vbKeyN And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("new"))
        Case KeyCode = vbKeyP And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("quickpay"))
        Case KeyCode = vbKeyD And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("add"))
        Case KeyCode = vbKeyR And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("var"))
        Case KeyCode = vbKeyO And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("post"))
        Case KeyCode = vbKeyA And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("approved"))
        Case KeyCode = vbKeyE And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("pending"))
        Case KeyCode = vbKeyH And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("hold"))
        Case KeyCode = vbKeyC And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("cancel"))
        
        Case KeyCode = vbKeyS And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("scan"))
        Case KeyCode = vbKeyU And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("documents"))
        Case KeyCode = vbKeyM And Shift > 1:   KeyCode = 0: Call Toolbar_ButtonClick(Toolbar.Buttons("email"))
        
    End Select
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim s  As String
    Dim b  As Boolean
    
    If Button.Enabled = False Then Exit Sub
    Select Case Button.Key
            
        Case "new", "cancel"
            If gEnterDist.EditWindow = 0 Then 'not editing a cell
                If Dirty Then
                    Select Case MsgBox("This invoice has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, App.ProductName)
                        Case vbCancel: Exit Sub
                        Case vbYes:    If Not SaveInvoice() Then Exit Sub
                        Case vbNo:     Dirty = False
                    End Select
                End If
                Call OpenInvoice("", "")
            End If
        
        Case "quickpay"
            If Dirty Then
                Select Case MsgBox("This invoice has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, App.ProductName)
                    Case vbCancel: Exit Sub
                    Case vbYes:    If Not SaveInvoice() Then Exit Sub
                End Select
            End If
            Call OpenInvoice("", "", True)
            If FPickList.Choose(HFApp.Databases(AccountingDB), "PO", SelectPO) Then
                txtVendor = Trim(FPickList.SelectedItem("Vendor"))
                mLoading = True
                txtVendor.Enabled = False
                mLoading = False
                Call txtVendor_Validate(b)
                txtInvoice.Text = FPickList.SelectedItem("PO")
                lblVendor.Enabled = False
                lblInvoice.Enabled = False
                dteInvoice.Value = VBA.Date
                numAmount = Val(FPickList.SelectedItem("RemainingAmt"))
                If AccountingDB = dbTimberlinePVdata Then
                    Call AddPO("isub=" & DbQuote(Str, FormatTSField(12, FPickList.SelectedItem(1))))
                Else
                    Call AddPO("isub=" & DbQuote(Str, FPickList.SelectedItem(1)))
                End If
                numTax = gEnterTail.Cell(flexcpValue, 1, 6)
                Call GetInvoiceNumber
                
                txtInvoice.Enabled = True
                Call SetCtrlFocus(txtInvoice)
            End If
        
        Case "add"
            If txtJob.Text <> "" Or txtVendor.tag <> "" Then
                Call txtJob_Validate(False)
                If Not mReadOnly Then
                    If FCommitments.Choose(txtVendor.tag, txtJob.Text, InvoiceRemaining) Then
                        Call AddPO(FCommitments.WhereClause)
                    End If
                End If
            End If
            
        Case "var"
            Call VarianceRemaining
            
            
        Case "email"
            Call MailDocuments(False)
            
        Case "documents"
            Call HFApp.RunTask("EditAttachments|" & InvoiceObjectID & "|Invoice " & txtInvoice.Text)
            Call SetDocButtonImage
        
        Case "approved", "hold", "pending"
            If SaveInvoice(Button.Key) Then Call OpenInvoice("", "")
        
        Case "post"
            If Not SaveInvoice("approved") Then Exit Sub
            If App.PostInvoices Then
                If InvoiceApproved("" & mInvoiceID) Then
                    Call PostInvoices(, "" & mInvoiceID)
                    Call OpenInvoice("", "")
                Else
                    MsgBox "Invoice has not been approved for posting.", vbInformation, App.ProductName
                End If
            Else
                MsgBox "Your security settings do not allow you to post invoices." & vbCrLf & "Contact an administrator to change your permissions.", vbInformation, App.ProductName
            End If
            
            
    End Select
    Exit Sub
eh: Call ErrHandler(SRCFILE & "Toolbar_ButtonClick")
End Sub

Private Function InvoiceObjectID(Optional Vendor As String, Optional Invoice As String) As String
    If Vendor = "" Then
        'InvoiceObjectID = "AP~" & txtVendor.Text & "~" & txtInvoice.Text
        InvoiceObjectID = "AP~" & txtVendor.tag & "~" & txtInvoice.Text
    Else
        InvoiceObjectID = "AP~" & Vendor & "~" & Invoice
    End If
End Function








Private Sub txtDescription_Change()
    Dirty = True
End Sub

Private Sub txtDescription_LostFocus()
    gEnterDist.Col = 1
End Sub


Private Sub txtJob_Change()
    Dirty = True
End Sub


Private Sub txtVendor_GotFocus()
    SelectAll txtVendor
End Sub

Private Sub txtVendor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim b As Boolean
    Select Case KeyCode
        Case vbKeyReturn
            If txtVendor <> "" Then
                Call txtVendor_Validate(b)
                If Not b Then Call TabToNextCtrl(Me)
            End If
  
        Case vbKeyF4:     If Shift = 0 Then Call lblVendor_Click
    End Select
End Sub
Private Sub txtInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim b As Boolean
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblInvoice_Click
        Case vbKeyReturn
            If txtInvoice.Text <> "" Then
                Call txtInvoice_Validate(b)
                If Not b Then Call SetCtrlFocus(dteInvoice)
            End If
    End Select
End Sub
Private Sub txtJob_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblJob_Click
        Case vbKeyReturn: Call TabToNextCtrl(Me)
    End Select
End Sub






Private Sub lblInvoice_Click()
    Dim s As String
    Dim b As Boolean
    
    SetCtrlFocus txtInvoice
    If Not txtInvoice.Enabled Then Exit Sub
    
    
    s = ""
    s = s & "SELECT Vendor" & vbCrLf
    s = s & "      ,VendorName" & vbCrLf
    s = s & "      ,Invoice" & vbCrLf
    s = s & "      ,InvoiceDate" & vbCrLf
    s = s & "      ,Description" & vbCrLf
    s = s & "      ,Status" & vbCrLf
    If App.Options(InvoiceCode1Usage) <> "Hidden" Then
        s = s & "      ,InvoiceCode1 """ & App.Options(InvoiceCode1Label) & """" & vbCrLf
    End If
    If App.Options(InvoiceCode2Usage) <> "Hidden" Then
        s = s & "      ,InvoiceCode2 """ & App.Options(InvoiceCode2Label) & """" & vbCrLf
    End If
    s = s & "      ,DStmp DateEntered " & vbCrLf
    s = s & "  FROM Invoices" & vbCrLf
    s = s & " WHERE DivisionID =" & HFApp.DivisionID & vbCrLf
    If txtVendor.Text <> "" Then
        s = s & "   and Vendor=" & DbQuote(Str, txtVendor.tag)
    End If
    
    If FPickList.Choose(HFApp.Databases(dbHomeFront), "Invoice", s, txtInvoice.Text, False, , , IIf(txtVendor.Text <> "", "Vendor,VendorName", "")) Then
        txtVendor.Text = FPickList.SelectedItem("vendor")
        txtVendor.tag = txtVendor.Text
        txtInvoice.Text = FPickList.SelectedItem("invoice")
        If OpenInvoice(txtVendor.tag, txtInvoice, True) Then
            On Error Resume Next
            txtJob.SetFocus
        End If
    End If
End Sub


Private Sub lblJob_Click()
    Dim b As Boolean
    SetCtrlFocus txtJob
    If Not txtJob.Enabled Then Exit Sub
    If FPickList.Choose(HFApp.Databases(AccountingDB), App.Options(Caption_Job), SelectJob, txtJob.Text) Then
        txtJob.Text = FPickList.SelectedItem(1)
        lblJobDesc = FPickList.SelectedItem(2)
        Call txtJob_Validate(b)
        If Not b Then
            Call SetCtrlFocus(txtJob)
            Call GetApprover(txtJob.Text)
            On Error Resume Next
            numAmount.SetFocus
        End If
    End If
End Sub

Private Sub txtVendor_Validate(Cancel As Boolean)
On Error GoTo eh
Static bInHere As Boolean
If mLoading Then Exit Sub
If bInHere Then Exit Sub
If Not txtVendor.Visible Then Exit Sub
bInHere = True
    
    
    Dim s  As String
    Dim rs As Recordset
    
    If Trim(txtVendor.Text) = "" Then
        Cancel = True
        MsgBox "Vendor is required", vbExclamation, App.ProductName
    Else
        s = ""
        s = s & "SELECT vendor_id vendor" & vbCrLf
        s = s & "      ,vendor_name   Name" & vbCrLf
        s = s & "      ,tradetype VendorType  " & vbCrLf
        s = s & "      ,addr1  Addr1" & vbCrLf
        s = s & "      ,addr2  Addr2" & vbCrLf
        s = s & "      ,city   City" & vbCrLf
        s = s & "      ,state  Prov" & vbCrLf
        s = s & "      ,zip    Postal" & vbCrLf
        s = s & "      ,PayTermDiscPercent DiscPercent" & vbCrLf
        s = s & "      ,PayTermDiscDays DiscDays" & vbCrLf
        s = s & "      ,PayTermDiscType DiscType" & vbCrLf
        s = s & "      ,PayTermNetDays NetDays" & vbCrLf
        s = s & "      ,PayTermNetType NetType" & vbCrLf
        s = s & "      ,PaymentTerms" & vbCrLf
        s = s & "      ,MiscDeductionRate" & vbCrLf
        s = s & "      ,CostCode  DefPhase" & vbCrLf
        s = s & "      ,Category  DefCategory" & vbCrLf
        s = s & "      ,DebitAccount  DefDebitAccount" & vbCrLf
        s = s & "      ,MaterialTaxGroup DefTaxGroup" & vbCrLf
        s = s & "      ,glinsexpdate   GlInsExpDate" & vbCrLf
        s = s & "      ,wcinsexpdate   WcInsExpDate" & vbCrLf
        s = s & "      ,umbinsexpdate  UmbInsExpDate" & vbCrLf
        s = s & "      ,autoinsexpdate AutoInsExpDate" & vbCrLf
        s = s & "      ,glinsRequired" & vbCrLf
        s = s & "      ,wcinsRequired" & vbCrLf
        s = s & "      ,umbinsRequired" & vbCrLf
        s = s & "      ,autoinsRequired" & vbCrLf
        s = s & "      ,isnull(inactive,0) Inactive" & vbCrLf
        s = s & "  from tblvendors" & vbCrLf
        Set rs = HFApp.SqlExec(s & " WHERE isnull(isTBD,0)=0 and DivisionID =" & HFApp.DivisionID & " and vendor_id=" & DbQuote(Str, txtVendor), dbHomeFront)
        If rs.EOF Then Set rs = HFApp.SqlExec(s & " WHERE isnull(isTBD,0)=0 and DivisionID =" & HFApp.DivisionID & " and vendor_name=" & DbQuote(Str, txtVendor), dbHomeFront)
        If rs.EOF Then
            txtAddress = ""
            MsgBox "Vendor not found", vbExclamation, App.ProductName
            Cancel = True
        Else
            txtCompany.ForeColor = vbWindowText
            
            If HFApp.Options(AccountingSystem) = asTimberline Then
                txtVendor = Trim("" & rs("vendor"))
            Else
                txtVendor = Trim("" & rs("name"))
            End If
            
            txtVendor.tag = Trim("" & rs("vendor"))
            txtCompany = Trim("" & rs("name"))
            txtAddress = Trim("" & rs("Addr1")) & vbCrLf & _
                         IIf(Trim("" & rs("Addr2")) & vbCrLf = vbCrLf, "", Trim("" & rs("Addr2")) & vbCrLf) & _
                         Trim("" & rs("City")) & ", " & Trim("" & rs("Prov")) & vbCrLf & _
                         Trim("" & rs("Postal")) & vbCrLf
            
            If "" & rs("Inactive") = "True" Then
                txtAddress = txtCompany & vbCrLf & txtAddress
                txtCompany = "INACTIVE"
                txtCompany.ForeColor = vbRed
            End If


            mWrapInsuranceRate = Val("" & rs("MiscDeductionRate")) / 100
            lblWrapInsuranceRate.Caption = Format(mWrapInsuranceRate, "percent")
            
            mVendorType = "" & rs("VendorType")
            If mVendorType = "Summary" Then
                txtCompany.Locked = False
                txtAddress.Locked = False
                txtCompany.TabStop = True
                txtAddress.TabStop = True
                txtCompany.Text = ""
                txtAddress.Text = ""
            Else
                txtCompany.Locked = True
                txtAddress.Locked = True
                txtCompany.TabStop = False
                txtAddress.TabStop = False
            End If
            
            
            'defaults
            mVDefPhase = Trim("" & rs("DefPhase"))
            mVDefCategory = Trim("" & rs("DefCategory"))
            mVDefDebitAccount = Trim("" & rs("DefDebitAccount"))
            
            mVDefTaxGroup = Trim("" & rs("DefTaxGroup"))
            If mVDefTaxGroup = "" Then mVDefTaxGroup = App.Options(DefaultTaxGroup)

            'discount and terms
            mNetType = "" & rs("NetType")
            mNetDays = Val("" & rs("NetDays"))
            mDiscType = "" & rs("DiscType")
            mDiscDays = Val("" & rs("DiscDays"))
            mDiscPercent = Val("" & rs("DiscPercent"))
            lblTerms = "" & rs("PaymentTerms")
            
            'check insurance
            s = ""
            If "" & rs("GlInsRequired") = "True" And IsDate(rs("GlInsExpDate")) Then
                If Now > rs("GlInsExpDate") Then s = s & vbCrLf & "General Liability insurance expired on " & Format(rs("GlInsExpDate"), App.Options(TimberlineDateFormat))
            End If
            If "" & rs("AutoInsRequired") = "True" And IsDate(rs("AutoInsExpDate")) Then
                If Now > rs("AutoInsExpDate") Then s = s & vbCrLf & "Automobile insurance expired on " & Format(rs("AutoInsExpDate"), App.Options(TimberlineDateFormat))
            End If
            If "" & rs("WcInsRequired") = "True" And IsDate(rs("WcInsExpDate")) Then
                If Now > rs("WcInsExpDate") Then s = s & vbCrLf & "Workers' Comp insurance expired on " & Format(rs("WcInsExpDate"), App.Options(TimberlineDateFormat))
            End If
            If "" & rs("UmbInsRequired") = "True" And IsDate(rs("UmbInsExpDate")) Then
                If Now > rs("UmbInsExpDate") Then s = s & vbCrLf & "Umbrella insurance expired on " & Format(rs("UmbInsExpDate"), App.Options(TimberlineDateFormat))
            End If
            If s <> "" Then
                s = "Warning!" & vbCrLf & vbCrLf & "Vendor's insurance is not up-to-date." & vbCrLf & s
                MsgBox s, vbExclamation, App.ProductName
            End If
            
            
            'check for back charged PO's
            Call CheckBackChargePOs(txtVendor.tag)
            
        End If
        
        Call dteInvoice_Change
        Call dteReceived_Change
        Dirty = False
        
    End If
        
    
        
bInHere = False
    Exit Sub
eh: Call ErrHandler(SRCFILE & "txtVendor_Validate")
End Sub

Private Sub txtJob_Validate(Cancel As Boolean)
On Error GoTo eh
    Dim s  As String
    Dim rs As Recordset
    
    If Trim(txtJob.Text) = "" Then
        txtJob = ""
        lblJobDesc = ""
        Exit Sub
    Else
        Set rs = ValidateJob(txtJob.Text)
        If rs.EOF Then
            lblJobDesc = "job not found"
        Else
            
            If "" & rs("Status") = "Closed" Then
                If MsgBox("Job is closed.", vbInformation + vbOKCancel, App.ProductName) = vbCancel Then
                    Cancel = True
                    Exit Sub
                End If
            End If
            
            txtJob.Text = "" & rs("Job")
            lblJobDesc = Trim("" & rs("description"))
            
            If App.Options(InvDescDefaultsToJobAddr) And Trim(txtDescription.Text) = "" Then
                txtDescription.Text = "" & rs("Municipal_address")
            End If
            
            
        End If
    
    
    End If
    
    
    Call GetApprover(txtJob.Text)
    
    Exit Sub
eh: Call ErrHandler(SRCFILE & "txtJob_Validate")
End Sub

Private Sub txtInvoice_Validate(Cancel As Boolean)
On Error GoTo eh

Static bInHere As Boolean
If mLoading Then Exit Sub
If bInHere Then Exit Sub
bInHere = True

    If Trim(txtInvoice.Text) = "" Then
        Cancel = True
        MsgBox "Invoice is required", vbExclamation, App.ProductName
    Else
        If mDirty Then
            'still editting
        Else
            Cancel = Not OpenInvoice(txtVendor.tag, txtInvoice)
        End If
    End If

bInHere = False
    Exit Sub
eh: Call ErrHandler(SRCFILE & "txtInvoice_Validate")
bInHere = False
End Sub


Private Function SaveInvoice(Optional Mode As String = "pending") As Boolean
'mode = hold,pending,approved
On Error GoTo eh

    Dim s        As String
    Dim c        As String
    Dim r      As Long
    Dim ItemID   As Long
    Dim rc       As Long
    Dim Status As String
    Dim Addr1 As String
    Dim Addr2 As String
    Dim City As String
    Dim Prov As String
    Dim Postal As String
    
    
    If dteAccounting.Text = "" Then dteAccounting = dteInvoice
    If Not ValidateInvoice(Mode, App.Options.Value(SaveIncompleteInvoices)) Then Exit Function
    
    
    'this will return the number of months between the invoices accounting date and todays date
    ' 2 means the accting date is 2 months ago
    ' 0 means the accting date is current
    '-2 means the accting date is 2 months from now
    If App.Options(WarnIfAcctDateNotCurrent) And DateDiff("m", dteAccounting, Now()) >= 1 Then
        rc = MsgBox("The accounting date is not current." & vbCrLf & vbCrLf & "Do you want to use todays date?", vbYesNoCancel + vbExclamation, App.ProductName)
        Select Case rc
            Case vbCancel: Exit Function
            Case vbNo:    'use whatever was entered
            Case vbYes:    dteAccounting = DateValue(Now())
        End Select
    End If

    If IsPOInvoice Then
        If App.Options(AutoApprovePOInvoices) Then Mode = "Approved"
        If App.Options(SchedulingUsage) = HoldUntilCompleted And Not ArePOsCompleted Then Mode = "Hold"
    End If
    
    Select Case True
        Case Mode = "hold":         Status = "Hold"
        Case Mode = "pending":      Status = "Pending"
        Case Mode = "approved":     Status = "Approved"
        Case Else:                  MsgBox "Whoa! What?": Stop
    End Select
    
    If Status = "Hold" And App.Options(AutoEmail) Then Call MailDocuments(True)
        
        
        
        
    
    
    'save headers
    Call ParseAddress(txtAddress, Addr1, Addr2, City, Prov, Postal)
    mInvoiceID = WriteInvoice(mInvoiceID, False, _
                              txtVendor.tag, mVendorType, txtCompany, Addr1, Addr2, City, Prov, Postal, _
                              txtInvoice, txtJob, lblJobDesc, Status, cboApprover.Text, txtComments, GetComboBoxListID(cboDepartment), numAmount - numTax, numTax, numDiscount, _
                              numWrapInsuranceAmount, mWrapInsuranceRate, dteDiscount, dteReceived, dteInvoice, dtePayment, dteAccounting, txtDescription, _
                              txtInvoiceCode1, txtInvoiceCode2, lblWarning, False, 0)

    
    'save details
    With gEnterDist
    
        'delete rows first
        For r = .Rows - 2 To 1 Step -1
            If .RowHidden(r) Then
                Call HFApp.SqlExec("DELETE FROM InvoiceItems WHERE ItemID=" & DbQuote(num, .ValueMatrix(r, .ColIndex("ItemID"))))
                Call .RemoveItem(r)
            End If
        Next
        
        'now insert or update items
        For r = 1 To .Rows - 2
            .TextMatrix(r, .ColIndex("ItemID")) = WriteInvoiceItem(mInvoiceID, .ValueMatrix(r, .ColIndex("ItemID")), txtVendor.tag, txtInvoice, _
                                                                  .TextMatrix(r, .ColIndex("CommitmentVendor")), .TextMatrix(r, .ColIndex("Commitment")), .TextMatrix(r, .ColIndex("CommitmentDesc")), .TextMatrix(r, .ColIndex("CommitmentItem")), .TextMatrix(r, .ColIndex("Job")), .TextMatrix(r, .ColIndex("JobDesc")), .TextMatrix(r, .ColIndex("Extra")), .TextMatrix(r, .ColIndex("ExtraDesc")), .TextMatrix(r, .ColIndex("Phase")), .TextMatrix(r, .ColIndex("PhaseDesc")), .TextMatrix(r, .ColIndex("Category")), .TextMatrix(r, .ColIndex("CategoryDesc")), .TextMatrix(r, .ColIndex("DebitAccount")), .TextMatrix(r, .ColIndex("DebitAccountDesc")), .TextMatrix(r, .ColIndex("Equipment")), .TextMatrix(r, .ColIndex("EquipmentDesc")), .TextMatrix(r, .ColIndex("EQCostCode")), .TextMatrix(r, .ColIndex("EQCostCodeDesc")), _
                                                                  .TextMatrix(r, .ColIndex("TaxGroup")), .TextMatrix(r, .ColIndex("TaxGroupDesc")), .ValueMatrix(r, .ColIndex("TaxRate")), _
                                                                  .ValueMatrix(r, .ColIndex("CommittedQuantity")), .ValueMatrix(r, .ColIndex("CommittedUnitPrice")), _
                                                                  .ValueMatrix(r, .ColIndex("InvoicedQuantity")), .ValueMatrix(r, .ColIndex("InvoicedUnitPrice")), _
                                                                  .ValueMatrix(r, .ColIndex("PreTax")), .ValueMatrix(r, .ColIndex("Tax")), _
                                                                  .ValueMatrix(r, .ColIndex("Retainage")), .ValueMatrix(r, .ColIndex("RetainageRate")), _
                                                                  .TextMatrix(r, .ColIndex("Description")), .Cell(flexcpChecked, r, .ColIndex("Billable")) = flexChecked, .Cell(flexcpChecked, r, .ColIndex("WrapInsuranceExempt")) = flexChecked, .TextMatrix(r, .ColIndex("JointPayee")))
        Next
    End With
    
    Call HFApp.SqlExec("exec dbo.Payables_CreateHBInvoice " & DbQuote(num, mInvoiceID))
    Call WriteApprovalLog(mInvoiceID, Status)
    Dirty = False
    
    
    SaveInvoice = True
    Exit Function
    
    
eh: If InStr(1, err.Description, "duplicate", vbTextCompare) Then
        MsgBox "Unable to save this invoice. It has already been entered.", vbExclamation, App.ProductName
    Else
        Call ErrHandler(SRCFILE & "SaveInvoice" & vbCrLf & err.Description, s)
    End If
End Function


Private Sub txtInvoiceCode1_Validate(Cancel As Boolean)
On Error GoTo eh

    If App.Options(InvoiceCode1Usage) = "Required" And txtInvoiceCode1.Text = "" Then
        MsgBox App.Options(InvoiceCode1Label) & " is required", vbExclamation, App.ProductName
        Cancel = True
        Exit Sub
    End If
    
    Select Case left(App.Options(InvoiceCode1Validate), 4)
        Case "None"
        Case "Stri"
            If InvoiceCodeExists(txtVendor.tag, txtInvoice.Text, 1, txtInvoiceCode1.Text) Then
                MsgBox App.Options(InvoiceCode1Label) & " """ & txtInvoiceCode1.Text & """ has already been entered.", vbExclamation, App.ProductName
                Cancel = True
                Exit Sub
            End If
        Case "Warn"
            If InvoiceCodeExists(txtVendor.tag, txtInvoice.Text, 1, txtInvoiceCode1.Text) Then
                If vbCancel = MsgBox(App.Options(InvoiceCode1Label) & " """ & txtInvoiceCode1.Text & """ has already been entered.", vbOKCancel + vbExclamation, App.ProductName) Then
                    Cancel = True
                    Exit Sub
                End If
            End If
    End Select
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "txtInvoiceCode1_Validate")
End Sub

Private Sub txtInvoiceCode2_Validate(Cancel As Boolean)
On Error GoTo eh

    If App.Options(InvoiceCode2Usage) = "Required" And txtInvoiceCode2.Text = "" Then
        MsgBox App.Options(InvoiceCode2Label) & " is required", vbExclamation, App.ProductName
        Cancel = True
        Exit Sub
    End If
    
    Select Case left(App.Options(InvoiceCode2Validate), 4)
        Case "None"
        Case "Stri"
            If InvoiceCodeExists(txtVendor.tag, txtInvoice.Text, 2, txtInvoiceCode2.Text) Then
                MsgBox App.Options(InvoiceCode2Label) & " """ & txtInvoiceCode2.Text & """ has already been entered.", vbExclamation, App.ProductName
                Cancel = True
                Exit Sub
            End If
        Case "Warn"
            If InvoiceCodeExists(txtVendor.tag, txtInvoice.Text, 2, txtInvoiceCode2.Text) Then
                If vbCancel = MsgBox(App.Options(InvoiceCode2Label) & " """ & txtInvoiceCode2.Text & """ has already been entered.", vbOKCancel + vbInformation, App.ProductName) Then
                    Cancel = True
                    Exit Sub
                End If
            End If
    End Select
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "txtInvoiceCode2_Validate")
End Sub


Private Function ValidateInvoice(Mode As String, HeaderOnly As Boolean) As Boolean
'mode = save,hold,approve,test
'partial only validates the header section - lines are left to be completed in approval portal and validated before posting.
' test is used by dirty() to see if invoice is complete. display no messages. just return true/false


    Dim rs As Recordset
    Dim r As Long
    Dim s As String
    Dim sErrors   As String
    Dim sWarnings As String
    
    Dim Accounts  As String
    
    
    Dim Commitment     As Boolean
    Dim CommitmentItem As Boolean
    Dim Job            As Boolean
    Dim Phase          As Boolean
    Dim Category       As Boolean
    Dim DebitAccount   As Boolean
    Dim CrossPay       As Boolean
    Dim TaxGroup       As Boolean
    Dim TaxValue       As Boolean
    
    Dim TotalTax       As Double
    
    
    
    If Mode <> "test" Then
        If ExceedsLimit() And Mode = "approved" Then sErrors = sErrors & " " & vbBullet & " Exceeds your approval limit" & vbCrLf
    End If
    If HasV99 And Mode <> "hold" Then sErrors = sErrors & " " & vbBullet & " Invoice has an unknown variance (" & mV99 & ")" & vbCrLf
    If txtVendor = "" Then sErrors = sErrors & " " & vbBullet & " Vendor is required" & vbCrLf
    If txtInvoice = "" Then sErrors = sErrors & " " & vbBullet & " Invoice number is required" & vbCrLf
    If App.Options(DepartmentIsRequired) And cboDepartment.Text = "" And Not IsPOInvoice() Then sErrors = sErrors & " " & vbBullet & " Department is required" & vbCrLf
    
    'these have options but are "builtin" no controls on settings screen
    If App.Options(InvoiceDateRequired) And dteInvoice.ValueIsNull Then sErrors = sErrors & " " & vbBullet & " Invoice date is required" & vbCrLf
    If App.Options(AccountingDateRequired) And dteAccounting.ValueIsNull Then sErrors = sErrors & " " & vbBullet & " Accounting date is required" & vbCrLf
    If App.Options(PaymentDateRequired) And dtePayment.ValueIsNull Then sErrors = sErrors & " " & vbBullet & " Payment date is required" & vbCrLf
    If App.Options(ReceivedDateRequired) And dteReceived.ValueIsNull Then sErrors = sErrors & " " & vbBullet & " Received date is required" & vbCrLf
    
    If App.Options(InvoiceCode1Usage) = "Required" And txtInvoiceCode1.Text = "" Then
        sErrors = sErrors & " " & vbBullet & " " & App.Options(InvoiceCode1Label) & " is required" & vbCrLf
    End If
    Select Case left(App.Options(InvoiceCode1Validate), 4)
        Case "None"
        Case "Stri"
            If InvoiceCodeExists(txtVendor.tag, txtInvoice.Text, 1, txtInvoiceCode1.Text) Then
                sErrors = sErrors & " " & vbBullet & " " & App.Options(InvoiceCode1Label) & " """ & txtInvoiceCode1.Text & """ has already been entered." & vbCrLf
            End If
        Case "Warn"
            If InvoiceCodeExists(txtVendor.tag, txtInvoice.Text, 1, txtInvoiceCode1.Text) Then
                sWarnings = sWarnings & " " & vbBullet & " " & App.Options(InvoiceCode1Label) & " """ & txtInvoiceCode1.Text & """ has already been entered." & vbCrLf
            End If
    End Select
    
    If App.Options(InvoiceCode2Usage) = "Required" And txtInvoiceCode2.Text = "" Then
        sErrors = sErrors & " " & vbBullet & " " & App.Options(InvoiceCode2Label) & " is required" & vbCrLf
    End If
    Select Case left(App.Options(InvoiceCode2Validate), 4)
        Case "None"
        Case "Stri":
            If InvoiceCodeExists(txtVendor.tag, txtInvoice.Text, 2, txtInvoiceCode2.Text) Then
                sErrors = sErrors & " " & vbBullet & " " & App.Options(InvoiceCode2Label) & " """ & txtInvoiceCode2.Text & """ has already been entered." & vbCrLf
            End If
        Case "Warn"
            If InvoiceCodeExists(txtVendor.tag, txtInvoice.Text, 2, txtInvoiceCode2.Text) Then
                sWarnings = sWarnings & " " & vbBullet & " " & App.Options(InvoiceCode2Label) & " """ & txtInvoiceCode2.Text & """ has already been entered." & vbCrLf
            End If
    End Select
    
ValidationDEBUG = sErrors
    
    If Not HeaderOnly Then
        If InvoiceRemaining <> 0 Then sErrors = sErrors & " " & vbBullet & " " & Format(InvoiceRemaining, "$#,##0.00") & " remaining." & vbCrLf
        With gEnterDist
        
            For r = 1 To .Rows - 2
                If Not .RowHidden(r) Then
                    TotalTax = TotalTax + Val("" & .TextMatrix(r, .ColIndex("Tax")))
                End If
            Next
            If Round(TotalTax, 2) <> Round(Val("" & numTax.Value), 2) And Mode <> "hold" And gEnterDist.Rows > 2 Then
                sWarnings = sWarnings & " " & vbBullet & " Taxes do not balance. The amount entered will be changed to match the amount distributed. (" & Format(TotalTax, "#,##0.00") & ")" & vbCrLf
            End If
            
            For r = 1 To .Rows - 2
                If Not .RowHidden(r) Then
                
                    Commitment = Trim(.TextMatrix(r, .ColIndex("Commitment"))) <> ""
                    CommitmentItem = Trim(.TextMatrix(r, .ColIndex("CommitmentItem"))) <> ""
                    Job = Trim(.TextMatrix(r, .ColIndex("Job"))) <> ""
                    Phase = Trim(.TextMatrix(r, .ColIndex("Phase"))) <> ""
                    Category = Trim(.TextMatrix(r, .ColIndex("Category"))) <> ""
                    DebitAccount = Trim(.TextMatrix(r, .ColIndex("DebitAccount"))) <> ""
                    CrossPay = Commitment And LCase(Trim(.TextMatrix(r, .ColIndex("CommitmentVendor")))) <> LCase(Trim(txtVendor.tag))
                    TaxGroup = Trim(.TextMatrix(r, .ColIndex("TaxGroup"))) <> ""
                    TaxValue = Val("" & .TextMatrix(r, .ColIndex("Tax"))) <> 0
                    
                    If Trim(.TextMatrix(r, .ColIndex("DebitAccount"))) <> "" And Trim(.TextMatrix(r, .ColIndex("Job"))) <> "" Then
                        Accounts = Accounts & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("DebitAccount")))
                    End If
                                    
                    If Commitment And Not (CommitmentItem) Then sErrors = sErrors & " " & vbBullet & " line " & r & " - item, " & App.Options(Caption_Job) & ", " & App.Options(Caption_Phase) & " and " & App.Options(Caption_Category) & " are required" & vbCrLf
                    If Job And Not (Phase And Category) Then sErrors = sErrors & " " & vbBullet & " line " & r & " - " & App.Options(Caption_Phase) & " and " & App.Options(Caption_Category) & " are required" & vbCrLf
                    
                    If Not (Job Or DebitAccount) Then sErrors = sErrors & " " & vbBullet & " line " & r & " - " & "A " & App.Options(Caption_Job) & " or a " & App.Options(Caption_DebitAccount) & " is required" & vbCrLf
                    
                    If HFApp.Options(AccountingSystem) = asSimply And Not DebitAccount Then sErrors = sErrors & " " & vbBullet & " line " & r & " - " & "A " & App.Options(Caption_DebitAccount) & " is required" & vbCrLf
                    
                    If CrossPay And Not App.Options(AllowCrossPayingPOs) Then sErrors = sErrors & " " & vbBullet & " line " & r & " - " & App.Options(Caption_Commitment) & " " & Trim(.TextMatrix(r, .ColIndex("Commitment"))) & " was issued to " & Trim(.TextMatrix(r, .ColIndex("CommitmentVendor"))) & vbCrLf
                    If TaxValue And Not TaxGroup Then sErrors = sErrors & " " & vbBullet & " line " & r & " - Tax Group is required if tax amount is not 0" & vbCrLf
                    
                End If
            Next
            
            Accounts = Mid(Accounts, 2)
            If HFApp.Options(AccountingSystem) = asSimply And Accounts <> "" Then
                Set rs = HFApp.SqlExec("select account,description,isnull(AllowJobAllocation,0) Allocate from glaccounts where DivisionID = " & HFApp.DivisionID & " and allowjoballocation = 0 and account IN(" & Accounts & ")")
                While Not rs.EOF
                    If rs!Allocate = 0 Then
                        sWarnings = sWarnings & " " & vbBullet & " Job allocations are not allowed on account """ & rs("account") & """. The invoice can be posted but it won't be fully allocated." & vbCrLf
                    End If
                    rs.MoveNext
                Wend
            End If
            
        End With
    End If
    
ValidationDEBUG = sErrors
    
    
    If Mode = "test" Then
        ValidateInvoice = sErrors = ""
    Else
        If sErrors = "" Then
            If sWarnings = "" Then
                ValidateInvoice = True
            Else
                If vbOK = MsgBox("Are you sure you want to save this invoice?" & vbCrLf & vbCrLf & sWarnings, vbOKCancel + vbInformation + vbDefaultButton2, App.ProductName) Then
                    ValidateInvoice = True
                    numTax.Value = TotalTax
                Else
                    ValidateInvoice = False
                End If
            End If
        Else
            MsgBox "Unable to save invoice" & vbCrLf & vbCrLf & sErrors, vbExclamation, App.ProductName
            ValidateInvoice = False
        End If
    End If
    
End Function


Private Function OpenInvoice(Vendor As String, Invoice As String, Optional Quiet As Boolean) As Boolean
    Dim s  As String
    Dim r As Long
    Dim rs As Recordset
    Dim miscDeduction As Double
    
    Screen.MousePointer = vbHourglass
    mLoading = True
    ReadOnly = Not App.EditInvoices
        
    On Error Resume Next
    mStatus = ""
    If InvoiceInHomefront(Vendor, Invoice, mStatus) Then
        If IsIn(mStatus, "Lien Voucher") Then
            ReadOnly = True
            If vbCancel = MsgBox(vbQuote & Vendor & vbQuote & " invoice " & vbQuote & Invoice & vbQuote & " has a lien voucher. You cannot edit it." & vbCrLf & vbCrLf & "Do you want to view it?", vbQuestion + vbOKCancel, App.ProductName) Then
                ReadOnly = False
                OpenInvoice = False
                mLoading = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        ElseIf IsIn(mStatus, "Exported", "Posted") And Not App.EditPostedInvoices Then
            ReadOnly = True
            If vbCancel = MsgBox(vbQuote & Vendor & vbQuote & " invoice " & vbQuote & Invoice & vbQuote & " has been posted. You cannot edit it." & vbCrLf & vbCrLf & "Do you want to view it?", vbQuestion + vbOKCancel, App.ProductName) Then
                ReadOnly = False
                OpenInvoice = False
                mLoading = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        Else
            If vbCancel = MsgBox(vbQuote & Vendor & vbQuote & " invoice " & vbQuote & Invoice & vbQuote & " has already been entered." & vbCrLf & vbCrLf & "Do you want to edit it?", vbQuestion + vbOKCancel, App.ProductName) Then
                ReadOnly = False
                OpenInvoice = False
                mLoading = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
    End If
    
    On Error GoTo 0
    
    
    
    'load invoice
    s = ""
    s = s & "select v.vendor_name,i.* " & vbCrLf
    s = s & "  from invoices i left outer join tblvendors v on i.DivisionID = v.DivisionID and i.vendor=v.vendor_id" & vbCrLf
    s = s & " where i.DivisionID =" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & "   and i.vendor=" & DbQuote(Str, Vendor) & vbCrLf
    s = s & "   and i.invoice=" & DbQuote(Str, Invoice) & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    If rs.EOF Then
        txtInvoice = Invoice
        mIsHBInvoice = False
        lblHoldbackInvoice.Caption = ""
        
        numAmount = 0
        numTax = 0
        numDiscount = 0
        numWrapInsuranceAmount = 0
        dteInvoice = Null
        dtePayment = Null
        txtDescription = ""
        txtInvoiceCode1 = ""
        txtInvoiceCode2 = ""
        txtComments = ""
        cboDepartment.ListIndex = -1
        cboApprover.ListIndex = -1
        
        gEnterDist.Rows = 1
        gEnterDist.AddItem ""
        gEnterDist.Cell(flexcpChecked, gEnterDist.Rows - 1, gEnterDist.ColIndex("Billable")) = IIf(mAPInvoicesDefaultToBillable, flexChecked, flexUnchecked)
        gEnterDist.Cell(flexcpData, gEnterDist.Rows - 1, 0, gEnterDist.Rows - 1, gEnterDist.Cols - 1) = "DefRequired"
        gEnterTail.Cell(flexcpText, 0, 1, 3, 1) = ""
        gEnterTail.Cell(flexcpText, 0, 3, 3, 3) = ""
        gEnterTail.Cell(flexcpText, 0, 5, 3, 7) = ""
        RecalcTotals
        OpenInvoice = True
        txtVendor.Enabled = Vendor = "" And Invoice = "" And Not ReadOnly
        lblVendor.Enabled = txtVendor.Enabled
        txtInvoice.Enabled = txtVendor.Enabled
        lblInvoice.Enabled = txtVendor.Enabled
        
        If Vendor = "" Then
            txtCompany.Text = ""
            txtAddress.Text = ""
            lblTerms.Caption = ""
        End If
        
        mInvoiceID = 0
        
        
        On Error Resume Next
        txtVendor.SetFocus
        mLoading = False
        Screen.MousePointer = vbNormal
        Dirty = False
        'clear doc button
        Toolbar.Buttons("documents").Image = "documents"
        
        
        
        Exit Function
    Else
        mInvoiceID = "" & rs("InvoiceID")
        
        'the opened invoice either is or has a holdback invoice associated with it. No new hold backs allowed
        If "" & rs("RetainageInvoice") = "True" Then
            lblHoldbackInvoice.Caption = "This is a " & LCase(App.Options.Value(Caption_Retainage)) & " invoice"
            mIsHBInvoice = True
        End If
        If Val("" & rs("RetainageID")) <> 0 Then
            lblHoldbackInvoice.Caption = "A " & LCase(App.Options.Value(Caption_Retainage)) & " invoice has been created for this invoice"
            mIsHBInvoice = True
        End If
        
        
        txtVendor.Text = "" & rs("vendor")
        mLoading = False
        Call txtVendor_Validate(b)
        
        mLoading = True
        
        mCodeMode = True
        txtInvoice = "" & rs("invoice")
        mCodeMode = False
        txtJob = "" & rs("job")
        lblJobDesc = "" & rs("jobdesc")
        txtDescription = "" & rs("description")
        txtInvoiceCode1 = "" & rs("InvoiceCode1")
        txtInvoiceCode2 = "" & rs("InvoiceCode2")
        txtComments = "" & rs("Comments")
        
        numAmount = Val("" & rs("pretax")) + Val("" & rs("tax"))
        numTax = Val("" & rs("tax"))
        numDiscount = Val("" & rs("discount"))
        dteReceived = "" & rs("ReceivedDate")
        dteInvoice = "" & rs("InvoiceDate")
        
        'this has to be done after recalctotals
        miscDeduction = Val("" & rs("miscdeductionamt"))
        
        dteDiscount = "" & rs("DiscountDate")
        dtePayment = "" & rs("PaymentDate")
        dteAccounting = "" & rs("AccountingDate")
        
        Call SetComboBoxListIndex(cboDepartment, , , Val("" & rs("DeptID")))
        Call SetComboBoxListIndex(cboApprover, "" & rs("Approver"))
        Call SetDocButtonImage
        
        
        'load distributions
        s = ""
        s = s & "select *" & vbCrLf
        s = s & "  from invoiceitems" & vbCrLf
        s = s & " where DivisionID =" & HFApp.DivisionID & " and InvoiceID=" & DbQuote(num, mInvoiceID) & vbCrLf
        s = s & "order by itemid" & vbCrLf
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        
        With gEnterDist
            .Rows = 1
            While Not rs.EOF
                .AddItem ""
                r = .Rows - 1
                
                
                .TextMatrix(r, .ColIndex("itemid")) = "" & rs("itemid")
                .TextMatrix(r, .ColIndex("commitment")) = "" & rs("commitment")
                .TextMatrix(r, .ColIndex("commitmentdesc")) = "" & rs("commitmentdesc")
                .TextMatrix(r, .ColIndex("commitmentitem")) = "" & IIf(rs("commitmentitem") = 0, "", rs("commitmentitem"))
                .TextMatrix(r, .ColIndex("commitmentvendor")) = "" & rs("commitmentvendor")
                
                .TextMatrix(r, .ColIndex("job")) = "" & rs("job")
                .TextMatrix(r, .ColIndex("jobdesc")) = "" & rs("jobdesc")
                .TextMatrix(r, .ColIndex("extra")) = "" & rs("extra")
                .TextMatrix(r, .ColIndex("extradesc")) = "" & rs("extradesc")
                .TextMatrix(r, .ColIndex("phase")) = "" & rs("phase")
                .TextMatrix(r, .ColIndex("phasedesc")) = "" & rs("phasedesc")
                .TextMatrix(r, .ColIndex("category")) = "" & rs("category")
                .TextMatrix(r, .ColIndex("categorydesc")) = "" & rs("categorydesc")
                .TextMatrix(r, .ColIndex("debitaccount")) = "" & rs("debitaccount")
                .TextMatrix(r, .ColIndex("debitaccountdesc")) = "" & rs("debitaccountDesc")
                
                .TextMatrix(r, .ColIndex("taxgroup")) = "" & rs("taxgroup")
                .TextMatrix(r, .ColIndex("taxgroupdesc")) = "" & rs("TaxGroupDesc")
                .TextMatrix(r, .ColIndex("taxrate")) = "" & rs("TaxRate")
                .TextMatrix(r, .ColIndex("pretax")) = "" & rs("Pretax")
                .TextMatrix(r, .ColIndex("tax")) = "" & rs("Tax")
                .TextMatrix(r, .ColIndex("retainage")) = "" & rs("Retainage")
                .TextMatrix(r, .ColIndex("description")) = "" & rs("Description")
                .TextMatrix(r, .ColIndex("committedquantity")) = "" & rs("CommittedQuantity")
                
                .TextMatrix(r, .ColIndex("committedunitprice")) = "" & rs("CommittedUnitPrice")
                .TextMatrix(r, .ColIndex("invoicedquantity")) = "" & rs("InvoicedQuantity")
                .TextMatrix(r, .ColIndex("invoicedunitprice")) = "" & rs("InvoicedUnitPrice")
                .TextMatrix(r, .ColIndex("retainagerate")) = "" & rs("RetainageRate") '* 100
                
                .TextMatrix(r, .ColIndex("JointPayee")) = "" & rs("JointPayee")
                
                .Cell(flexcpChecked, r, .ColIndex("billable")) = IIf("" & rs("billable") = "True", flexChecked, flexUnchecked)
                .Cell(flexcpChecked, r, .ColIndex("WrapInsuranceExempt")) = IIf("" & rs("WrapInsuranceExempt") = "True", flexChecked, flexUnchecked)
                .TextMatrix(r, .ColIndex("EQCostCode")) = "" & rs("EquipmentCostCode")
                .TextMatrix(r, .ColIndex("EQCostCodeDesc")) = "" & rs("EquipmentCostCodeDesc")
                .TextMatrix(r, .ColIndex("Equipment")) = "" & rs("Equipment")
                .TextMatrix(r, .ColIndex("EquipmentDesc")) = "" & rs("EquipmentDesc")
                
                
                
                rs.MoveNext
            Wend
            .AddItem ""
            .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "DefRequired"
            .Cell(flexcpChecked, .Rows - 1, .ColIndex("Billable")) = IIf(mAPInvoicesDefaultToBillable, flexChecked, flexUnchecked)
        End With
        
        OpenInvoice = True
        txtVendor.Enabled = False
        lblVendor.Enabled = False
        txtInvoice.Enabled = False
        lblInvoice.Enabled = False
        
            
    End If
    mLoading = False
    Dirty = False
    
    Call RecalcTotals
    Dirty = False
    'set this last so it doesnt get overwritten by recalculation
    numWrapInsuranceAmount = miscDeduction
    
    Screen.MousePointer = vbNormal
    
        
End Function

Private Sub RecalcRetainage(Row As Long)
    Dim amt As Double
    With gEnterDist
        amt = 0
        amt = amt + .ValueMatrix(Row, .ColIndex("PreTax")) * .ValueMatrix(Row, .ColIndex("RetainageRate")) / 100
        .TextMatrix(Row, .ColIndex("Retainage")) = IIf(.ValueMatrix(Row, .ColIndex("RetainageRate")) = 0, "", Round(amt, 2))
    End With
End Sub

Private Sub RecalcTotals()
    
    Dim Row       As Long
    Dim WrapInsurance   As Double
    Dim Pretax    As Double
    Dim Tax       As Double
    With gEnterDist
        WrapInsurance = 0
        Pretax = 0
        Tax = 0
        mTotalHBAmount = 0
        mTotalHBTax = 0
        
        For Row = 1 To .Rows - 2
        
            .TextMatrix(Row, .ColIndex("PreTax")) = Round(.ValueMatrix(Row, .ColIndex("PreTax")), 2)
            .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("Tax")), 2)
            .TextMatrix(Row, .ColIndex("Retainage")) = Round(.ValueMatrix(Row, .ColIndex("Retainage")), 2)
            
            .TextMatrix(Row, .ColIndex("Retainage")) = IIf(.ValueMatrix(Row, .ColIndex("RetainageRate")) = 0, "", .ValueMatrix(Row, .ColIndex("Retainage")))
            .TextMatrix(Row, .ColIndex("RetainageRate")) = IIf(.ValueMatrix(Row, .ColIndex("RetainageRate")) = 0, "", Round(.ValueMatrix(Row, .ColIndex("RetainageRate")), 2))
                        
            .Cell(flexcpForeColor, Row, .ColIndex("PreTax")) = IIf(.ValueMatrix(Row, .ColIndex("PreTax")) < -0.001, vbRed, vbWindowText)
            .Cell(flexcpForeColor, Row, .ColIndex("Tax")) = IIf(.ValueMatrix(Row, .ColIndex("Tax")) < -0.001, vbRed, vbWindowText)
            .Cell(flexcpForeColor, Row, .ColIndex("Retainage")) = IIf(.ValueMatrix(Row, .ColIndex("Retainage")) < -0.001, vbRed, vbWindowText)
                    
            If Not .RowHidden(Row) Then
                mTotalHBAmount = mTotalHBAmount + .ValueMatrix(Row, .ColIndex("Retainage"))
                If .ValueMatrix(Row, .ColIndex("TaxRate")) <> 0 Then
                    mTotalHBTax = mTotalHBTax + (Round(.ValueMatrix(Row, .ColIndex("Tax")) * .ValueMatrix(Row, .ColIndex("RetainageRate")) / .ValueMatrix(Row, .ColIndex("TaxRate")) * .ValueMatrix(Row, .ColIndex("RetainageRate")) / 100, 2))
                End If
                Pretax = Pretax + .ValueMatrix(Row, .ColIndex("pretax"))
                Tax = Tax + .ValueMatrix(Row, .ColIndex("tax"))
                
                If .TextMatrix(Row, .ColIndex("Job")) <> "" And .Cell(flexcpChecked, Row, .ColIndex("WrapInsuranceExempt")) = flexUnchecked Then
                    WrapInsurance = WrapInsurance + .ValueMatrix(Row, .ColIndex("pretax")) * mWrapInsuranceRate
                End If
                
            End If
        Next
    End With
    numWrapInsuranceAmount = WrapInsurance

    With gEnterTail
        .Cell(flexcpText, 0, 5) = Format(numAmount.Value - numTax.Value, "$#,##0.00")
        .Cell(flexcpText, 0, 6) = Format(numTax.Value, "$#,##0.00")
        .Cell(flexcpText, 0, 7) = Format(.Cell(flexcpValue, 0, 5) + .Cell(flexcpValue, 0, 6), "$#,##0.00")
        
        .Cell(flexcpText, 1, 5) = Format(Pretax, "$#,##0.00")
        .Cell(flexcpText, 1, 6) = Format(Tax, "$#,##0.00")
        .Cell(flexcpText, 1, 7) = Format(.Cell(flexcpValue, 1, 5) + .Cell(flexcpValue, 1, 6), "$#,##0.00")

        .Cell(flexcpText, 2, 5) = Format(numAmount.Value - numTax.Value - Pretax, "$#,##0.00")
        .Cell(flexcpText, 2, 6) = Format(numTax.Value - Tax, "$#,##0.00")
        .Cell(flexcpText, 2, 7) = Format(.Cell(flexcpValue, 2, 5) + .Cell(flexcpValue, 2, 6), "$#,##0.00")
        
        .Cell(flexcpForeColor, 0, 5) = IIf(Round(.Cell(flexcpValue, 0, 5), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, 0, 6) = IIf(Round(.Cell(flexcpValue, 0, 6), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, 0, 7) = IIf(Round(.Cell(flexcpValue, 0, 7), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, 1, 5) = IIf(Round(.Cell(flexcpValue, 1, 5), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, 1, 6) = IIf(Round(.Cell(flexcpValue, 1, 6), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, 1, 7) = IIf(Round(.Cell(flexcpValue, 1, 7), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, 2, 5) = IIf(Round(.Cell(flexcpValue, 2, 5), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, 2, 6) = IIf(Round(.Cell(flexcpValue, 2, 6), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpForeColor, 2, 7) = IIf(Round(.Cell(flexcpValue, 2, 7), 2) < 0, vbRed, vbWindowText)
        .Cell(flexcpFontBold, 2, 5) = False
        .Cell(flexcpFontBold, 2, 6) = False
        
    End With
    
    
    If App.Options(AutoCalcHeaderTotals) Then
        numAmount = Pretax + Tax
        numTax = Tax
    End If
End Sub

Private Property Get InvoiceRemaining() As Double
    InvoiceRemaining = Round(gEnterTail.ValueMatrix(2, 5), 2) + Round(gEnterTail.ValueMatrix(2, 6), 2)
End Property

Private Sub gEnterDist_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh

Static bInHere As Boolean
If bInHere Then Exit Sub
bInHere = True

    Dim bOverride As Boolean
    
    Dim bHasPO   As Boolean
    Dim bHasJob  As Boolean
    Dim bHasTaxGroup As Boolean
    Dim bDefReq  As Boolean
    Dim sDefault As String
    
    
    If ReadOnly Then
        Cancel = True
        bInHere = False
        Exit Sub
    End If
    bOverride = App.Options(AllowJobDrOverride) Or HFApp.Options(AccountingSystem) = asSimply
    
    
    With gEnterDist
    
        bHasPO = Trim(.TextMatrix(Row, .ColIndex("Commitment"))) <> ""
        bHasJob = Trim(.TextMatrix(Row, .ColIndex("Job"))) <> ""
        
        'dont trim this. when tax row is added to po with no tax the tax group is set to " "
        bHasTaxGroup = .TextMatrix(Row, .ColIndex("TaxGroup")) <> ""
        
'If .ColKey(Col) = "Phase" Then Stop
        bDefReq = .Cell(flexcpData, Row, Col) = "DefRequired" 'And (InvoiceRemaining <> 0 Or Row <> .Rows - 1)
        If Row <> .Rows - 1 Then .Cell(flexcpData, Row, Col) = ""
        
        sDefault = ""
        
        .ComboList = ""
        Select Case .ColKey(Col)
        
            Case "Commitment":
                .ComboList = "|..."
            
            Case "CommitmentItem":
                .ComboList = "|..."
                Cancel = Not bHasPO
                
            Case "Job":
                .ComboList = "|..."
                Cancel = bHasPO:
                sDefault = IIf(Row > 1 And .TextMatrix(Row - 1, Col) <> "", .TextMatrix(Row - 1, Col), txtJob.Text)
                
            Case "Extra":
                .ComboList = "|..."
                Cancel = bHasPO Or Not bHasJob:
                sDefault = IIf(Row > 1 And .TextMatrix(Row - 1, Col) <> "", .TextMatrix(Row - 1, Col), "")
                
            Case "Phase":
                .ComboList = "|..."
                Cancel = bHasPO Or Not bHasJob:
                sDefault = IIf(Row > 1 And .TextMatrix(Row - 1, Col) <> "", .TextMatrix(Row - 1, Col), mVDefPhase)
            
            Case "Category"
                .ComboList = "|..."
                Cancel = bHasPO Or Not bHasJob:
                sDefault = IIf(Row > 1 And .TextMatrix(Row - 1, Col) <> "", .TextMatrix(Row - 1, Col), mVDefCategory)
                
            Case "CategoryDesc", "DebitAccoutDesc", "PhaseDesc":
                .ComboList = "..."
                
            Case "DebitAccount":
                .ComboList = "|..."
                Cancel = bHasJob And Not bOverride:
                If .TextMatrix(Row, Col) = "" Then
                    sDefault = IIf(Row > 1, .TextMatrix(Row - 1, Col), "")
                    If sDefault = "" Then sDefault = mVDefDebitAccount
                    If sDefault = "" Then sDefault = App.Options(DefaultGLAccount)
                End If
                 
            
            Case "Equipment":
                .ComboList = "|..."
            Case "EQCostCode":
                .ComboList = "|..."
            
            Case "InvoicedQuantity":
            
            Case "InvoicedUnitPrice":
            
            Case "Billable"
                Cancel = .TextMatrix(Row, .ColIndex("Job")) = ""
                
            Case "TaxGroup"
                .ComboList = "|..."
                If bHasPO Then
                    If bHasTaxGroup Then
                        Cancel = True
                    Else
                        sDefault = IIf(Row > 1 And .TextMatrix(Row - 1, Col) <> "", .TextMatrix(Row - 1, Col), mVDefTaxGroup)
                        If sDefault = "" Then sDefault = App.Options(DefaultTaxGroup)
                    End If
                Else
                
                    's = "select dbo.purch_getdefaulttaxgroup(" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Job"))) & ",'','','','','',''," & DbQuote(Str, txtVendor.tag) & "," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Category"))) & "," & DbQuote(Num, HFApp.DivisionID) & ")"
                    'sDefault = IIf(Row > 1 And .TextMatrix(Row - 1, Col) <> "", .TextMatrix(Row - 1, Col), mVDefTaxGroup)
                
                    On Error Resume Next
                    sDefault = ""
                    sDefault = "" & HFApp.SqlExec("select dbo.purch_getdefaulttaxgroup(" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Job"))) & ",'','','','','',''," & DbQuote(Str, txtVendor.tag) & "," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Category"))) & "," & DbQuote(num, HFApp.DivisionID) & ")")(0)
                    On Error GoTo eh
                    
                    If sDefault = "" Then sDefault = IIf(Row > 1 And .TextMatrix(Row - 1, Col) <> "", .TextMatrix(Row - 1, Col), mVDefTaxGroup)
                    If sDefault = "" Then sDefault = App.Options(DefaultTaxGroup)
                End If
                
            Case "PreTax"
                If .TextMatrix(Row, .ColIndex("InvoicedQuantity")) = "" Then
                    sDefault = Format(InvoiceRemaining / (1 + .ValueMatrix(Row, .ColIndex("TaxRate")) / 100), "#,##0.00")
                Else
                    sDefault = Format(.ValueMatrix(Row, .ColIndex("InvoicedQuantity")) * .ValueMatrix(Row, .ColIndex("InvoicedUnitPrice")), "#,##0.00")
                End If
                
            Case "Tax"
                Cancel = (Not bHasTaxGroup)
                
            Case "RetainageRate", "Retainage"
                Cancel = mIsHBInvoice
            
                
            Case "Description"
                If HFApp.Options(AccountingSystem) = asTimberline Then
                    .EditMaxLength = 30
                End If
                sDefault = IIf(App.Options(DefaultDescription), txtDescription.Text, "")
                
        End Select
        
        
        If bDefReq And sDefault <> "" And Not Cancel And .ComboList <> "..." Then
            .Text = sDefault
            Call gEnterDist_AfterEdit(Row, Col)
        End If
        
        If mReadOnly Then
            .ComboList = ""
            Cancel = True
        End If
    End With
    
    bInHere = False
    Exit Sub
eh: Call ErrHandler(SRCFILE & "gEnterDist_BeforeEdit")
    bInHere = False
End Sub

Private Sub gEnterDist_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo eh
    Dim s           As String
    Dim h           As String
    Dim a           As String
    Dim hJob        As String 'values from invoice header
    Dim hVendor     As String
    Dim sCommitment As String 'values from current row
    Dim sJob        As String
    Dim sExtra      As String
    Dim sPhase      As String
    Dim sTaxGroup   As String
    Dim sHiddenCols As String
    
    With gEnterDist
        
        'make sure we are not in edit mode
        .Editable = flexEDNone
        .Editable = flexEDKbdMouse
        '------------------
        
        
        hJob = txtJob.Text
        hVendor = txtVendor.tag
        
        sTaxGroup = .TextMatrix(Row, .ColIndex("TaxGroup"))
        sCommitment = .TextMatrix(Row, .ColIndex("Commitment"))
        
        sJob = .TextMatrix(Row, .ColIndex("Job"))
        sExtra = .TextMatrix(Row, .ColIndex("Extra"))
        sPhase = .TextMatrix(Row, .ColIndex("Phase"))
        sHiddenCols = ""
        
        Select Case .ColKey(Col)
            Case "Commitment"
                a = SelectPO(hVendor, hJob)
                If hJob <> "" Then
                    a = App.Options(Caption_Job) & " " & App.Options(Caption_Commitment) & "s" & Chr(1) & SelectPO(, hJob) & Chr(0) & _
                        App.Options(Caption_Vendor) & " " & App.Options(Caption_Commitment) & "s" & Chr(1) & SelectPO(hVendor)
                End If
                                   
            Case "CommitmentItem":   a = SelectCommitmentItem(sCommitment)
            Case "Job":              a = SelectJob()
            Case "Extra":            a = SelectExtra(sJob)
            
            
            Case "Equipment"
                h = "select Equipment,Description from equipment where divisionid=" & DbQuote(num, HFApp.DivisionID)
            
            Case "EQCostCode"
                h = "select CostCode,Description from standardeqcostcodes where divisionid=" & DbQuote(num, HFApp.DivisionID)
            
            Case "Phase"
                a = App.Options(Caption_Job) & " " & App.Options(Caption_Phase) & "s" & Chr(1) & SelectJobExtraCostCode(sJob, sExtra) & Chr(0) & _
                    "Standard " & App.Options(Caption_Phase) & "s" & Chr(1) & SelectJobExtraCostCode()
                
            Case "PhaseDesc"
                Col = .ColIndex("Phase")
                If Not HFApp.Options(AccountingSystem) = asTimberline Then sHiddenCols = App.Options(Caption_Phase)
                a = App.Options(Caption_Job) & " " & App.Options(Caption_Phase) & "s" & Chr(1) & SelectJobExtraCostCode(sJob, sExtra) & Chr(0) & _
                    "Standard " & App.Options(Caption_Phase) & "s" & Chr(1) & SelectJobExtraCostCode()
                
            Case "Category"
                a = App.Options(Caption_Phase) & " Categories" & Chr(1) & SelectJobExtraCostCodeCategory(sJob, sExtra, sPhase) & Chr(0) & _
                    "Standard Categories" & Chr(1) & SelectJobExtraCostCodeCategory()
                                       
            Case "CategoryDesc"
                Col = .ColIndex("Category")
                If Not HFApp.Options(AccountingSystem) = asTimberline Then sHiddenCols = App.Options(Caption_Category)
                a = App.Options(Caption_Phase) & " Categories" & Chr(1) & SelectJobExtraCostCodeCategory(sJob, sExtra, sPhase) & Chr(0) & _
                    "Standard Categories" & Chr(1) & SelectJobExtraCostCodeCategory()
            
            Case "DebitAccount", "DebitAccountDesc"
                Col = .ColIndex("DebitAccount")
                a = SelectGLAccount()
            
            Case "TaxGroup":
                If HFApp.Options(AccountingSystem) <> asTimberline Then sHiddenCols = App.Options(Caption_Retainage)
                h = "SELECT taxgroup " & Quote(App.Options(Caption_TaxGroup)) & ",Description,groupRate TaxRate,retainagerate " & Quote(App.Options(Caption_Retainage)) & " FROM taxgroups where DivisionID =" & HFApp.DivisionID
            
        End Select
        
        
        If mCancelEdit Then
            Call .FinishEditing(True)
            .Text = ""
        End If
        
        
        If h <> "" Then
            If FPickList.Choose(HFApp.Databases(dbHomeFront), .TextMatrix(0, Col), h, .Text, , , , sHiddenCols) Then
                s = FPickList.SelectedItem(1)
                .Text = s
                Call gEnterDist_AfterEdit(Row, Col)
            End If
        ElseIf a <> "" Then
            If FPickList.Choose(HFApp.Databases(AccountingDB), .TextMatrix(0, Col), a, .Text, , , , sHiddenCols) Then
                s = Trim(FPickList.SelectedItem(1))
                .Cell(flexcpText, .Row, Col, .RowSel, Col) = s
                Call gEnterDist_AfterEdit(.Row, Col)
            End If
        End If
        
    
    End With
    Exit Sub
eh: Call ErrHandler(SRCFILE & "gEnterDist_CellButtonClick")
End Sub

Private Sub AddTaxRow(Row As Long, TaxGroup As String, TaxGroupDesc As String, TaxRate As Double)
    Dim i As Long
    With gEnterDist
        i = .Row + 1
        .AddItem "", i
        .TextMatrix(i, .ColIndex("CommitmentVendor")) = ""
        .TextMatrix(i, .ColIndex("Commitment")) = ""
        .TextMatrix(i, .ColIndex("CommitmentItem")) = ""
        .TextMatrix(i, .ColIndex("Job")) = .TextMatrix(Row, .ColIndex("Job"))
        .TextMatrix(i, .ColIndex("JobDesc")) = .TextMatrix(Row, .ColIndex("JobDesc"))
        .TextMatrix(i, .ColIndex("Extra")) = .TextMatrix(Row, .ColIndex("Extra"))
        .TextMatrix(i, .ColIndex("ExtraDesc")) = .TextMatrix(Row, .ColIndex("ExtraDesc"))
        .TextMatrix(i, .ColIndex("Phase")) = .TextMatrix(Row, .ColIndex("Phase"))
        .TextMatrix(i, .ColIndex("PhaseDesc")) = .TextMatrix(Row, .ColIndex("PhaseDesc"))
        .TextMatrix(i, .ColIndex("Category")) = .TextMatrix(Row, .ColIndex("Category"))
        .TextMatrix(i, .ColIndex("CategoryDesc")) = .TextMatrix(Row, .ColIndex("CategoryDesc"))
        .TextMatrix(i, .ColIndex("TaxGroup")) = TaxGroup
        .TextMatrix(i, .ColIndex("TaxGroupDesc")) = TaxGroupDesc
        .TextMatrix(i, .ColIndex("TaxRate")) = TaxRate
        .TextMatrix(i, .ColIndex("CommittedQuantity")) = ""
        .TextMatrix(i, .ColIndex("CommittedUnitPrice")) = ""
        .TextMatrix(i, .ColIndex("InvoicedQuantity")) = ""
        .TextMatrix(i, .ColIndex("InvoicedUnitPrice")) = ""
        
        
        
        .TextMatrix(i, .ColIndex("PreTax")) = 0
        .TextMatrix(i, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("Pretax")) * TaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("RetainageRate")) = .TextMatrix(Row, .ColIndex("RetainageRate"))
        .TextMatrix(i, .ColIndex("Description")) = "tax"
        
        .Cell(flexcpChecked, i, .ColIndex("Billable")) = IIf(mAPInvoicesDefaultToBillable, flexChecked, flexUnchecked)
        
        Call RecalcRetainage(i)
    End With
End Sub


Public Function Quote(s As String, Optional RemoveFormatting As Boolean = True, Optional Length As Long, Optional RemoveSpecial As Boolean = True) As String
    
    If RemoveFormatting Then
        s = Replace(s, vbTab, " ")
        s = Replace(s, vbLf, " ")
        s = Replace(s, vbCr, " ")
    End If
        
    'replace special characters
    If RemoveSpecial Then s = RemoveSpecialChar(s)
    
    'replace quote with 2 apostrophes
    s = Replace(s, vbQuote, "''")

    If Length <> 0 Then
        s = left(s, Length)
    End If
    
    Quote = vbQuote & s & vbQuote
    
End Function

Private Function RemoveSpecialChar(Text As String) As String
    Dim i As Long
    Dim j As Long
    Dim Char As String

    i = 1
    For j = 1 To Len(Text)
        Char = Mid$(Text, j, 1)
        If (AscW(Char) And &HFFFF&) <= &H7F& Then
            Mid$(Text, i, 1) = Char
            i = i + 1
        End If
    Next
    
    RemoveSpecialChar = left$(Text, i - 1)
End Function




Private Sub gEnterDist_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo eh
    
'lookup taxgroup using the lookup function
'select [dbo].[Purch_GetDefaultTaxGroup](@Job,@Community,'','','','','',@Vendor,@JCCategory,@DivisionID)

    
    Dim s             As String
    Dim hJob          As String 'values from invoice header
    Dim hVendor       As String
    Dim sCommitment   As String 'values from current row
    Dim sCommitmentItem       As Long
    Dim sJob          As String
    Dim sExtra        As String
    Dim sPhase        As String
    Dim sCategory     As String
    Dim sDebitAccount As String
    Dim sTaxGroup     As String
    Dim sEquipment    As String
    Dim sEQCostCode   As String
    Dim rs            As Recordset
    
    Dim Amount  As Double
    Dim Pretax  As Double
    Dim Tax     As Double
    Dim TaxRate As Double
    
    Dim AddCCOK  As Boolean
    Dim AddCatOK As Boolean
    
    If mLoading Then Exit Sub
    
    With gEnterDist
        For Row = Min(.Row, .RowSel) To Max(.Row, .RowSel)
        mCancelEdit = False
        
        'get all these values
        hJob = Trim(txtJob.Text)
        hVendor = txtVendor.tag
        sCommitment = .TextMatrix(Row, .ColIndex("Commitment"))
        sCommitmentItem = .ValueMatrix(Row, .ColIndex("CommitmentItem"))
        sJob = .TextMatrix(Row, .ColIndex("Job"))
        sExtra = .TextMatrix(Row, .ColIndex("Extra"))
        sPhase = .TextMatrix(Row, .ColIndex("Phase"))
        sCategory = .TextMatrix(Row, .ColIndex("Category"))
        sDebitAccount = .TextMatrix(Row, .ColIndex("DebitAccount"))
        sTaxGroup = .TextMatrix(Row, .ColIndex("TaxGroup"))
        sEquipment = .TextMatrix(Row, .ColIndex("Equipment"))
        sEQCostCode = .TextMatrix(Row, .ColIndex("EqCostCode"))

        
        Select Case True
            Case .ColKey(Col) = "WrapInsuranceExempt"
                Call RecalcTotals
                
                
            Case .ColKey(Col) = "Commitment"
                If sCommitment = "" Then
                    .TextMatrix(Row, .ColIndex("CommitmentVendor")) = ""
                    .TextMatrix(Row, .ColIndex("Commitment")) = ""
                    .TextMatrix(Row, .ColIndex("CommitmentItem")) = ""
                    .TextMatrix(Row, .ColIndex("Job")) = ""
                    .TextMatrix(Row, .ColIndex("JobDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Extra")) = ""
                    .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Phase")) = ""
                    .TextMatrix(Row, .ColIndex("PhaseDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                    .TextMatrix(Row, .ColIndex("DebitAccount")) = ""
                    .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Description")) = ""
                    .TextMatrix(Row, .ColIndex("WrapInsuranceExempt")) = ""
                Else
                    Set rs = HFApp.SqlExec(ValidateCommitment(sCommitment, hVendor), AccountingDB)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found." & vbCrLf & "It or the " & App.Options(Caption_Job) & " may have been closed.", vbExclamation, App.ProductName
                    Else
                        If Not SchedulingIsFinished(.TextMatrix(Row, .ColIndex("Commitment"))) Then
                            mCancelEdit = True
                        Else
                            If rs("CrossPay") And Not App.Options(AllowCrossPayingPOs) Then
                                MsgBox "Vendor " & txtVendor.Text & " cannot invoice " & .TextMatrix(0, Col) & " " & .TextMatrix(Row, Col) & vbCrLf & .TextMatrix(0, Col) & " " & .TextMatrix(Row, Col) & " was issued to vendor " & Trim("" & rs("vendor")), vbExclamation, App.ProductName
                                mCancelEdit = True
                            Else
                                .TextMatrix(Row, .ColIndex("CommitmentVendor")) = Trim("" & rs("vendor"))
                                .TextMatrix(Row, .ColIndex("Commitment")) = Trim("" & rs("ponumber"))
                                .TextMatrix(Row, .ColIndex("CommitmentItem")) = ""
                                .TextMatrix(Row, .ColIndex("Job")) = ""
                                .TextMatrix(Row, .ColIndex("JobDesc")) = ""
                                .TextMatrix(Row, .ColIndex("Extra")) = ""
                                .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                                .TextMatrix(Row, .ColIndex("Phase")) = ""
                                .TextMatrix(Row, .ColIndex("PhaseDesc")) = ""
                                .TextMatrix(Row, .ColIndex("Category")) = ""
                                .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                                .TextMatrix(Row, .ColIndex("DebitAccount")) = ""
                                .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = ""
                                .TextMatrix(Row, .ColIndex("Description")) = ""
                                .TextMatrix(Row, .ColIndex("WrapInsuranceExempt")) = "" & rs("WrapInsuranceExempt")
                            End If
                        End If
                    End If
                End If
                Call RecalcTotals
                
                
                
            Case .ColKey(Col) = "CommitmentItem"
                If sCommitmentItem = "0" Then
                    .TextMatrix(Row, .ColIndex("CommitmentItem")) = ""
                    .TextMatrix(Row, .ColIndex("Job")) = ""
                    .TextMatrix(Row, .ColIndex("JobDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Extra")) = ""
                    .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Phase")) = ""
                    .TextMatrix(Row, .ColIndex("PhaseDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Description")) = ""
                    .TextMatrix(Row, .ColIndex("TaxGroup")) = ""
                    .TextMatrix(Row, .ColIndex("TaxGroupDesc")) = ""
                    .TextMatrix(Row, .ColIndex("TaxRate")) = ""
                    
                    .TextMatrix(Row, .ColIndex("CommittedQuantity")) = ""
                    .TextMatrix(Row, .ColIndex("CommittedUnitPrice")) = ""
                    .TextMatrix(Row, .ColIndex("InvoicedQuantity")) = ""
                    .TextMatrix(Row, .ColIndex("InvoicedUnitPrice")) = ""

                    .TextMatrix(Row, .ColIndex("PreTax")) = ""
                    .TextMatrix(Row, .ColIndex("Tax")) = ""
                    .TextMatrix(Row, .ColIndex("Retainage")) = ""
                    .TextMatrix(Row, .ColIndex("RetainageRate")) = ""
                    .TextMatrix(Row, .ColIndex("Description")) = ""
                    .TextMatrix(Row, .ColIndex("DebitAccount")) = ""
                    .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = ""
                Else
                    Set rs = HFApp.SqlExec(ValidateCommitmentItem(sCommitment, sCommitmentItem), AccountingDB)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found", vbExclamation, App.ProductName
                        
                    Else
                        .TextMatrix(Row, .ColIndex("CommitmentVendor")) = Trim("" & rs("CommitmentVendor"))
                        .TextMatrix(Row, .ColIndex("Commitment")) = Trim("" & rs("Commitment"))
                        .TextMatrix(Row, .ColIndex("CommitmentItem")) = Trim("" & rs("CommitmentItem"))
                        .TextMatrix(Row, .ColIndex("Job")) = Trim("" & rs("Job"))
                        .TextMatrix(Row, .ColIndex("JobDesc")) = Trim("" & rs("JobDesc"))
                        .TextMatrix(Row, .ColIndex("Extra")) = Trim("" & rs("Extra"))
                        .TextMatrix(Row, .ColIndex("ExtraDesc")) = Trim("" & rs("ExtraDesc"))
                        .TextMatrix(Row, .ColIndex("Phase")) = Trim("" & rs("Phase"))
                        .TextMatrix(Row, .ColIndex("PhaseDesc")) = Trim("" & rs("PhaseDesc"))
                        .TextMatrix(Row, .ColIndex("Category")) = Trim("" & rs("Category"))
                        .TextMatrix(Row, .ColIndex("CategoryDesc")) = Trim("" & rs("CategoryDesc"))
                        .TextMatrix(Row, .ColIndex("TaxGroup")) = Trim("" & rs("TaxGroup"))
                        .TextMatrix(Row, .ColIndex("TaxGroupDesc")) = Trim("" & rs("TaxGroupDesc"))
                        .TextMatrix(Row, .ColIndex("CommittedQuantity")) = Val("" & rs("Quantity"))
                        .TextMatrix(Row, .ColIndex("CommittedUnitPrice")) = Val("" & rs("UnitPrice"))
                        .TextMatrix(Row, .ColIndex("InvoicedQuantity")) = Val("" & rs("RemainingQuantity"))
                        .TextMatrix(Row, .ColIndex("InvoicedUnitPrice")) = .ValueMatrix(Row, .ColIndex("CommittedUnitPrice"))
                        .TextMatrix(Row, .ColIndex("TaxRate")) = Val("" & rs("TaxRate"))
                        .TextMatrix(Row, .ColIndex("RetainageRate")) = Val("" & rs("RetainageRate"))
                        .TextMatrix(Row, .ColIndex("DebitAccount")) = Trim("" & rs("DebitAccount"))
                        .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = Trim("" & rs("DebitAccountDesc"))
                        
                        ' Figure out the amount
                        TaxRate = .ValueMatrix(Row, .ColIndex("TaxRate"))
                        If App.Options(ShowQtyAndUnitPrice) Then
                            Pretax = Round(.ValueMatrix(Row, .ColIndex("InvoicedQuantity")) * .ValueMatrix(Row, .ColIndex("InvoicedUnitPrice")), 2)
                            Tax = Round(Pretax * TaxRate / 100, 2)
                        Else
                            Amount = CommitmentItemRemaining(.TextMatrix(Row, .ColIndex("Commitment")), .TextMatrix(Row, .ColIndex("CommitmentItem")))
                            
                            If Not App.Options(AutoCalcHeaderTotals) Then
                            
                                ' if this is a credit memo
                                If InvoiceRemaining < 0 Then
                                    'need to reverse the logic
                                    If Amount < InvoiceRemaining Then Amount = InvoiceRemaining
                                Else
                                    'do normal comparison
                                    If Amount > InvoiceRemaining Then Amount = InvoiceRemaining
                                End If
                            End If
                            Pretax = Round(Amount / (TaxRate / 100 + 1), 2)
                            Tax = Round(Amount - Pretax, 2)
                        End If
                        
                        .TextMatrix(Row, .ColIndex("AmountRemaining")) = Pretax
                        .TextMatrix(Row, .ColIndex("PreTax")) = Pretax
                        .TextMatrix(Row, .ColIndex("Tax")) = Tax
                        .TextMatrix(Row, .ColIndex("Description")) = Trim("" & rs("CommitmentItemDesc"))
                        .Cell(flexcpData, .Row, 0, .Row, .Cols - 1) = ""
                    End If
                End If
                Call RecalcRetainage(Row)
                Call RecalcTotals
                Call VerifyAmounts(Row)
            
            Case .ColKey(Col) = "Job"
                If sJob = "" Then
                    .TextMatrix(Row, .ColIndex("Job")) = ""
                    .TextMatrix(Row, .ColIndex("JobDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Extra")) = ""
                    .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Phase")) = ""
                    .TextMatrix(Row, .ColIndex("PhaseDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
   '                 .TextMatrix(Row, .ColIndex("DebitAccount")) = ""
   '                 .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = ""
                    .TextMatrix(Row, .ColIndex("WrapInsuranceExempt")) = ""
                Else
                    If sJob <> "" Then
                       Set rs = ValidateJob(sJob)
                       If rs.EOF Then
                           mCancelEdit = True
                           MsgBox .TextMatrix(0, Col) & " not found." & vbCrLf & "It may have been closed.", vbExclamation, App.ProductName
                       Else
                       
                            If "" & rs("Status") = "Closed" And sJob <> txtJob.Text Then
                                If MsgBox("Job is closed.", vbInformation & vbOKCancel, App.ProductName) = vbCancel Then
                                    Exit Sub
                                End If
                            End If

                           .TextMatrix(Row, .ColIndex("Job")) = Trim("" & rs("job"))
                           .TextMatrix(Row, .ColIndex("JobDesc")) = Trim("" & rs("description"))
                           .TextMatrix(Row, .ColIndex("Extra")) = ""
                           .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                           .TextMatrix(Row, .ColIndex("Phase")) = ""
                           .TextMatrix(Row, .ColIndex("PhaseDesc")) = ""
                           .TextMatrix(Row, .ColIndex("Category")) = ""
                           .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
    '                       .TextMatrix(Row, .ColIndex("DebitAccount")) = ""
    '                       .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = ""
                           .TextMatrix(Row, .ColIndex("WrapInsuranceExempt")) = "" & rs("WrapInsuranceExempt")
                           
                       End If
                    End If
                End If
                Call RecalcTotals
                
            Case .ColKey(Col) = "Extra"
                If sExtra = "" Then
                    .TextMatrix(Row, .ColIndex("Extra")) = ""
                    .TextMatrix(Row, .ColIndex("ExtraDesc")) = ""
                Else
                    Set rs = HFApp.SqlExec(ValidateJobExtra(sJob, sExtra), AccountingDB)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found", vbExclamation, App.ProductName
                        
                    Else
                        .TextMatrix(Row, .ColIndex("Extra")) = Trim("" & rs("extra"))
                        .TextMatrix(Row, .ColIndex("ExtraDesc")) = Trim("" & rs("description"))
                        .TextMatrix(Row, .ColIndex("Phase")) = ""
                        .TextMatrix(Row, .ColIndex("PhaseDesc")) = ""
                        .TextMatrix(Row, .ColIndex("Category")) = ""
                        .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
'                        .TextMatrix(Row, .ColIndex("DebitAccount")) = ""
'                        .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = ""
                    End If
                End If
            
            Case .ColKey(Col) = "Phase"
                If sPhase = "" Then
                    .TextMatrix(Row, .ColIndex("Phase")) = ""
                    .TextMatrix(Row, .ColIndex("PhaseDesc")) = ""
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
  '                  .TextMatrix(Row, .ColIndex("DebitAccount")) = ""
  '                  .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = ""
                Else
                    
                    If CodeOnInvoice(sJob, sExtra, sPhase) Then
                        s = ""
                        s = s & "select 1,s.costcode,s.description,a.account debitaccount,a.description debitaccountdesc" & vbCrLf
                        s = s & "  from standardcostcodes s" & vbCrLf
                        s = s & "  left outer join glaccounts a on(s.DivisionID = a.DivisionID and s.debitaccount=a.account)"
                        s = s & " where s.DivisionID = " & HFApp.DivisionID & " And CostCode = " & DbQuote(Str, sPhase) & vbCrLf
                        Set rs = HFApp.SqlExec(s, dbHomeFront)
                    Else
                        s = ValidateJobExtraPhase(sJob, sExtra, sPhase)
                        Set rs = HFApp.SqlExec(s, AccountingDB)
                    End If
                    
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found", vbExclamation, App.ProductName
                    Else
                        If Val("" & rs(0)) = 1 Then
                            'found in job cost codes
                            .TextMatrix(Row, .ColIndex("Phase")) = Trim("" & rs("costcode"))
                            .TextMatrix(Row, .ColIndex("PhaseDesc")) = Trim("" & rs("description"))
                            .TextMatrix(Row, .ColIndex("Category")) = ""
                            .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                            .TextMatrix(Row, .ColIndex("DebitAccount")) = Trim("" & rs("DebitAccount"))
                            .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = Trim("" & rs("DebitAccountDesc"))
                        Else
                            'not found in job cost codes
                            If AddCCOK Then
                                .TextMatrix(Row, .ColIndex("Phase")) = Trim("" & rs("costcode"))
                                .TextMatrix(Row, .ColIndex("PhaseDesc")) = Trim("" & rs("description"))
                                .TextMatrix(Row, .ColIndex("Category")) = ""
                                .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                                .TextMatrix(Row, .ColIndex("DebitAccount")) = Trim("" & rs("DebitAccount"))
                                .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = Trim("" & rs("DebitAccountDesc"))
                            ElseIf MsgBox(vbQuote & FullTrim("" & rs("description")) & vbQuote & " is not setup for job """ & FullTrim("" & .TextMatrix(Row, .ColIndex("Job"))) & " - " & FullTrim(.TextMatrix(Row, .ColIndex("JobDesc"))) & """" & vbCrLf & vbCrLf & "Do you want to add this cost code?", vbQuestion + vbYesNo) = vbYes Then
                                AddCCOK = True
                                .TextMatrix(Row, .ColIndex("Phase")) = Trim("" & rs("costcode"))
                                .TextMatrix(Row, .ColIndex("PhaseDesc")) = Trim("" & rs("description"))
                                .TextMatrix(Row, .ColIndex("Category")) = ""
                                .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                                .TextMatrix(Row, .ColIndex("DebitAccount")) = Trim("" & rs("DebitAccount"))
                                .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = Trim("" & rs("DebitAccountDesc"))
                            Else
                                mCancelEdit = True
                            End If
                        End If
                    End If
                End If
                
            Case .ColKey(Col) = "Category"
                If sCategory = "" Then
                    .TextMatrix(Row, .ColIndex("Category")) = ""
                    .TextMatrix(Row, .ColIndex("CategoryDesc")) = ""
                Else
                    If CodeOnInvoice(sJob, sExtra, sPhase, sCategory) Then
                        s = ""
                        s = s & "select 1,ca.category,ca.description,gl.account DebitAccount,gl.description DebitAccountDesc" & vbCrLf
                        s = s & "from standardcategories ca" & vbCrLf
                        s = s & "left outer join standardcostcodes cc on(ca.divisionid=cc.divisionid and cc.costcode=" & DbQuote(Str, sPhase) & ")" & vbCrLf
                        s = s & "left outer join glaccounts gl on(ca.DivisionID=gl.DivisionID and gl.account=isnull(nullif(cc.debitaccount,''),ca.debitaccount))" & vbCrLf
                        s = s & "where ca.DivisionID = " & HFApp.DivisionID & " and ca.category=" & DbQuote(Str, sCategory) & vbCrLf
                        Set rs = HFApp.SqlExec(s, dbHomeFront)
                    Else
                        Set rs = HFApp.SqlExec(ValidateJobExtraPhaseCategory(sJob, sExtra, sPhase, sCategory), AccountingDB)
                    End If
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found", vbExclamation, App.ProductName
                        .Text = ""
                    Else
                        .TextMatrix(Row, .ColIndex("DebitAccount")) = Trim("" & rs("DebitAccount"))
                        .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = Trim("" & rs("DebitAccountDesc"))
                        If Val("" & rs(0)) = 1 Then
                            'found in job categories
                            .TextMatrix(Row, .ColIndex("Category")) = Trim("" & rs("category"))
                            .TextMatrix(Row, .ColIndex("CategoryDesc")) = Trim("" & rs("description"))
                        Else
                            'not found in job categories
                            If AddCatOK Then
                                .TextMatrix(Row, .ColIndex("Category")) = Trim("" & rs("category"))
                                .TextMatrix(Row, .ColIndex("CategoryDesc")) = Trim("" & rs("description"))
                            ElseIf MsgBox(vbQuote & FullTrim("" & rs("description")) & """ is not setup for """ & FullTrim(.TextMatrix(Row, .ColIndex("PhaseDesc"))) & """" & vbCrLf & vbCrLf & "Do you want to add this category?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                                AddCatOK = True
                                .TextMatrix(Row, .ColIndex("Category")) = Trim("" & rs("category"))
                                .TextMatrix(Row, .ColIndex("CategoryDesc")) = Trim("" & rs("description"))
                            Else
                                mCancelEdit = True
                                
                            End If
                        End If
                    End If
                    If Not mCancelEdit Then
                        Call VerifyAmounts(Row)
                    End If
                    
                End If
                                       
            Case .ColKey(Col) = "DebitAccount"
                If sDebitAccount = "" Then
                    .TextMatrix(Row, .ColIndex("DebitAccount")) = ""
                    .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = ""
                Else
                    If HFApp.Options(AccountingSystem) = asTimberline And Not App.Options(AllowJobDrOverride) Then
                        .TextMatrix(Row, .ColIndex("Job")) = ""
                        .TextMatrix(Row, .ColIndex("Extra")) = ""
                        .TextMatrix(Row, .ColIndex("Phase")) = ""
                        .TextMatrix(Row, .ColIndex("Category")) = ""
                    End If
                    
                    Set rs = HFApp.SqlExec(ValidateGlAccount(sDebitAccount), AccountingDB)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found", vbExclamation, App.ProductName
                        .Text = ""
                    Else
                        .TextMatrix(Row, .ColIndex("DebitAccount")) = Trim("" & rs("account"))
                        .TextMatrix(Row, .ColIndex("DebitAccountDesc")) = Trim("" & rs("description"))
                    End If
                    
                End If
                
                
            Case .ColKey(Col) = "Equipment"
                If sEquipment = "" Then
                    .TextMatrix(Row, .ColIndex("EquipmentDesc")) = ""
                Else
                    s = "select equipment,description FROM equipment WHERE DivisionID =" & HFApp.DivisionID & " and equipment=" & DbQuote(Str, sEquipment)
                    Set rs = HFApp.SqlExec(s, dbHomeFront)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found", vbExclamation, App.ProductName
                        .Text = ""
                    Else
                        .TextMatrix(Row, .ColIndex("EquipmentDesc")) = Trim("" & rs("description"))
                    End If
                End If
                
            Case .ColKey(Col) = "EQCostCode"
                If sEQCostCode = "" Then
                    .TextMatrix(Row, .ColIndex("EqCostCodeDesc")) = ""
                Else
                    s = "select Costcode,description FROM standardeqcostcodes WHERE DivisionID =" & HFApp.DivisionID & " and costcode=" & DbQuote(Str, sEQCostCode)
                    Set rs = HFApp.SqlExec(s, dbHomeFront)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found", vbExclamation, App.ProductName
                        .Text = ""
                    Else
                        .TextMatrix(Row, .ColIndex("EqCostCodeDesc")) = Trim("" & rs("description"))
                    End If
                End If
                
            Case .ColKey(Col) = "TaxGroup"
                If sTaxGroup = "" Then
                    .TextMatrix(Row, .ColIndex("TaxGroup")) = ""
                    .TextMatrix(Row, .ColIndex("TaxGroupDesc")) = ""
                    .TextMatrix(Row, .ColIndex("TaxRate")) = 0
                    .TextMatrix(Row, .ColIndex("Tax")) = 0
                Else
                    s = "select taxgroup,description,grouprate,RetainageRate FROM taxgroups where DivisionID =" & HFApp.DivisionID & " and taxgroup=" & DbQuote(Str, sTaxGroup)
                    Set rs = HFApp.SqlExec(s, dbHomeFront)
                    If rs.EOF Then
                        mCancelEdit = True
                        MsgBox .TextMatrix(0, Col) & " not found", vbExclamation, App.ProductName
                        .Text = ""
                    Else
                        If sCommitment = "" Then
                            .TextMatrix(Row, .ColIndex("TaxGroup")) = Trim("" & rs("taxgroup"))
                            .TextMatrix(Row, .ColIndex("TaxGroupDesc")) = Trim("" & rs("description"))
                            .TextMatrix(Row, .ColIndex("TaxRate")) = Val("" & rs("grouprate"))
                            .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("PreTax")) * Val("" & rs("grouprate")) / 100, 2)
                            Call RecalcRetainage(Row)
                            Call RecalcTotals
                        Else
                            .TextMatrix(Row, .ColIndex("TaxGroup")) = " "
                            Call AddTaxRow(Row, Trim("" & rs("taxgroup")), Trim("" & rs("description")), Val("" & rs("grouprate")))
                            Call RecalcRetainage(Row)
                            Call RecalcTotals
                        End If
                    End If
                End If
                
                
            Case .ColKey(Col) = "InvoicedQuantity"
'                'check if over commited amount
'                'If .ValueMatrix(Row, .ColIndex("InvoicedQuantity")) * .ValueMatrix(Row, .ColIndex("InvoicedUnitPrice")) > .ValueMatrix(Row, .ColIndex("AmountRemaining")) Then
'                If ExceedsCommittedAmount() Then
'                   mCancelEdit = True
'                    MsgBox "Exceeds amount committed.", vbExclamation, App.ProductName
'                Else
                    .TextMatrix(Row, .ColIndex("InvoicedQuantity")) = .ValueMatrix(Row, .ColIndex("InvoicedQuantity"))
                    .TextMatrix(Row, .ColIndex("PreTax")) = .ValueMatrix(Row, .ColIndex("InvoicedQuantity")) * .ValueMatrix(Row, .ColIndex("InvoicedUnitPrice"))
                    .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("PreTax")) * .ValueMatrix(Row, .ColIndex("TaxRate")) / 100, 2)
                    Call RecalcRetainage(Row)
                    Call RecalcTotals
                    Call VerifyAmounts(Row)
'                End If
                
            
            Case .ColKey(Col) = "InvoicedUnitPrice"
                'check if over commited price
                If .ValueMatrix(Row, Col) > .ValueMatrix(Row, .ColIndex("CommittedUnitPrice")) And .ValueMatrix(Row, .ColIndex("CommittedUnitPrice")) <> 0 Then
                    mCancelEdit = True
                    MsgBox "Unit price exceeds committed unit price", vbExclamation, App.ProductName
                    .TextMatrix(Row, Col) = .ValueMatrix(Row, .ColIndex("CommittedUnitPrice"))
'                'check if over commited amount
'                'ElseIf .ValueMatrix(Row, .ColIndex("InvoicedQuantity")) * .ValueMatrix(Row, .ColIndex("InvoicedUnitPrice")) > .ValueMatrix(Row, .ColIndex("AmountRemaining")) Then
'                ElseIf ExceedsCommittedAmount() Then
'                    mCancelEdit = True
'                    MsgBox "Exceeds amount committed.", vbExclamation, App.ProductName
'                    .TextMatrix(Row, Col) = .ValueMatrix(Row, .ColIndex("CommittedUnitPrice"))
                Else
                    .TextMatrix(Row, .ColIndex("InvoicedUnitPrice")) = .ValueMatrix(Row, .ColIndex("InvoicedUnitPrice"))
                    .TextMatrix(Row, .ColIndex("PreTax")) = .ValueMatrix(Row, .ColIndex("InvoicedQuantity")) * .ValueMatrix(Row, .ColIndex("InvoicedUnitPrice"))
                    .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("PreTax")) * .ValueMatrix(Row, .ColIndex("TaxRate")) / 100, 2)
                    Call RecalcRetainage(Row)
                    Call RecalcTotals
                    Call VerifyAmounts(Row)
                End If
            
            Case .ColKey(Col) = "PreTax"
'                'check if over commited amount
'                'If .ValueMatrix(Row, Col) > .ValueMatrix(Row, .ColIndex("AmountRemaining")) Then
'                If ExceedsCommittedAmount() Then
'                    mCancelEdit = True
'                    MsgBox "Exceeds amount committed.", vbExclamation, App.ProductName
'                    .TextMatrix(Row, Col) = .ValueMatrix(Row, .ColIndex("AmountRemaining"))
'                End If
                .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("PreTax")) * .ValueMatrix(Row, .ColIndex("TaxRate")) / 100, 2)
                Call RecalcRetainage(Row)
                Call RecalcTotals
                Call VerifyAmounts(Row)
                
                
            Case .ColKey(Col) = "Tax"
                Call RecalcRetainage(Row)
                Call RecalcTotals
                        
            Case .ColKey(Col) = "RetainageRate"
                Call RecalcRetainage(Row)
                Call RecalcTotals
            
            Case .ColKey(Col) = "Retainage"
                .TextMatrix(Row, .ColIndex("RetainageRate")) = .ValueMatrix(Row, .ColIndex("Retainage")) / .ValueMatrix(Row, .ColIndex("Pretax")) * 100
                Call RecalcTotals
                        
        End Select
        
        
On Error Resume Next
        If mCancelEdit Then
            'Call .Select(Row, Col)
            Call .EditCell
            Exit Sub
        Else
            'they have edited the "newline" so add new "newline"
            'If Row = .Rows - 1 Then Stop
            If Row = .Rows - 1 And Trim(.Text) <> "" Then
                .AddItem ""
                .Cell(flexcpChecked, .Rows - 1, .ColIndex("Billable")) = IIf(mAPInvoicesDefaultToBillable, flexChecked, flexUnchecked)
                .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "DefRequired"
            
            
                'weirdness. the grid with lots of rows sometimes doesnt show the new last row. you can mouse into it..
                'added this junk to try to fix it
                gEnterDist.Redraw = flexRDNone
                gEnterDist.Redraw = flexRDBuffered
                Call .ShowCell(.Rows - 1, Col)
                
            End If
            
            Dirty = True
        End If
    Next
    End With
Exit Sub
eh: Call ErrHandler(SRCFILE & "gEnterDist_AfterEdit")
End Sub

Private Function ExceedsCommittedAmount(InvoiceAmt As Double, Row As Long) As Boolean
    
'    Dim CommittedAmt As Double
'
'
'    With gEnterDist
'        CommittedAmt = .ValueMatrix(Row, .ColIndex("CommittedUnitPrice")) * .ValueMatrix(Row, .ColIndex("CommittedQuantity"))
'
'
'    End With
'
'    ExceedsCommittedAmount = InvoicedAmt > OriginalAmt
End Function

Private Property Get CommitmentItemRemaining(Commitment As String, CommitmentItem As String, Optional Amount, Optional DirectPay) As Double
    Dim s          As String
    Dim rs         As Recordset
    Dim CrossPay   As Double
    Dim PendingInv As Double
    Dim CurrentInv As Double
    
    If IsMissing(Amount) Then
        If TimberlineAccounting Then
            s = ""
            s = s & "select iamt+iappcoa Amount" & vbCrLf
            s = s & "      ,iamtinv      Invoiced" & vbCrLf
            s = s & "  from master_jcm_record_13 " & vbCrLf
            If AccountingDB = dbTimberlinePVdata Then
                s = s & " where isub=" & DbQuote(Str, FormatTSField(12, Commitment)) & vbCrLf
            Else
                s = s & " where isub=" & DbQuote(Str, Commitment) & vbCrLf
            End If
            s = s & "   and item=" & DbQuote(num, CommitmentItem) & vbCrLf
        Else
            s = ""
            s = s & "select lineremainingpretax+lineremainingtax Amount" & vbCrLf
            s = s & "      ,lineinvoicedpretax+lineinvoicedtax Invoiced" & vbCrLf
            s = s & "  from jcpodetails" & vbCrLf
            s = s & " where DivisionID=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
            s = s & "   and PO=" & DbQuote(Str, Commitment) & vbCrLf
            s = s & "   and Line=" & DbQuote(num, CommitmentItem) & vbCrLf
        End If
        Set rs = HFApp.SqlExec(s, AccountingDB)
        If Not rs.EOF Then
            Amount = Val("" & rs("Amount"))
            DirectPay = Val("" & rs("Invoiced"))
        End If
    End If
    
    s = ""
    s = s & "select sum(pretax + tax)" & vbCrLf
    s = s & "  from invoiceitems" & vbCrLf
    s = s & " where DivisionID =" & HFApp.DivisionID & " and commitment=" & DbQuote(Str, Commitment) & vbCrLf
    s = s & "   and commitmentitem=" & DbQuote(num, CommitmentItem) & vbCrLf
    s = s & "   and vendor<>commitmentvendor" & vbCrLf
    CrossPay = Val("" & HFApp.SqlExec(s, dbHomeFront)(0))
    
    s = ""
    s = s & "select sum(invoiceitems.pretax + invoiceitems.tax)" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "  join invoiceitems on Invoices.InvoiceID=InvoiceItems.InvoiceID" & vbCrLf
    s = s & " where invoices.DivisionID =" & HFApp.DivisionID & " and invoiceitems.commitment=" & DbQuote(Str, Commitment) & vbCrLf
    s = s & "   and invoiceitems.commitmentitem=" & DbQuote(num, CommitmentItem) & vbCrLf
    s = s & "   and invoiceitems.vendor=invoiceitems.commitmentvendor" & vbCrLf
    s = s & "   and invoices.Status<>'Posted' and invoices.Status<>'Exported'" & vbCrLf
    If TimberlineAccounting Then
        PendingInv = Val("" & HFApp.SqlExec(s, dbHomeFront)(0))
    Else
        PendingInv = 0
    End If
    CurrentInv = InvoiceTotal(Commitment, CommitmentItem)
    
    CommitmentItemRemaining = Amount - CurrentInv - CrossPay - DirectPay - PendingInv


End Property



Private Sub VerifyAmounts(Row As Long, Optional AlwaysDisplay As Boolean = False)
    
    Dim Commitment     As String
    Dim CommitmentItem As Long
    Dim Job            As String
    Dim Extra          As String
    Dim Phase          As String
    Dim PhaseDesc      As String
    Dim Category       As String
    Dim CategoryDesc   As String
    
    Dim s   As String
    Dim rs  As Recordset
    
    Dim CommitmentAmount    As Double
    Dim CommitmentCurrent   As Double
    Dim CommitmentHold      As Double
    Dim CommitmentInvoiced  As Double
    
    Dim PhaseAmount         As Double
    Dim PhaseCurrent        As Double
    Dim PhaseHold           As Double
    Dim PhaseInvoiced       As Double
    Dim CategoryAmount      As Double
    Dim CategoryCurrent     As Double
    Dim CategoryHold        As Double
    Dim CategoryInvoiced    As Double
    Dim CommittedQuantity   As Double
    Dim CommittedUnitPrice  As Double
    Dim InvoicedQuantity    As Double
    Dim InvoicedUnitPrice   As Double
    
    With gEnterDist
    
        If Row = .Rows - 1 Then Exit Sub
        If Row < 1 Then Exit Sub
        If .TextMatrix(Row, .ColIndex("Job")) = "" Then Exit Sub
    
        Commitment = .TextMatrix(Row, .ColIndex("Commitment"))
        CommitmentItem = .ValueMatrix(Row, .ColIndex("CommitmentItem"))
        Job = .TextMatrix(Row, .ColIndex("Job"))
        Extra = .TextMatrix(Row, .ColIndex("Extra"))
        Phase = .TextMatrix(Row, .ColIndex("Phase"))
        PhaseDesc = .TextMatrix(Row, .ColIndex("PhaseDesc"))
        Category = .TextMatrix(Row, .ColIndex("Category"))
        CategoryDesc = .TextMatrix(Row, .ColIndex("CategoryDesc"))
    
        CommittedQuantity = .ValueMatrix(Row, .ColIndex("CommittedQuantity"))
        CommittedUnitPrice = .ValueMatrix(Row, .ColIndex("CommittedUnitPrice"))
        InvoicedQuantity = .ValueMatrix(Row, .ColIndex("InvoicedQuantity"))
        InvoicedUnitPrice = .ValueMatrix(Row, .ColIndex("InvoicedUnitPrice"))
    
    End With
    
    If Job <> "" And App.Options(VerifyBudgets) Then
        
        Screen.MousePointer = vbHourglass
    
        'get phase budget
        Set rs = HFApp.SqlExec(GetBudget(Job, Extra, Phase), AccountingDB)
        PhaseAmount = Val("" & rs("Budget"))
        PhaseCurrent = InvoiceTotal(, , Job, Extra, Phase, "", True)
        PhaseInvoiced = Val("" & rs("Actual"))
        
        'get category budget
        Set rs = HFApp.SqlExec(GetBudget(Job, Extra, Phase, Category), AccountingDB)
        CategoryAmount = Val("" & rs("Budget"))
        CategoryCurrent = InvoiceTotal(, , Job, Extra, Phase, Category, True)
        CategoryInvoiced = Val("" & rs("Actual"))
        
        
        
        'get posted phase amount
        s = ""
        s = s & "SELECT SUM(d.Pretax+d.jctax) Actual" & vbCrLf
        s = s & "  FROM Invoices m " & vbCrLf
        s = s & "      ,InvoiceItems d " & vbCrLf
        s = s & " WHERE m.DivisionID =" & HFApp.DivisionID & " and m.InvoiceID=d.InvoiceID" & vbCrLf
        s = s & "   AND m.InvoiceID<>" & DbQuote(num, mInvoiceID) & vbCrLf
        s = s & "   AND m.Status IN('Exported','Posted','Rejected')" & vbCrLf
        s = s & "   AND d.Job  =" & DbQuote(Str, Trim(Job)) & vbCrLf
        s = s & "   AND d.Extra=" & DbQuote(Str, Trim(Extra)) & vbCrLf
        s = s & "   AND d.Phase=" & DbQuote(Str, Trim(Phase)) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        PhaseInvoiced = PhaseInvoiced + Val("" & rs("Actual"))
        
        'get posted category amount
        s = s & "   AND d.Category=" & DbQuote(Str, Trim(Category))
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        CategoryInvoiced = CategoryInvoiced + Val("" & rs("Actual"))
        
        
        
        'get held phase amount
        s = ""
        s = s & "SELECT SUM(d.Pretax+d.jctax) Actual" & vbCrLf
        s = s & "  FROM Invoices m " & vbCrLf
        s = s & "      ,InvoiceItems d " & vbCrLf
        s = s & " WHERE m.DivisionID =" & HFApp.DivisionID & " and m.InvoiceID=d.InvoiceID" & vbCrLf
        s = s & "   AND m.InvoiceID<>" & DbQuote(num, mInvoiceID) & vbCrLf
        s = s & "   AND m.Status ='Hold'" & vbCrLf
        s = s & "   AND d.Job  =" & DbQuote(Str, Trim(Job)) & vbCrLf
        s = s & "   AND d.Extra=" & DbQuote(Str, Trim(Extra)) & vbCrLf
        s = s & "   AND d.Phase=" & DbQuote(Str, Trim(Phase)) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        PhaseHold = Val("" & rs("Actual"))
        
        'get held category amount
        s = s & "   AND d.Category=" & DbQuote(Str, Trim(Category))
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        CategoryHold = Val("" & rs("Actual"))
        
    End If
    
    
    If Commitment <> "" And CommitmentItem <> 0 Then
    
        CommitmentCurrent = InvoiceTotal(Commitment, CommitmentItem, , , , , True)
         
        Set rs = HFApp.SqlExec(GetCommitmentItemAmount(Commitment, CommitmentItem), AccountingDB)
        If Not rs.EOF Then
            CommitmentAmount = Val("" & rs("Pretax"))
            CommitmentInvoiced = Round(Val("" & rs("InvoicedTaxIn")) / (1 + Val("" & rs("TaxRate")) / 100), 2)
        End If
        'add crosspayed invoices that are not on hold
        s = ""
        s = s & "select sum(d.pretax+d.jctax)" & vbCrLf
        s = s & "  from Invoices m " & vbCrLf
        s = s & "      ,InvoiceItems d " & vbCrLf
        s = s & " where m.DivisionID =" & HFApp.DivisionID & " and m.InvoiceID=d.InvoiceID" & vbCrLf
        s = s & "   and d.commitment=" & DbQuote(Str, Commitment) & vbCrLf
        s = s & "   and d.commitmentitem=" & DbQuote(num, CommitmentItem) & vbCrLf
        s = s & "   and m.Status<>'Hold'" & vbCrLf
        s = s & "   and d.vendor<>d.commitmentvendor" & vbCrLf
        CommitmentInvoiced = CommitmentInvoiced + Val("" & HFApp.SqlExec(s)(0))
        
        
        'add correct payed invoices
        s = ""
        s = s & "select sum(d.pretax+d.jctax)" & vbCrLf
        s = s & "  from Invoices m " & vbCrLf
        s = s & "      ,InvoiceItems d " & vbCrLf
        s = s & " where m.DivisionID =" & HFApp.DivisionID & " and m.InvoiceID=d.InvoiceID" & vbCrLf
        s = s & "   and d.commitment=" & DbQuote(Str, Commitment) & vbCrLf
        s = s & "   and d.commitmentitem=" & DbQuote(num, CommitmentItem) & vbCrLf
        s = s & "   and m.Status in('Exported','Posted','Rejected')" & vbCrLf
        s = s & "   and d.vendor=d.commitmentvendor" & vbCrLf
        CommitmentInvoiced = CommitmentInvoiced + Val("" & HFApp.SqlExec(s)(0))
        
        'get held commitment amount
        s = ""
        s = s & "SELECT SUM(d.Pretax+d.jctax) Actual" & vbCrLf
        s = s & "  FROM Invoices m " & vbCrLf
        s = s & "      ,InvoiceItems d " & vbCrLf
        s = s & " WHERE m.DivisionID =" & HFApp.DivisionID & " and m.InvoiceID=d.InvoiceID" & vbCrLf
        s = s & "   AND m.InvoiceID<>" & DbQuote(num, mInvoiceID) & vbCrLf
        s = s & "   AND m.Status ='Hold'" & vbCrLf
        s = s & "   AND d.Commitment=" & DbQuote(Str, Trim(Commitment)) & vbCrLf
        s = s & "   AND d.CommitmentItem=" & DbQuote(num, Trim(CommitmentItem)) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        CommitmentHold = Val("" & rs("Actual"))
    
    End If
    
    
    CommitmentAmount = Round(CommitmentAmount, 2)
    CommitmentCurrent = Round(CommitmentCurrent, 2)
    CommitmentHold = Round(CommitmentHold, 2)
    CommitmentInvoiced = Round(CommitmentInvoiced, 2)
    PhaseAmount = Round(PhaseAmount, 2)
    PhaseCurrent = Round(PhaseCurrent, 2)
    PhaseHold = Round(PhaseHold, 2)
    PhaseInvoiced = Round(PhaseInvoiced, 2)
    CategoryAmount = Round(CategoryAmount, 2)
    CategoryCurrent = Round(CategoryCurrent, 2)
    CategoryHold = Round(CategoryHold, 2)
    CategoryInvoiced = Round(CategoryInvoiced, 2)
    
    If AlwaysDisplay _
        Or ((CommitmentAmount <> 0) And (CommitmentAmount - CommitmentCurrent - CommitmentInvoiced - CommitmentHold < 0)) _
        Or (App.Options(VerifyBudgets) And (PhaseAmount - PhaseCurrent - PhaseInvoiced - PhaseHold < 0)) _
        Or (App.Options(VerifyBudgets) And (CategoryAmount - CategoryCurrent - CategoryInvoiced - CategoryHold < 0)) _
        Or (App.Options(ShowQtyAndUnitPrice) And InvoicedQuantity > CommittedQuantity And CommittedQuantity <> 0) _
        Or (App.Options(ShowQtyAndUnitPrice) And InvoicedUnitPrice > CommittedUnitPrice And CommittedUnitPrice <> 0) Then
        
        mLoading = True
        Call FAmounts.ShowAmounts(AlwaysDisplay, PhaseDesc, CategoryDesc, _
                                  CommitmentAmount, CommitmentCurrent, CommitmentHold, CommitmentInvoiced, _
                                  PhaseAmount, PhaseCurrent, PhaseHold, PhaseInvoiced, _
                                  CategoryAmount, CategoryCurrent, CategoryHold, CategoryInvoiced, _
                                  CommittedQuantity, CommittedUnitPrice, _
                                  InvoicedQuantity, InvoicedUnitPrice)
        mLoading = False
    End If
    Screen.MousePointer = vbDefault
    
    
End Sub



Private Sub gEnterDist_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo eh 'Resume Next
    Dim r As Long
    Dim c As Long
    Dim Cancel As Boolean
    
If ReadOnly Then Exit Sub
    
    Select Case KeyCode
    
        Case vbKeyF9
            Call VerifyAmounts(gEnterDist.Row, True)
            
        Case vbKeyDelete
            If Shift = 0 Then
                'clear cell
                Call gEnterDist_BeforeEdit(gEnterDist.Row, gEnterDist.Col, Cancel)
                If Not Cancel Then
                    gEnterDist.Text = ""
                    r = gEnterDist.Row
                    c = gEnterDist.Col
                    Call gEnterDist_AfterEdit(r, c)
                    Call RecalcTotals
                End If
            Else
                'delete line
                If IsBetween(gEnterDist.Row, 1, gEnterDist.Rows - 2) Then
                    If gEnterDist.ValueMatrix(gEnterDist.Row, gEnterDist.ColIndex("ItemID")) = 0 Then
                        Call gEnterDist.RemoveItem(gEnterDist.Row)
                    Else
                        gEnterDist.RowHidden(gEnterDist.Row) = True

                    End If
                    Call RecalcTotals

                    Call gEnterDist_RowColChange
                End If
            End If
            Dirty = True
    
        Case vbKeyF4
            If Shift = 0 Then
                Call gEnterDist_BeforeEdit(gEnterDist.Row, gEnterDist.Col, Cancel)
                If gEnterDist.ComboList <> "" And Not Cancel Then
                    Call gEnterDist_CellButtonClick(gEnterDist.Row, gEnterDist.Col)
                End If
            End If
            
        Case vbKeyReturn
            KeyCode = 0
            If gEnterDist.Col = gEnterDist.Cols - 1 Then
                Call gEnterDist.Select(gEnterDist.Row + 1, 0)
            Else
                Call gEnterDist.Select(gEnterDist.Row, gEnterDist.Col + 1)
            End If
            
            
    End Select
    Exit Sub
eh:
    If err.Number <> 0 Then
        MsgBox err.Description
    End If
End Sub

Private Sub gEnterDist_RowColChange()
On Error Resume Next
    With gEnterDist
        gEnterTail.TextMatrix(0, 1) = .TextMatrix(.Row, .ColIndex("JobDesc"))
        gEnterTail.TextMatrix(1, 1) = .TextMatrix(.Row, .ColIndex("ExtraDesc"))
        gEnterTail.TextMatrix(2, 1) = .TextMatrix(.Row, .ColIndex("PhaseDesc"))
        gEnterTail.TextMatrix(3, 1) = .TextMatrix(.Row, .ColIndex("CategoryDesc"))
        
        gEnterTail.TextMatrix(0, 3) = .TextMatrix(.Row, .ColIndex("EquipmentDesc"))
        gEnterTail.TextMatrix(1, 3) = .TextMatrix(.Row, .ColIndex("EQCostCodeDesc"))
        gEnterTail.TextMatrix(2, 3) = .TextMatrix(.Row, .ColIndex("DebitAccountDesc"))
        gEnterTail.TextMatrix(3, 3) = .TextMatrix(.Row, .ColIndex("TaxGroupDesc"))
        
        If .ColHidden(.Col) Then .Col = .Col + 1
        If .RowHidden(.Row) Then .Row = .Row + 1
    End With
End Sub


Private Sub AddPO(WhereClause As String)
On Error GoTo eh
    Dim s  As String
    Dim rs As Recordset
    Dim r  As Long
    Dim Amount  As Double
    Dim Retainage  As Double
    Dim Pretax  As Double
    Dim Tax     As Double
    Dim TaxRate As Double
    
    If App.Options(SchedulingUsage) <> DoNotUse Then
        Set rs = HFApp.SqlExec(GetSelectedCommitments(WhereClause), AccountingDB)
        While Not rs.EOF
            If Not SchedulingIsFinished("" & rs(0)) Then
                Exit Sub
            End If
            rs.MoveNext
        Wend
    End If
    
    s = GetSelectedCommitmentItems(WhereClause)
    Set rs = HFApp.SqlExec(s, AccountingDB)
    With gEnterDist
        .Redraw = flexRDNone
        
        If Not rs.EOF Then
            While Not rs.EOF
                
                If App.Options(AutoCalcHeaderTotals) Then
                    Amount = CommitmentItemRemaining(Trim("" & rs("Commitment")), Val("" & rs("CommitmentItem")), Val("" & rs("OriginalAmount")), Val("" & rs("InvoicedAmount")))
                    Pretax = Amount / (Val("" & rs("TaxRate")) / 100 + 1)
                    
                Else
                    Call RecalcTotals
                    Amount = CommitmentItemRemaining(Trim("" & rs("Commitment")), Val("" & rs("CommitmentItem")), Val("" & rs("OriginalAmount")), Val("" & rs("InvoicedAmount")))
                    If Amount > InvoiceRemaining Then
                        Amount = InvoiceRemaining
                    End If
                End If
                 
                If Amount <> 0 Then
                    TaxRate = Val("" & rs("TaxRate"))
                    Pretax = Round(Amount / (1 + TaxRate / 100), 2)
                    Tax = Round(Amount - Pretax, 2)
                    
                    
                    If Pretax <> 0 Or Tax <> 0 Then
                    
                        r = .Rows - 1
                        .AddItem "", r
                        .TextMatrix(r, .ColIndex("CommitmentVendor")) = Trim("" & rs("CommitmentVendor"))
                        .TextMatrix(r, .ColIndex("Commitment")) = Trim("" & rs("Commitment"))
                        .TextMatrix(r, .ColIndex("CommitmentItem")) = Trim("" & rs("CommitmentItem"))
                        .TextMatrix(r, .ColIndex("Job")) = Trim("" & rs("Job"))
                        .TextMatrix(r, .ColIndex("JobDesc")) = Trim("" & rs("JobDesc"))
                        .TextMatrix(r, .ColIndex("Extra")) = Trim("" & rs("Extra"))
                        .TextMatrix(r, .ColIndex("ExtraDesc")) = Trim("" & rs("ExtraDesc"))
                        .TextMatrix(r, .ColIndex("Phase")) = Trim("" & rs("Phase"))
                        .TextMatrix(r, .ColIndex("PhaseDesc")) = Trim("" & rs("PhaseDesc"))
                        .TextMatrix(r, .ColIndex("Category")) = Trim("" & rs("Category"))
                        .TextMatrix(r, .ColIndex("CategoryDesc")) = Trim("" & rs("CategoryDesc"))
                        
                        .TextMatrix(r, .ColIndex("DebitAccount")) = Trim("" & rs("Account"))
                        .TextMatrix(r, .ColIndex("DebitAccountDesc")) = Trim("" & rs("AccountDesc"))
                        
                        .TextMatrix(r, .ColIndex("TaxGroup")) = Trim("" & rs("TaxGroup"))
                        .TextMatrix(r, .ColIndex("TaxGroupDesc")) = Trim("" & rs("TaxGroupDesc"))
                        .TextMatrix(r, .ColIndex("TaxRate")) = TaxRate
                        .TextMatrix(r, .ColIndex("CommittedQuantity")) = Val("" & rs("CommittedQuantity"))
                        .TextMatrix(r, .ColIndex("CommittedUnitPrice")) = Val("" & rs("CommittedUnitPrice"))
                        .TextMatrix(r, .ColIndex("InvoicedQuantity")) = Val("" & rs("RemainingQuantity"))
                        .TextMatrix(r, .ColIndex("InvoicedUnitPrice")) = .ValueMatrix(r, .ColIndex("CommittedUnitPrice"))
                        .TextMatrix(r, .ColIndex("PreTax")) = Pretax
                        .TextMatrix(r, .ColIndex("Tax")) = Tax
                        
                        .TextMatrix(r, .ColIndex("RetainageRate")) = Val("" & rs("RetainageRate"))
                        
                        If HFApp.Options(AccountingSystem) = asTimberline Then
                            .TextMatrix(r, .ColIndex("Description")) = left(Trim("" & rs("CommitmentItemDesc")), 30)
                        Else
                            .TextMatrix(r, .ColIndex("Description")) = Trim("" & rs("CommitmentItemDesc"))
                        End If
                        
                        If App.Options(InvDescDefaultsToJobAddr) And Trim(txtDescription.Text) = "" Then txtDescription.Text = "" & rs("JobAddr")
                        
                        Call RecalcRetainage(r)
'                        Call VerifyAmounts(r)
                        
                        .Redraw = flexRDDirect
                    End If
                End If
                rs.MoveNext
            Wend
            
            Call RecalcTotals
            If App.Options(AutoVarianceRemaining) Then
                Call VarianceRemaining
            End If
            If .Rows > 1 Then
                Call GetInvoiceNumber
            End If
            Dirty = True
        End If
        .Redraw = flexRDDirect
    End With
Exit Sub
eh: Call ErrHandler(SRCFILE & "AddPO")
End Sub

Private Sub VarianceRemaining()
On Error Resume Next
    Dim d As Double
    Dim r As Long
    Dim i As Long
    Dim s As String
    
    d = InvoiceRemaining
    If d = 0 Then Exit Sub
    With gEnterDist
        
        'which row are we going to take coding from??
        i = .Row 'current
        If i < 1 Then i = .Rows - 2 'if not valid then use last row
        
        'add new row
        r = .Rows - 1
        .AddItem "", r
        
        
        
        'set coding
        If r > 1 Then
            .TextMatrix(r, .ColIndex("DebitAccount")) = .TextMatrix(i, .ColIndex("DebitAccount"))
            .TextMatrix(r, .ColIndex("DebitAccountDesc")) = .TextMatrix(i, .ColIndex("DebitAccountDesc"))
            
            .TextMatrix(r, .ColIndex("Job")) = .TextMatrix(i, .ColIndex("Job"))
            .TextMatrix(r, .ColIndex("JobDesc")) = .TextMatrix(i, .ColIndex("JobDesc"))
            .TextMatrix(r, .ColIndex("Extra")) = .TextMatrix(i, .ColIndex("Extra"))
            .TextMatrix(r, .ColIndex("ExtraDesc")) = .TextMatrix(i, .ColIndex("ExtraDesc"))
            .TextMatrix(r, .ColIndex("Phase")) = .TextMatrix(i, .ColIndex("Phase"))
            .TextMatrix(r, .ColIndex("PhaseDesc")) = .TextMatrix(i, .ColIndex("PhaseDesc"))
            
            .TextMatrix(r, .ColIndex("TaxGroup")) = App.Options(DefaultTaxGroup)
            .TextMatrix(r, .ColIndex("TaxGroupDesc")) = App.Options(DefaultTaxGroupDesc)
            .TextMatrix(r, .ColIndex("TaxRate")) = App.Options(DefaultTaxGroupRate)
        
        
        End If
        
        .TextMatrix(r, .ColIndex("Category")) = mV99
        .TextMatrix(r, .ColIndex("CategoryDesc")) = "un-coded variance"
        .TextMatrix(r, .ColIndex("PreTax")) = d / (1 + .ValueMatrix(r, .ColIndex("TaxRate")) / 100)
        
        
        
        On Error Resume Next
        s = "select dbo.purch_getdefaulttaxgroup(" & DbQuote(Str, .TextMatrix(r, .ColIndex("Job"))) & ",'','','','','',''," & DbQuote(Str, txtVendor.tag) & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Category"))) & "," & DbQuote(num, HFApp.DivisionID) & ")"
        .TextMatrix(r, .ColIndex("TaxGroup")) = "" & HFApp.SqlExec(s)(0)
        .Row = r
        Call gEnterDist_AfterEdit(r, .ColIndex("TaxGroup"))
        On Error GoTo 0

        
        
        .TextMatrix(r, .ColIndex("Tax")) = d - .ValueMatrix(r, .ColIndex("PreTax"))
        
        On Error Resume Next
        Call .Select(r, .ColIndex("Category"))
        .SetFocus
    
    End With
    Call RecalcTotals
    Dirty = Dirty
End Sub

Private Sub ApplySettings()
On Error Resume Next

Dim d As Boolean
Dim rs As Recordset
d = Dirty
    
    
    With App.Options
    
        Toolbar.Buttons("documents").Visible = True '.Value(AllowScanning)
        
        'enter frame
        
        txtInvoice.MaxLength = MaxInvoiceLength()
        
        
        dteInvoice.Format = .Value(TimberlineDateFormat)
        dteInvoice.DisplayFormat = .Value(TimberlineDateFormat)
        dteReceived.Format = .Value(TimberlineDateFormat)
        dteReceived.DisplayFormat = .Value(TimberlineDateFormat)
        dtePayment.Format = .Value(TimberlineDateFormat)
        dtePayment.DisplayFormat = .Value(TimberlineDateFormat)
        dteDiscount.Format = .Value(TimberlineDateFormat)
        dteDiscount.DisplayFormat = .Value(TimberlineDateFormat)
        dteAccounting.Format = .Value(TimberlineDateFormat)
        dteAccounting.DisplayFormat = .Value(TimberlineDateFormat)
        
        lblVendor.Text = .Value(Caption_Vendor)
        lblJob.Caption = .Value(Caption_Job)
            
            
            
        lblWrapInsuranceAmount.Caption = .Value(Caption_WrapInsurance)
            
        lblWrapInsuranceAmount.Visible = HFApp.Options(AccountingSystem) = asTimberline
        numWrapInsuranceAmount.Visible = HFApp.Options(AccountingSystem) = asTimberline
        lblWrapInsuranceRate.Visible = HFApp.Options(AccountingSystem) = asTimberline
            
            
        lblInvoiceCode1.Caption = .Value(InvoiceCode1Label)
        lblInvoiceCode2.Caption = .Value(InvoiceCode2Label)
        
        txtInvoiceCode1.Visible = .Value(InvoiceCode1Usage) <> "Hidden"
        lblInvoiceCode1.Visible = txtInvoiceCode1.Visible
        lblInvoiceCode1.Caption = .Value(InvoiceCode1Label)
        txtInvoiceCode2.Visible = .Value(InvoiceCode2Usage) <> "Hidden"
        lblInvoiceCode2.Visible = txtInvoiceCode2.Visible
        lblInvoiceCode2.Caption = .Value(InvoiceCode2Label)
        
        Toolbar.Buttons("add").ToolTipText = "Add " & LCase(.Value(Caption_Commitment)) & " items to the invoice"
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Commitment")) = .Value(Caption_Commitment)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Job")) = .Value(Caption_Job)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Extra")) = .Value(Caption_Extra)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Phase")) = .Value(Caption_Phase)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Category")) = .Value(Caption_Category)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("DebitAccount")) = .Value(Caption_DebitAccount)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("TaxGroup")) = .Value(Caption_TaxGroup)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Retainage")) = .Value(Caption_Retainage)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("RetainageRate")) = .Value(Caption_RetainagePct)
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("WrapInsuranceExempt")) = .Value(Caption_WrapInsurance) & " Exempt"
        
        gEnterDist.ColHidden(gEnterDist.ColIndex("InvoicedQuantity")) = Not App.Options(ShowQtyAndUnitPrice)
        gEnterDist.ColHidden(gEnterDist.ColIndex("InvoicedUnitPrice")) = Not App.Options(ShowQtyAndUnitPrice)
        
                
        If HFApp.Options(AccountingSystem) = asTimberline Then
            If gEnterDist.TextMatrix(0, gEnterDist.ColIndex("JointPayee")) = "" Then gEnterDist.TextMatrix(0, gEnterDist.ColIndex("JointPayee")) = "Joint Payee"
            If gViewDist.TextMatrix(0, gViewDist.ColIndex("JointPayee")) = "" Then gViewDist.TextMatrix(0, gViewDist.ColIndex("JointPayee")) = "Joint Payee"
        Else
            gEnterDist.ColHidden(gEnterDist.ColIndex("JointPayee")) = True
            gEnterDist.TextMatrix(0, gEnterDist.ColIndex("JointPayee")) = ""
            gViewDist.ColHidden(gEnterDist.ColIndex("JointPayee")) = True
            gViewDist.TextMatrix(0, gViewDist.ColIndex("JointPayee")) = ""
        End If
        
        gEnterTail.TextMatrix(0, 0) = .Value(Caption_Job)
        gEnterTail.TextMatrix(1, 0) = .Value(Caption_Extra)
        gEnterTail.TextMatrix(2, 0) = .Value(Caption_Phase)
        gEnterTail.TextMatrix(3, 0) = .Value(Caption_Category)
        gEnterTail.TextMatrix(0, 2) = "Equipment"
        gEnterTail.TextMatrix(1, 2) = "EQ Cost Code"
        gEnterTail.TextMatrix(2, 2) = .Value(Caption_DebitAccount)
        gEnterTail.TextMatrix(3, 2) = .Value(Caption_TaxGroup)
        
        
        numAmount.Enabled = Not App.Options(AutoCalcHeaderTotals)
        numTax.Enabled = Not App.Options(AutoCalcHeaderTotals)
        numDiscount.TabStop = Not App.Options(AutoCalcHeaderTotals)
        
        'view frame
        dteSearchInvoiceDate.Format = .Value(TimberlineDateFormat)
        dteSearchInvoiceDate.DisplayFormat = .Value(TimberlineDateFormat)
        dteSearchReceivedDate.Format = .Value(TimberlineDateFormat)
        dteSearchReceivedDate.DisplayFormat = .Value(TimberlineDateFormat)
        dteSearchPaymentDate.Format = .Value(TimberlineDateFormat)
        dteSearchPaymentDate.DisplayFormat = .Value(TimberlineDateFormat)
        dteSearchAccountingDate.Format = .Value(TimberlineDateFormat)
        dteSearchAccountingDate.DisplayFormat = .Value(TimberlineDateFormat)
        
        lblSearchVendor.Caption = .Value(Caption_Vendor)
        lblSearchJob.Caption = .Value(Caption_Job)
        lblSearchCommitment.Caption = .Value(Caption_Commitment)
        
        gViewInv.ColFormat(gViewInv.ColIndex("InvoiceDate")) = App.Options(TimberlineDateFormat)
        gViewInv.ColFormat(gViewInv.ColIndex("AccountingDate")) = App.Options(TimberlineDateFormat)
        gViewInv.ColFormat(gViewInv.ColIndex("DiscountDate")) = App.Options(TimberlineDateFormat)
        gViewInv.ColFormat(gViewInv.ColIndex("PaymentDate")) = App.Options(TimberlineDateFormat)
        gViewInv.ColFormat(gViewInv.ColIndex("DStmp")) = App.Options(TimberlineDateFormat)
        
        gViewInv.TextMatrix(0, gViewInv.ColIndex("Vendor")) = .Value(Caption_Vendor)
        gViewInv.TextMatrix(0, gViewInv.ColIndex("VendorName")) = .Value(Caption_Vendor) & " Name"
        gViewInv.TextMatrix(0, gViewInv.ColIndex("Job")) = .Value(Caption_Job)
        gViewInv.TextMatrix(0, gViewInv.ColIndex("JobDesc")) = .Value(Caption_Job) & " Description"
        
        
        
        gViewInv.TextMatrix(0, gViewInv.ColIndex("InvoiceCode1")) = IIf(.Value(InvoiceCode1Usage) = "Hidden", "", .Value(InvoiceCode1Label))
        gViewInv.TextMatrix(0, gViewInv.ColIndex("InvoiceCode2")) = IIf(.Value(InvoiceCode2Usage) = "Hidden", "", .Value(InvoiceCode2Label))
        gViewInv.ColHidden(gViewInv.ColIndex("InvoiceCode1")) = .Value(InvoiceCode1Usage) = "Hidden"
        gViewInv.ColHidden(gViewInv.ColIndex("InvoiceCode2")) = .Value(InvoiceCode2Usage) = "Hidden"
        
        
        gViewDist.TextMatrix(0, gViewDist.ColIndex("Commitment")) = .Value(Caption_Commitment)
        gViewDist.TextMatrix(0, gViewDist.ColIndex("CommitmentDesc")) = .Value(Caption_Commitment) & " Description"
        gViewDist.TextMatrix(0, gViewDist.ColIndex("Job")) = .Value(Caption_Job)
        gViewDist.TextMatrix(0, gViewDist.ColIndex("JobDesc")) = .Value(Caption_Job) & " Description"
        gViewDist.TextMatrix(0, gViewDist.ColIndex("Extra")) = .Value(Caption_Extra)
        gViewDist.TextMatrix(0, gViewDist.ColIndex("ExtraDesc")) = .Value(Caption_Extra) & " Description"
        gViewDist.TextMatrix(0, gViewDist.ColIndex("Phase")) = .Value(Caption_Phase)
        gViewDist.TextMatrix(0, gViewDist.ColIndex("PhaseDesc")) = .Value(Caption_Phase) & " Description"
        gViewDist.TextMatrix(0, gViewDist.ColIndex("Category")) = .Value(Caption_Category)
        gViewDist.TextMatrix(0, gViewDist.ColIndex("CategoryDesc")) = .Value(Caption_Category) & " Description"
        gViewDist.TextMatrix(0, gViewDist.ColIndex("DebitAccount")) = .Value(Caption_DebitAccount)
        gViewDist.TextMatrix(0, gViewDist.ColIndex("DebitAccountDesc")) = .Value(Caption_DebitAccount) & " Description"
        gViewDist.TextMatrix(0, gViewDist.ColIndex("TaxGroup")) = .Value(Caption_TaxGroup)
        gViewDist.TextMatrix(0, gViewDist.ColIndex("TaxGroupDesc")) = .Value(Caption_TaxGroup) & " Description"
        gViewDist.TextMatrix(0, gViewDist.ColIndex("Retainage")) = .Value(Caption_Retainage)
        
        
        If HFApp.Options(AccountingSystem) <> asTimberline Then
            gViewDist.TextMatrix(0, gViewDist.ColIndex("Retainage")) = ""
            gViewDist.ColHidden(gViewDist.ColIndex("Retainage")) = True
        End If
        
        
    End With
    
    Dirty = d

End Sub


Private Function ItemOnInvoice(Commitment As String, CommitmentItem As Long) As Boolean
On Error Resume Next
    Dim r As Long
    With gEnterDist
        For r = 1 To .Rows - 2
            If Not .RowHidden(r) And Commitment = .TextMatrix(r, .ColIndex("Commitment")) Then
                If CommitmentItem = 0 Or CommitmentItem = .ValueMatrix(r, .ColIndex("CommitmentItem")) Then
                    ItemOnInvoice = True
                    Exit Function
                End If
            End If
        Next
    End With
    ItemOnInvoice = False
End Function

Private Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property
Private Property Let ReadOnly(RHS As Boolean)
    mReadOnly = RHS
    txtVendor.Enabled = Not RHS
    txtInvoice.Enabled = Not RHS
    lblVendor.Enabled = Not RHS
    lblInvoice.Enabled = Not RHS
    
    txtJob.Enabled = Not RHS
    lblJob.Enabled = Not RHS
    
    txtInvoiceCode1.Enabled = Not RHS
    txtInvoiceCode2.Enabled = Not RHS
    
    txtDescription.Enabled = Not RHS
    txtComments.Enabled = Not RHS
    
    cboDepartment.Enabled = Not RHS
    cboApprover.Enabled = Not RHS
    
    
    numAmount.Enabled = Not RHS
    numTax.Enabled = Not RHS
    
    numAmount.Enabled = Not App.Options(AutoCalcHeaderTotals) And Not RHS
    numTax.Enabled = Not App.Options(AutoCalcHeaderTotals) And Not RHS
    
    numWrapInsuranceAmount.Enabled = False
    numDiscount.Enabled = Not RHS
    dteInvoice.Enabled = Not RHS
    dteReceived.Enabled = Not RHS
    dtePayment.Enabled = Not RHS
    dteDiscount.Enabled = Not RHS
    dteAccounting.Enabled = Not RHS
    Dirty = False
End Property

Public Property Get TotalCommitmentValue() As Double
    'returns the amount of this invoice allocated to commitment items
    Dim r As Long
    
    Dim v As Double
    With gEnterDist
        For r = 1 To .Rows - 2
            If Not .RowHidden(r) And .TextMatrix(r, .ColIndex("Commitment")) <> "" Then
                v = v + .ValueMatrix(r, .ColIndex("PreTax")) + .ValueMatrix(r, .ColIndex("Tax"))
            End If
        Next
    End With
    TotalCommitmentValue = v
End Property

Public Property Get ExceedsLimit() As Boolean
    Dim invAmt   As Double
    Dim ComAmt   As Double
    Dim variance As Double

    If IsPOInvoice Then
        'invoices that are entirely coded to POs.
        ExceedsLimit = False
    Else
        'invoices that have lines NOT coded to POs.
                
        'invoice total must not exceed limit
        invAmt = Val("" & numAmount)
        If invAmt > App.InvoiceMax Then
            ExceedsLimit = True
            Exit Property
        End If
        
        'variance amount must not exceed limits
        ComAmt = TotalCommitmentValue
        If ComAmt > 0 Then
            variance = Round(Max(0, invAmt - ComAmt), 2)
            ExceedsLimit = variance > App.InvOverrideMax _
                        Or variance > Abs(ComAmt * App.InvOverridePcnt / 100)
        End If
    End If
    
End Property

Private Function InvoiceApproved(InvoiceIDs As String, Optional filter As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "where status not in('Rejected','SiteApproved','Approved')" & vbCrLf
    If filter = "" Then
        s = s & "and invoiceid in(" & InvoiceIDs & ")" & vbCrLf
    Else
        s = s & filter
    End If
    Set rs = HFApp.SqlExec(s)
    
    InvoiceApproved = rs.EOF
End Function

Private Function InvoicesPosted(InvoiceIDs As String, Optional filter As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "where status in('Exported','Posted')" & vbCrLf
    If filter = "" Then
        s = s & "and invoiceid in(" & InvoiceIDs & ")" & vbCrLf
    Else
        s = s & filter
    End If
    Set rs = HFApp.SqlExec(s)
    
    InvoicesPosted = Not rs.EOF
End Function

Private Function ExceedsLimitDB(InvoiceIDs As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    
    s = ""
    s = s & "select" & vbCrLf
    s = s & "   i.invoiceid" & vbCrLf
    s = s & "  ,i.pretax+i.tax InvoiceTotal" & vbCrLf
    s = s & "  ,sum(case when isnull(d.commitment,'')='' then 1 else 0 end) NonPORows" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "  ,i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) PO" & vbCrLf
    s = s & "  ,                 sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) Variance" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "  --ExceedsInvMax = InvoiceTotal > InvoiceMax" & vbCrLf
    s = s & "  ,case when i.pretax+i.tax > " & DbQuote(num, App.InvoiceMax) & " then 1 else 0 end ExceedsInvMax" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "  --ExceedsOvrMax = PO>0 and Variance > OverrideMax" & vbCrLf
    s = s & "  ,case when i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > 0 " & vbCrLf
    s = s & "        and sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > " & DbQuote(num, App.InvOverrideMax) & vbCrLf
    s = s & "        then 1 else 0 end ExceedsOvrMax" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "  --ExceedsOvrPct = PO>0 and Variance > PO * OverridePcnt" & vbCrLf
    s = s & "  ,case when i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > 0 " & vbCrLf
    s = s & "        and sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > (i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end)) * " & DbQuote(num, App.InvOverridePcnt) & "/100" & vbCrLf
    s = s & "        then 1 else 0 end ExceedsOvrPct" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "from invoices i" & vbCrLf
    s = s & "join invoiceitems d on i.invoiceid=d.invoiceid" & vbCrLf
    s = s & "where i.invoiceid in (" & InvoiceIDs & ")" & vbCrLf
    s = s & "group by i.invoiceid,i.pretax,i.tax" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "having " & vbCrLf
    s = s & "(" & vbCrLf
    s = s & "   --ExceedsInvMax = InvoiceTotal > InvoiceMax" & vbCrLf
    s = s & "   case when i.pretax+i.tax > " & DbQuote(num, App.InvoiceMax) & " then 1 else 0 end = 1" & vbCrLf
    s = s & "OR" & vbCrLf
    s = s & "   --ExceedsOvrMax = PO>0 and Variance > OverrideMax" & vbCrLf
    s = s & "   case when i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > 0 " & vbCrLf
    s = s & "        and sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > " & DbQuote(num, App.InvOverrideMax) & vbCrLf
    s = s & "        then 1 else 0 end =1" & vbCrLf
    s = s & "OR" & vbCrLf
    s = s & "   --ExceedsOvrPct = PO>0 and Variance > PO * OverridePcnt" & vbCrLf
    s = s & "  case when i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > 0 " & vbCrLf
    s = s & "       and sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end) > (i.pretax+i.tax - sum(case when isnull(d.commitment,'')='' then d.pretax+d.tax else 0 end)) * " & DbQuote(num, App.InvOverridePcnt) & "/100" & vbCrLf
    s = s & "       then 1 else 0 end =1" & vbCrLf
    s = s & ")" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    ExceedsLimitDB = Not rs.EOF

End Function
Private Property Get InvoiceTotal(Optional Commitment, Optional CommitmentItem, Optional Job, Optional Extra, Optional Phase, Optional Category, Optional Pretax As Boolean) As Double
    'return amount of current invoice in job/extra/phase/cat
    Dim Row       As Long
    Dim total     As Double
    With gEnterDist
        If Not IsMissing(Commitment) Then
            For Row = 1 To .Rows - 2
                If Not .RowHidden(Row) Then
                    If .TextMatrix(Row, .ColIndex("Commitment")) = Commitment Then
                        If .TextMatrix(Row, .ColIndex("CommitmentItem")) = CommitmentItem Then
                            If Pretax Then
                                total = total + .ValueMatrix(Row, .ColIndex("PreTax"))
                            Else
                                total = total + .ValueMatrix(Row, .ColIndex("PreTax")) + .ValueMatrix(Row, .ColIndex("Tax"))
                            End If
                        End If
                    End If
                End If
            Next
        Else
            For Row = 1 To .Rows - 2
                If Not .RowHidden(Row) Then
                    If .TextMatrix(Row, .ColIndex("Job")) = Job Then
                        If .TextMatrix(Row, .ColIndex("Extra")) = Extra Then
                            If .TextMatrix(Row, .ColIndex("Phase")) = Phase Then
                                If (Category = "" Or .TextMatrix(Row, .ColIndex("Category")) = Category) Then
                                    If Pretax Then
                                        total = total + .ValueMatrix(Row, .ColIndex("PreTax"))
                                    Else
                                        total = total + .ValueMatrix(Row, .ColIndex("PreTax")) + .ValueMatrix(Row, .ColIndex("Tax"))
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End With
    InvoiceTotal = total
End Property

Private Sub VerifyPostedInvoices()
    Dim rs  As Recordset
    Dim rs1 As Recordset
    Dim s  As String
    Screen.MousePointer = vbHourglass
    
    If HFApp.Options(AccountingSystem) <> asTimberline Then
        s = ""
        s = s & "update invoices set status='Posted' where status='Exported'" & vbCrLf
        Call HFApp.SqlExec(s)
        Exit Sub
    End If
    
    s = ""
    s = s & "select InvoiceID,Vendor,Invoice,invoiceDate" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & " where DivisionID =" & HFApp.DivisionID & " and status='Exported'" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        s = ""
        s = s & "select * " & vbCrLf
        s = s & "  from new_api_record_1" & vbCrLf
        s = s & " where oivnd=" & DbQuote(Str, "" & rs("Vendor")) & vbCrLf
        s = s & "   and oiinv=" & DbQuote(Str, "" & rs("Invoice"), , True) & vbCrLf
        On Error Resume Next
        Set rs1 = HFApp.SqlExec(s, dbAccountingDictionary)
        If err.Number <> 0 Or rs1.EOF Then
            On Error GoTo 0
            s = ""
            s = s & "select * " & vbCrLf
            s = s & "  from master_apm_record_1" & vbCrLf
            s = s & " where oivnd=" & DbQuote(Str, "" & rs("Vendor")) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, "" & rs("Invoice"), , True) & vbCrLf
            's = s & "   and oidate=" & DbQuote(Date, "" & rs("Invoicedate")) & vbCrLf
            Set rs1 = HFApp.SqlExec(s, dbAccountingDictionary)
            If rs1.EOF Then
                s = "update invoices set status='Rejected' where invoiceid=" & rs("InvoiceID")
                Call HFApp.SqlExec(s)
            Else
                s = "update invoices set BatchNumber=" & DbQuote(Str, "" & rs1("oisessn"), , True) & ",status='Posted' where invoiceid=" & rs("InvoiceID")
                Call HFApp.SqlExec(s)
                Call SetCommitmentCompleted(rs("InvoiceID"))
            End If
        Else
            s = "update invoices set BatchNumber=" & DbQuote(Str, "" & rs1("oisessn"), , True) & ",status='Posted' where invoiceid=" & rs("InvoiceID")
            Call HFApp.SqlExec(s)
            Call SetCommitmentCompleted(rs("InvoiceID"))
        End If
        rs.MoveNext
    Wend
    Screen.MousePointer = vbDefault
End Sub







Private Function CleanFileName(s As String) As String
    s = Replace(s, "\/:*?<>|", "")
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, ":", "")
    s = Replace(s, "*", "")
    s = Replace(s, "?", "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "|", "")
    s = Replace(s, """", "")
    CleanFileName = s
End Function
Private Function DocumentPath(Vendor As String) As String
    Dim s As String
   
    s = PathAppend(App.Options(DocumentRoot), CleanFileName(Vendor))
    Call CreatePath("Document", s)
    DocumentPath = s
    
End Function

Private Function DocumentFilename(Vendor As String, Invoice As String, AllFiles As Boolean)
    Dim path As String
    Dim File As String
    Dim s    As String
    Dim i    As Long
    
    
    'remove illegal char
    File = Invoice
    File = Replace(File, "\/:*?<>|", "")
    File = Replace(File, "\", "")
    File = Replace(File, "/", "")
    File = Replace(File, ":", "")
    File = Replace(File, "*", "")
    File = Replace(File, "?", "")
    File = Replace(File, "<", "")
    File = Replace(File, ">", "")
    File = Replace(File, "|", "")
    File = Replace(File, """", "")
    
    'C:\Invoices\Vendor\Invoice
    path = PathAppend(DocumentPath(Vendor), File)
    
    
    If AllFiles Then
        DocumentFilename = path & "*.tif"
    Else
        'add page number
        s = Dir(path & "*")
        While s <> ""
            i = i + 1
            s = Dir
        Wend
        path = path & "_" & i + 1
        'add extension
        path = path & ".tif"
        DocumentFilename = path
    End If
    
End Function


Private Sub MailDocuments(Hold As Boolean)
On Error Resume Next
    Dim s        As String
    Dim File     As String
    Dim path     As String
    Dim email    As String
    
    path = DocumentFilename(txtVendor.Text, txtInvoice.Text, True)
    File = Dir(path)
    s = ""
    While File <> ""
        s = s & ";" & PathAppend(FilePath(path), File)
        File = Dir
    Wend
    File = Mid(s, 2)
    
    email = HFApp.SqlExec("SELECT Email FROM user_manager WHERE User_ID=" & DbQuote(Str, Trim(cboApprover.Text)))(0)
    If email = "" Then email = cboApprover.Text
    
On Error GoTo eh
    If Hold Then
        Call HFApp.SendMail(True, email, "", ReplaceEmailParms(App.Options(EmailSubject)), ReplaceEmailParms(App.Options(EmailMessage)), File)
    Else
        Call HFApp.SendMail(True, " ", "", "Invoice " & txtInvoice.Text, "", File)
    End If
Exit Sub
eh: MsgBox err.Description, vbExclamation, App.ProductName
End Sub

Private Function ReplaceEmailParms(s As String) As String
    s = Replace(s, vbQuote, "")
    s = Replace(s, "<%Vendor%>", txtVendor.tag)
    s = Replace(s, "<%VendorName%>", txtVendor.Text)
    s = Replace(s, "<%Invoice%>", txtInvoice.Text)
    s = Replace(s, "<%Job%>", txtJob.Text)
    s = Replace(s, "<%Reason%>", txtComments.Text)
    ReplaceEmailParms = s
End Function
Public Function SendMail(Preview As Boolean, SendTo As String, CCTo As String, Subject As String, Body As String, Attachments As String) As Boolean
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim s As String
    Dim Msg   As MAPIMessage
    Dim rec() As MapiRecip
    Dim att() As MapiFile
    
    'build recipients
    i = Parse(SendTo, , ";") - 1
    ReDim rec(i)
    For i = 0 To UBound(rec)
        s = Parse(SendTo, i + 1, ";")
        rec(i).RecipClass = MAPI_TO
        rec(i).Name = s
    Next
    
    'build cc list
    If CCTo <> "" Then
        j = Parse(CCTo, , ";") - 1
        ReDim Preserve rec(i + j)
        For k = 0 To j
            s = Parse(CCTo, k + 1, ";")
            rec(i + k).RecipClass = MAPI_CC
            rec(i + k).Name = s
        Next
    End If
    
    'build attachments
    i = Parse(Attachments, , ";") - 1
    ReDim att(i)
    For i = 0 To UBound(att)
        s = Trim(Parse(Attachments, i + 1, ";"))
        att(i).Position = i + 1
        att(i).PathName = s
    Next
    
    
    
    'build message
    Msg.FileCount = 0
    Msg.NoteText = Body
    Msg.Subject = Subject
    Msg.RecipCount = UBound(rec) + 1
    Msg.FileCount = IIf(Attachments = "", 0, UBound(att) + 1)
    
    If Len(Body) < Msg.FileCount Then
        Msg.NoteText = Msg.NoteText & Space(Msg.FileCount + 1)
    End If
    
    Select Case MAPISendMail(0, 0, Msg, rec, att, MAPI_LOGON_UI + IIf(Preview, MAPI_DIALOG, 0), 0)
        Case MAPI_E_USER_ABORT:                SendMail = False
        Case MAPI_E_FAILURE:                   Call err.Raise(5, , "MAPI_E_FAILURE")
        Case MAPI_E_LOGIN_FAILURE:             SendMail = False
        Case MAPI_E_LOGON_FAILURE:             SendMail = False
        Case MAPI_E_DISK_FULL:                 Call err.Raise(5, , "MAPI_E_DISK_FULL")
        Case MAPI_E_INSUFFICIENT_MEMORY:       Call err.Raise(5, , "MAPI_E_INSUFFICIENT_MEMORY")
        Case MAPI_E_BLK_TOO_SMALL:             Call err.Raise(5, , "MAPI_E_INSUFFICIENT_MEMORY")
        Case MAPI_E_TOO_MANY_SESSIONS:         Call err.Raise(5, , "MAPI_E_TOO_MANY_SESSIONS")
        Case MAPI_E_TOO_MANY_FILES:            Call err.Raise(5, , "MAPI_E_TOO_MANY_FILES")
        Case MAPI_E_TOO_MANY_RECIPIENTS:       Call err.Raise(5, , "MAPI_E_TOO_MANY_RECIPIENTS")
        Case MAPI_E_ATTACHMENT_NOT_FOUND:      Call err.Raise(5, , "MAPI_E_ATTACHMENT_NOT_FOUND")
        Case MAPI_E_ATTACHMENT_OPEN_FAILURE:   Call err.Raise(5, , "MAPI_E_ATTACHMENT_OPEN_FAILURE")
        Case MAPI_E_ATTACHMENT_WRITE_FAILURE:  Call err.Raise(5, , "MAPI_E_ATTACHMENT_WRITE_FAILURE")
        Case MAPI_E_UNKNOWN_RECIPIENT:         Call err.Raise(5, , "MAPI_E_UNKNOWN_RECIPIENT")
        Case MAPI_E_BAD_RECIPTYPE:             Call err.Raise(5, , "MAPI_E_BAD_RECIPTYPE")
        Case MAPI_E_NO_MESSAGES:               Call err.Raise(5, , "MAPI_E_NO_MESSAGES")
        Case MAPI_E_INVALID_MESSAGE:           Call err.Raise(5, , "MAPI_E_INVALID_MESSAGE")
        Case MAPI_E_TEXT_TOO_LARGE:            Call err.Raise(5, , "MAPI_E_TEXT_TOO_LARGE")
        Case MAPI_E_INVALID_SESSION:           Call err.Raise(5, , "MAPI_E_INVALID_SESSION")
        Case MAPI_E_TYPE_NOT_SUPPORTED:        Call err.Raise(5, , "MAPI_E_TYPE_NOT_SUPPORTED")
        Case MAPI_E_AMBIGUOUS_RECIPIENT:       Call err.Raise(5, , "MAPI_E_AMBIGUOUS_RECIPIENT")
        Case MAPI_E_AMBIG_RECIP:               Call err.Raise(5, , "MAPI_E_AMBIG_RECIP")
        Case MAPI_E_MESSAGE_IN_USE:            Call err.Raise(5, , "MAPI_E_MESSAGE_IN_USE")
        Case MAPI_E_NETWORK_FAILURE:           Call err.Raise(5, , "MAPI_E_NETWORK_FAILURE")
        Case MAPI_E_INVALID_EDITFIELDS:        Call err.Raise(5, , "MAPI_E_INVALID_EDITFIELDS")
        Case MAPI_E_INVALID_RECIPS:            Call err.Raise(5, , "MAPI_E_INVALID_RECIPS")
        Case MAPI_E_NOT_SUPPORTED:             SendMail = False
        Case Else:                             SendMail = True
    End Select

End Function


Private Sub GetInvoiceNumber()
    Dim s  As String
    Dim r As Long
    Dim CommitmentNumber As String
    Dim InvoiceNumber As String
    
    If txtInvoice.Text <> "" Then Exit Sub
    
    
    'get first commitment referenced on invoice
    With gEnterDist
        For r = 1 To .Rows - 1
            If .TextMatrix(r, .ColIndex("Commitment")) <> "" Then
                CommitmentNumber = .TextMatrix(r, .ColIndex("Commitment"))
                Exit For
            End If
        Next
    End With
    If CommitmentNumber = "" Then Exit Sub
    
    
    InvoiceNumber = ""
    
    'get next invoice number for this commitment
    s = ""
    s = s & "select invoice " & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & " where DivisionID =" & HFApp.DivisionID & vbCrLf
    s = s & "   and vendor=" & DbQuote(Str, txtVendor.tag) & vbCrLf
    s = s & "   and retainageinvoice=0" & vbCrLf
    s = s & "   and invoice like " & DbQuote(Str, CommitmentNumber & "%") & vbCrLf
    s = s & "order by 1 desc"
    On Error Resume Next
    InvoiceNumber = Max(InvoiceNumber, HFApp.SqlExec(s)(0))
    On Error GoTo 0
    
    If InvoiceNumber = "" Then
        InvoiceNumber = CommitmentNumber
    Else
        InvoiceNumber = CommitmentNumber & "." & Val(Mid(InvoiceNumber, Len(CommitmentNumber) + 2)) + 1
    End If
    
    mLoading = True
    txtInvoice.Text = InvoiceNumber
    mLoading = False
    txtVendor.Enabled = False
    mLoading = True
    lblVendor.Enabled = False
    txtInvoice.Enabled = False
    lblInvoice.Enabled = False
    mLoading = False
        
End Sub

Private Sub SetCommitmentCompleted(InvoiceID As Long)
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim csv As String
    
    If Not App.Options(AutomaticallyClosePO) Then Exit Sub
    
    ' get list of commitments referenced by this invoice
    s = ""
    s = s & "SELECT Commitment" & vbCrLf
    s = s & "  FROM InvoiceItems" & vbCrLf
    s = s & " WHERE DivisionID =" & HFApp.DivisionID & " and IFNULL(Commitment,'')<>''" & vbCrLf
    s = s & "   AND InvoiceID=" & DbQuote(num, InvoiceID)
    Set rs = HFApp.SqlExec(s)
    
    csv = ""
    While Not rs.EOF
        csv = csv & "," & DbQuote(Str, "" & rs(0))
        rs.MoveNext
    Wend
    csv = Mid(csv, 2)
    
    'close commitments with $0.00 remaining
    If csv <> "" Then
        s = ""
        s = s & "update master_jcm_record_12" & vbCrLf
        s = s & "   set sclosed=1" & vbCrLf
        s = s & " where sclosed=0" & vbCrLf
        s = s & "   and samt+sapprco-samtinv=0" & vbCrLf
        s = s & "   and sub in(" & csv & ")" & vbCrLf
        Call HFApp.SqlExec(s, dbAccountingDictionary)
    End If
Exit Sub
eh: Call ErrHandler(SRCFILE & "SetCommitmentCompleted")
End Sub

Private Function ArePOsCompleted() As Boolean
    Dim s As String
    Dim Pos As String
    Dim r As Long
    Dim rs As Recordset
    
    If App.Options(SchedulingUsage) = DoNotUse Then
        ArePOsCompleted = True
    Else
    
        With gEnterDist
        Pos = ""
        For r = 1 To .Rows - 2
            If .TextMatrix(r, .ColIndex("Commitment")) <> "" Then
                Pos = Pos & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Commitment")))
            End If
        Next
        Pos = Mid(Pos, 2)
        End With
        
        If Pos = "" Then
            ArePOsCompleted = True
            Exit Function
        End If
        
        s = ""
        s = s & "select p.approveddate,m.pmname,m.email" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "left outer join tblJobs j on(p.job=j.job_no)" & vbCrLf
        s = s & "left outer join tblProjectmanager m on(m.pm=j.pm)" & vbCrLf
        s = s & "left outer join tblPOIndex x on(p.divisionid=x.divisionid and p.poindex=x.poindex)" & vbCrLf
        s = s & "where isfieldpo=0 " & vbCrLf
        s = s & "and x.requirespaymentapproval=1" & vbCrLf
        s = s & "and p.approveddate is null" & vbCrLf
        s = s & "and p.ponumber in(" & Pos & ")"
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        ArePOsCompleted = rs.EOF
        
    End If
End Function


Public Function SchedulingIsFinished(Commitment As String, Optional PayPoint As Integer) As Boolean

    Dim s As String
    Dim rs As Recordset
    If App.Options(SchedulingUsage) = DoNotUse Then
        SchedulingIsFinished = True
    Else
        s = ""
        s = s & "select p.approveddate,m.pmname,m.email" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "left outer join tblJobs j on(p.job=j.job_no)" & vbCrLf
        s = s & "left outer join tblProjectmanager m on(m.pm=j.pm)" & vbCrLf
        s = s & "left outer join tblPOIndex x on(p.poindex=x.poindex and p.divisionid=x.divisionid)" & vbCrLf
        s = s & "where isfieldpo=0 " & vbCrLf
        s = s & "and x.requirespaymentapproval=1" & vbCrLf
        If PayPoint = 0 Then
            s = s & "and p.approveddate is null" & vbCrLf
        Else
            s = s & "and p.paypoint" & PayPoint & "approveddate is null" & vbCrLf
        End If
        s = s & "and p.ponumber=" & DbQuote(Str, Commitment) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        
        If rs.EOF Then
            SchedulingIsFinished = True
        Else
            If App.Options(SchedulingUsage) = RestrictIfIncomplete Then
                If "" & rs(0) = "" Then
                    s = App.Options(Caption_Commitment) & " '" & Trim(Commitment) & "' " & IIf(PayPoint = 0, "", "pay point " & PayPoint) & " has not been completed and" & vbCrLf & "cannot be invoiced. Contact the project manager" & vbCrLf & "for further instructions."
                Else
                    s = App.Options(Caption_Commitment) & " '" & Trim(Commitment) & "' " & IIf(PayPoint = 0, "", "pay point " & PayPoint) & " has not been completed and" & vbCrLf & "cannot be invoiced. Contact " & rs(1) & vbCrLf & " for further instructions."
                End If
                Call MsgBox(s, vbInformation, App.ProductName)
                SchedulingIsFinished = False
            Else 'warn or hold
                If "" & rs(0) = "" Then
                    s = App.Options(Caption_Commitment) & " '" & Trim(Commitment) & "' " & IIf(PayPoint = 0, "", "pay point " & PayPoint) & " has not been completed. Contact" & vbCrLf & "the project manager for further information." & vbCrLf & vbCrLf & "Do you want to invoice it anyway?"
                Else
                    s = App.Options(Caption_Commitment) & " '" & Trim(Commitment) & "' " & IIf(PayPoint = 0, "", "pay point " & PayPoint) & " has not been completed. Contact" & vbCrLf & rs(1) & " for further information." & vbCrLf & vbCrLf & "Do you want to invoice it anyway?"
                End If
                SchedulingIsFinished = vbYes = MsgBox(s, vbYesNo + vbQuestion, App.ProductName)
            End If
        End If
    End If
End Function



Private Function TranslateBackChargeVendorTypes() As String
    Dim s As String
    s = App.Options.Value(BackChargeVendorTypes)
    s = Replace(s, "1", "'Subcontractor'")
    s = Replace(s, "2", "'Supplier'")
    s = Replace(s, "3", "'Other'")
    s = Replace(s, "4", "'Equip Supplier'")
    TranslateBackChargeVendorTypes = s
End Function
Private Sub CheckBackChargePOs(Optional Vendor As String)
    Dim s As String
    Dim i As Long
    
    'not enabled
    If Not TimberlineAccounting Then Exit Sub
    
    
    If Not App.Options(EnableBackChargePOs) Then Exit Sub
    
    
    If Vendor <> "" Then
        'check if there are any
        s = ""
        s = s & "select count(*)" & vbCrLf
        s = s & "from master_apm_record_9 v" & vbCrLf
        s = s & "     inner join master_jcm_record_12 c on (c.smsalp1=ltrim(v.vendor))" & vbCrLf
        s = s & "where v.vtype in(" & TranslateBackChargeVendorTypes & ")" & vbCrLf
        s = s & "  and 0=c.smsccl" & App.Options(BackChargeStatusItem) & vbCrLf
        s = s & "  and v.vendor=" & DbQuote(Str, Vendor) & vbCrLf
        If Val("" & HFApp.SqlExec(s, dbAccountingDictionary)(0)) = 0 Then Exit Sub
        
        'ask if they want to process them
        If vbCancel = MsgBox("This vendor has pending back charges." & vbCrLf & vbCrLf & "Do you want to process them now?", vbQuestion + vbOKCancel) Then Exit Sub
    End If
    
    'pick which ones
    s = ""
    s = s & "select ltrim(v.vendor) Vendor" & vbCrLf
    s = s & "      ,v.vname  ""Back Charge To""" & vbCrLf
    s = s & "      ,c.samt+c.sapprco*(100+" & Val(App.Options(BackChargeMarkup)) & ")/100 Amount" & vbCrLf
    s = s & "      ,s.vname  Supplier" & vbCrLf
    s = s & "      ,ltrim(c.sub) " & App.Options(Caption_Commitment) & vbCrLf
    s = s & "      ,c.sdate  Date" & vbCrLf
    s = s & "      ,c.sdesc  Description" & vbCrLf
    s = s & "from master_apm_record_9 v" & vbCrLf
    s = s & "     inner join master_jcm_record_12 c on (c.smsalp1=v.vendor)" & vbCrLf
    s = s & "     inner join master_apm_record_9 s on (c.svendor=s.vendor)" & vbCrLf
    s = s & "where v.vtype in(" & TranslateBackChargeVendorTypes & ")" & vbCrLf
    s = s & "  and 0=c.smsccl" & App.Options(BackChargeStatusItem) & vbCrLf
    If Vendor <> "" Then
        s = s & "  and v.vendor=" & DbQuote(Str, Vendor) & vbCrLf
    End If
    If Not FPickBackCharges.Choose(s, , "Vendor", dteAccounting.Value) Then Exit Sub
    
    'make credit memos
    s = ""
    For i = 1 To FPickBackCharges.SelectedItems
        s = s & "," & DbQuote(Str, FPickBackCharges.SelectedItem(App.Options(Caption_Commitment), i))
    Next
    s = Mid(s, 2)
    
    If s <> "" Then Call CreateBackCharges(s, FPickBackCharges.SelectedItem("AccountingDate"))
    
    
End Sub


Private Sub CreateBackCharges(CommitmentsCSV As String, AccountingDate As String)
On Error GoTo eh
    Dim s    As String
    Dim rs   As Recordset
    Dim i As Long
    Dim Commitment As String
    Dim InvoiceID As Long
    Dim InvoiceIDs As String
    
    For i = 1 To Parse(CommitmentsCSV)
        Commitment = Parse(CommitmentsCSV, i)
        If Commitment <> "" Then
        
            s = ""
            s = s & "select v.Vendor " & vbCrLf
            s = s & "      ,v.vtype   VendorType" & vbCrLf
            s = s & "      ,v.vname   VendorName" & vbCrLf
            s = s & "      ,c.sdate   ReceivedDate" & vbCrLf
            s = s & "      ,c.sdate   InvoiceDate" & vbCrLf
            s = s & "      ,c.sub     Commitment" & vbCrLf
            s = s & "      ,c.sdesc   CommitmentDesc" & vbCrLf
            s = s & "      ,d.item    CommitmentItem" & vbCrLf
            s = s & "      ,c.svendor CommitmentVendor" & vbCrLf
            s = s & "      ,d.ijob    Job" & vbCrLf
            s = s & "      ,j.jdesc   JobDesc" & vbCrLf
            s = s & "      ,d.iextra  Extra" & vbCrLf
            s = s & "      ,x.xdesc   ExtraDesc" & vbCrLf
            s = s & "      ,d.iphase  Phase" & vbCrLf
            s = s & "      ,p.pdesc   PhaseDesc" & vbCrLf
            s = s & "      ,d.icat    Category" & vbCrLf
            s = s & "      ,ct.cdesc  CategoryDesc" & vbCrLf
            s = s & "      ,d.itxgrp  TaxGroup" & vbCrLf
            s = s & "      ,t.gdesc   TaxGroupDesc" & vbCrLf
            s = s & "      ,t.grate   TaxRate" & vbCrLf
            s = s & "      ,round((d.iamt-d.itxamt+d.iappcoa-d.iaptcoa)*(100+" & Val(App.Options(BackChargeMarkup)) & ")/100,2) Pretax" & vbCrLf
            s = s & "      ,round((d.iamt-d.itxamt+d.iappcoa-d.iaptcoa)*(100+" & Val(App.Options(BackChargeMarkup)) & ")/100,2) * ifnull(t.grate,0)/100 Tax" & vbCrLf
            s = s & "      ,d.idesc   Description" & vbCrLf
            s = s & "      ,d.iunits  CommittedQuantity" & vbCrLf
            s = s & "      ,d.iuntcst CommittedUnitPrice" & vbCrLf
            s = s & "      ,d.iunits  InvoicedQuantity" & vbCrLf
            s = s & "      ,d.iuntcst InvoicedUnitPrice" & vbCrLf
            s = s & "from master_jcm_record_12 c" & vbCrLf
            s = s & "     inner join master_apm_record_9 v on (c.smsalp1=v.vendor)" & vbCrLf
            s = s & "     inner join master_jcm_record_13 d on(c.sub=d.isub)" & vbCrLf
            s = s & "     inner join master_jcm_record_1_1 j on(d.ijob=j.job)" & vbCrLf
            s = s & "     left outer join master_jcm_record_2 x on(d.ijob=x.xjob and d.iextra=x.extra)" & vbCrLf
            s = s & "     left outer join master_jcm_record_3 p on(d.ijob=p.pjob and d.iextra=p.pextra and d.iphase=p.phase)" & vbCrLf
            s = s & "     left outer join master_jcm_record_4 ct on(d.ijob=ct.cjob and d.iextra=ct.cextra and d.iphase=ct.cphase and d.icat=ct.cat)" & vbCrLf
            s = s & "     left outer join master_txm_record_2 t on (d.itxgrp=t.ggroup)" & vbCrLf
            s = s & "where smsccl" & App.Options(BackChargeStatusItem) & "=0" & vbCrLf
            s = s & "  and c.sub=" & Commitment & vbCrLf
            Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
            If Not rs.EOF Then
                
                'mark PO as processed
                s = ""
                s = s & "update master_jcm_record_12" & vbCrLf
                s = s & "   set sinscr=1" & vbCrLf
                s = s & " where sub=" & DbQuote(Str, "" & rs("Commitment"), , True)
                On Error Resume Next
                Call HFApp.SqlExec(s, dbAccountingDictionary)
                If err.Number <> 0 Then
                    On Error GoTo eh
                    Call MsgBox("Unable to access " & App.Options(Caption_Commitment) & ": " & Trim("" & rs("Commitment")) & vbCrLf & "Some one may be editting it.", vbInformation, App.ProductName)
                Else
                    On Error GoTo eh
                    'write invoice
                    s = ""
                    s = s & "insert into invoices(DivisionID,vendor,vendortype,vendorname,invoice,job,jobdesc,status"
                    s = s & ",receiveddate,invoicedate,accountingdate,description,ustmp,tstmp,dstmp) values(" & HFApp.DivisionID & "," & vbCrLf
                    s = s & " " & DbQuote(Str, "" & rs("vendor"), , True) & vbCrLf
                    s = s & "," & DbQuote(Str, "" & rs("vendortype")) & vbCrLf
                    s = s & "," & DbQuote(Str, "" & rs("vendorname"), , True) & vbCrLf
                    s = s & "," & DbQuote(Str, "BC:" & LTrim(rs("Commitment")), , , 15) & vbCrLf
                    s = s & "," & DbQuote(Str, "" & rs("Job"), , True) & vbCrLf
                    s = s & "," & DbQuote(Str, "" & rs("JobDesc")) & vbCrLf
                    s = s & "," & DbQuote(Str, "Pending") & vbCrLf
                    s = s & "," & DbQuote(Date, "" & rs("receiveddate")) & vbCrLf
                    s = s & "," & DbQuote(Date, "" & rs("invoicedate")) & vbCrLf
                    s = s & "," & DbQuote(Date, AccountingDate) & vbCrLf
                    s = s & "," & DbQuote(Date, "" & rs("CommitmentDesc"), , True) & vbCrLf
                    s = s & "," & DbQuote(Str, HFApp.LoginID) & vbCrLf
                    s = s & "," & DbQuote(Time, Now()) & vbCrLf
                    s = s & "," & DbQuote(Date, Now()) & ")"
                    Call HFApp.SqlExec(s, dbHomeFront)
                    InvoiceID = HFApp.SqlIdentity("invoices", dbHomeFront)
                    InvoiceIDs = InvoiceIDs & "," & InvoiceID
                    
                    'write invoice nonzero distribution
                    While Not rs.EOF
                        If Val("" & rs("Pretax")) <> 0 Then
                            s = ""
                            s = s & "insert into invoiceitems(DivisionID,invoiceid,vendor,invoice,job,jobdesc,extra,extradesc,phase,phasedesc,category,categorydesc"
                            s = s & ",taxgroup,taxgroupdesc,taxrate,pretax,tax,description,ustmp,tstmp,dstmp,committedquantity,committedunitprice,invoicedquantity,invoicedunitprice) " & vbCrLf
                            s = s & "values(" & HFApp.DivisionID & "," & vbCrLf
                            s = s & " " & DbQuote(num, InvoiceID) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("vendor"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "BC:" & LTrim(rs("Commitment")), , , 15) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("Job"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("JobDesc")) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("Extra"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("ExtraDesc"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("Phase"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("PhaseDesc"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("Category"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("CategoryDesc"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("TaxGroup"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("TaxGroupDesc"), , True) & vbCrLf
                            s = s & "," & DbQuote(num, "" & rs("TaxRate")) & vbCrLf
                            s = s & "," & DbQuote(num, -1 * Val("" & rs("Pretax"))) & vbCrLf
                            s = s & "," & DbQuote(num, -1 * Val("" & rs("Tax"))) & vbCrLf
                            s = s & "," & DbQuote(Str, "" & rs("Description"), , True) & vbCrLf
                            s = s & "," & DbQuote(Str, HFApp.LoginID) & vbCrLf
                            s = s & "," & DbQuote(Time, Now()) & vbCrLf
                            s = s & "," & DbQuote(Date, Now()) & vbCrLf
                            s = s & "," & DbQuote(num, "" & rs("CommittedQuantity")) & vbCrLf
                            s = s & "," & DbQuote(num, "" & rs("CommittedUnitPrice")) & vbCrLf
                            s = s & "," & DbQuote(num, "" & rs("InvoicedQuantity")) & vbCrLf
                            s = s & "," & DbQuote(num, "" & rs("InvoicedUnitPrice")) & ")"
                            Call HFApp.SqlExec(s, dbHomeFront)
                        End If
                        rs.MoveNext
                    Wend
                    
                End If
                    
            End If
            
        End If
    Next
    
    If InvoiceIDs <> "" Then
        'update invoice totals
        s = ""
        s = s & "update invoices " & vbCrLf
        s = s & "set pretax=(select sum(pretax) from invoiceitems where DivisionID =" & HFApp.DivisionID & " and invoices.invoiceid=invoiceitems.invoiceid)" & vbCrLf
        s = s & "   ,tax=(select sum(tax) from invoiceitems where DivisionID =" & HFApp.DivisionID & " and invoices.invoiceid=invoiceitems.invoiceid)" & vbCrLf
        s = s & "where DivisionID =" & HFApp.DivisionID & " and invoiceid in(" & Mid(InvoiceIDs, 2) & ")" & vbCrLf
        Call HFApp.SqlExec(s, dbHomeFront)
    End If
    
    
    
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "CreateBackCharges")
End Sub




Private Sub PostInvoices(Optional UIDs As String, Optional InvoiceIDs As String)
    
    'both parameters are csv lists
    Dim s        As String
    Dim filter   As String
    Dim Count    As Long
    Dim rc       As Long
    
    If Not App.PostInvoices Then
        MsgBox "Your security settings do not allow you to post invoices." & vbCrLf & "Contact an administrator to change your permissions.", vbInformation, App.ProductName
        Exit Sub
    End If
    
    'build qry filter and ask permission
    Select Case True
        Case UIDs <> ""
            filter = "   and Invoices.status in('SiteApproved','Approved') and Invoices.ustmp in(" & UIDs & ")"
            If MsgBox("Post all approved invoices entered by " & UIDs & "?", vbQuestion + vbYesNo, App.ProductName) <> vbYes Then Exit Sub
        Case InvoiceIDs <> ""
            filter = "   and Invoices.InvoiceID in(" & InvoiceIDs & ")"
            If Parse(InvoiceIDs) = 1 Then
                If MsgBox("Post this invoice?", vbQuestion + vbYesNo, App.ProductName) <> vbYes Then Exit Sub
            Else
                If MsgBox("Post these " & Parse(InvoiceIDs) & " invoices?", vbQuestion + vbYesNo, App.ProductName) <> vbYes Then Exit Sub
            End If
        Case Else
            filter = "   and Invoices.status in('SiteApproved','Approved')"
            If MsgBox("Post ALL invoices?", vbQuestion + vbYesNo, App.ProductName) <> vbYes Then Exit Sub
    End Select
      
   
    s = ""
    s = s & "select count(*)" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "join tblvendors v on invoices.divisionid=v.divisionid and invoices.vendor=v.vendor_id and isnull(v.istbd,0)=0" & vbCrLf
    s = s & "where invoices.DivisionID = " & HFApp.DivisionID & " and month(invoices.accountingdate)>month(getdate()) and year(invoices.accountingdate) = year(getdate()) " & vbCrLf
    s = s & filter
    
    'this will return the number of invoices in this batch with an accounting date of last month
    Count = HFApp.SqlExec(s, dbHomeFront)(0)
    If App.Options(WarnIfAcctDateNotCurrent) And Count >= 1 Then
        rc = MsgBox(Count & " of these invoices have accounting date which are not current." & vbCrLf & vbCrLf & "Do you want to use todays date for these " & Count & " invoices?", vbYesNoCancel + vbExclamation, App.ProductName)
        Select Case rc
            Case vbCancel: Exit Sub
            Case vbNo:     'continue
            Case vbYes
                s = ""
                s = s & "update invoices" & vbCrLf
                s = s & "set accountingdate = getdate()" & vbCrLf
                s = s & "from invoices join tblvendors v on invoices.divisionid=v.divisionid and invoices.vendor=v.vendor_id and isnull(v.istbd,0)=0" & vbCrLf
                s = s & "where invoices.DivisionID = " & HFApp.DivisionID & " and month(invoices.accountingdate)>month(getdate()) and year(invoices.accountingdate) = year(getdate()) " & vbCrLf
                s = s & filter
                Call HFApp.SqlExec(s, dbHomeFront)
        End Select
    End If
    
    s = filter
    If UIDs = "" Then
        If Not InvoiceApproved("", s) Then
            MsgBox "Invoice has not been approved for posting.", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    
    Call PostJobsToAccounting(filter)
      
    Select Case HFApp.Options(AccountingSystem)
        Case asTimberline:
            'this was for macroless postings which were never completed
            'If HFApp.Options.ValueByName("DisableSageMacaroni") <> "True" Then
            '    Call PostToSage300(filter)
            'Else
                Call PostToTimberline(filter)
            'End If
            
            
        Case asIntacct:          Call SendInvoicesToIntacct(filter)
        Case asSimply:           Call PostToSimply(filter)
        Case asQuickBooks:       Call PostToQuickBooks(filter)
        Case asQuickBooksOnline: Call PostToQuickBooksOnlineUS(filter)
        Case asXero:             Call PostToXero(filter)
        Case asMasterBuilder:    Call PostToSage100(filter)
    End Select
    
    
    
    Screen.MousePointer = vbDefault
End Sub
Private Function PrePostValidate(WhereClause As String) As Boolean
On Error GoTo eh
    
    'dont do this anymore. it costs more than its worth. just let accounting reject the invoice..
    PrePostValidate = True
    Exit Function


'    Dim s As String
'    Dim rs As Recordset
'
'    'using timberline. cant validate.
'    If HFApp.Options(AccountingSystem) = asTimberline And App.Options(UseTimberlineAsPOSource) Then
'        PrePostValidate = True
'        Exit Function
'    End If
'
'    s = ""
'    s = s & "select distinct i.vendorname,i.invoice" & vbCrLf
'    s = s & "from invoices i" & vbCrLf
'    s = s & "join InvoiceValidationErrors e on i.invoiceid=e.invoiceid" & vbCrLf
'    s = s & " where i.DivisionID = " & HFApp.DivisionID & vbCrLf
'    s = s & WhereClause & vbCrLf
'    s = s & "order by 1,2" & vbCrLf
'    Set rs = HFApp.SqlExec(s)
'    If rs.EOF Then
'        'no errors found. validation passes
'        PrePostValidate = True
'    Else
'        s = "Invoices are incomplete or incorrect and cannot be posted." & vbCrLf & vbCrLf
'        While rs.EOF
'            s = s & rs(0) & " invoice number " & rs(1) & vbCrLf
'        Wend
'        MsgBox s, vbExclamation, App.ProductName
'        PrePostValidate = False
'    End If

    
Exit Function
eh: Call ErrHandler(SRCFILE & "PrePostValidate")
End Function
Private Sub PostToTimberline(WhereClause As String)
On Error GoTo eh
    Dim prevJob   As String
    Dim prevExtra As String

    Dim invImp   As Integer
    Dim invHst   As Integer
    Dim estImp   As Integer
    Dim estHst   As Integer
    Dim s        As String
    Dim rs       As Recordset
    Dim lastID   As Long
    Dim TLDateFormat As String
    
    Dim postedInvoiceIDs As String
    
    TLDateFormat = HFApp.Options.ValueByName("TimberlineDateFormat")
    If TLDateFormat = "" Then
        TLDateFormat = "mm\/dd\/yyyy"
    End If
    
    'delete the print file - not necessary except to cause error if it's in use.
    If PathExists(App.Options(ImportPrintFile)) Then Call Kill(App.Options(ImportPrintFile))
    
    'open the files
    Screen.MousePointer = vbHourglass
    Call CreatePath("Estimate Import File", FilePath(App.Options(ImportEstimateFile)))
    Call CreatePath("Invoice Import File", FilePath(App.Options(ImportInvoiceFile)))
    Call CreatePath("History Folder", App.Options(ImportHistoryFolder))
    invImp = FreeFile: Open App.Options(ImportInvoiceFile) For Output Access Write Lock Read Write As invImp
    invHst = FreeFile: Open App.Options(ImportHistoryFolder) & "\" & Format(Now(), "yyyy.mm.dd.hh.mm.ss") & ".txt" For Output Access Write Lock Read Write As invHst
    estImp = FreeFile: Open App.Options(ImportEstimateFile) For Output Access Write Lock Read Write As estImp
    estHst = FreeFile: Open App.Options(ImportHistoryFolder) & "\" & Format(Now(), "yyyy.mm.dd.hh.mm.ss") & ".jce" For Output Access Write Lock Read Write As estHst
    
    'open invoices query
    s = ""
    s = s & "select Invoices.InvoiceID" & vbCrLf
    s = s & "      ,Invoices.Vendor" & vbCrLf
    s = s & "      ,tblvendors.tradetype VendorType" & vbCrLf
    s = s & "      ,Invoices.Invoice" & vbCrLf
    s = s & "      ,left(Invoices.Description,30) InvoiceDescription" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode1" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode2" & vbCrLf
    s = s & "      ,Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & "      ,Invoices.Tax                   InvoiceTax" & vbCrLf
    s = s & "      ,Invoices.Discount              InvoiceDiscount" & vbCrLf
    s = s & "      ,Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & "      ,Invoices.ReceivedDate          InvoiceReceivedDate" & vbCrLf
    s = s & "      ,Invoices.DiscountDate          InvoiceDiscountDate" & vbCrLf
    s = s & "      ,Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & "      ,Invoices.AccountingDate        InvoiceAccountingDate" & vbCrLf
    s = s & "      ,Invoices.VendorName            InvoiceVendorName" & vbCrLf
    s = s & "      ,Invoices.VendorAddr1           InvoiceVendorAddr1" & vbCrLf
    s = s & "      ,Invoices.VendorAddr2           InvoiceVendorAddr2" & vbCrLf
    s = s & "      ,Invoices.VendorCity            InvoiceVendorCity" & vbCrLf
    s = s & "      ,Invoices.VendorProv            InvoiceVendorProv" & vbCrLf
    s = s & "      ,Invoices.VendorPostal          InvoiceVendorPostal" & vbCrLf
    s = s & "      ,invoiceitems.JointPayee        " & vbCrLf
    s = s & "      ,invoiceitems.ItemID            ItemID" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentVendor  ItemCommitmentVendor" & vbCrLf
    s = s & "      ,invoiceitems.Commitment        ItemCommitment" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentItem    ItemCommitmentItem" & vbCrLf
    s = s & "      ,invoiceitems.Job               ItemJob" & vbCrLf
    s = s & "      ,invoiceitems.Extra             ItemExtra" & vbCrLf
    s = s & "      ,invoiceitems.Phase             ItemPhase" & vbCrLf
    s = s & "      ,invoiceitems.Category          ItemCategory" & vbCrLf
    s = s & "      ,invoiceitems.DebitAccount      ItemDebitAccount" & vbCrLf
    s = s & "      ,invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedQuantity  ItemQuantity" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedUnitPrice ItemUnitPrice" & vbCrLf
    
    s = s & "      ,case when isnull(InvoiceItems.wrapinsuranceexempt,0)=1 then 0 " & vbCrLf
    s = s & "            else isnull(invoiceitems.PreTax,0) * isnull(Invoices.MiscDeductionRate,0) end MiscDeductionAmt" & vbCrLf
    
    s = s & "      ,invoiceitems.PreTax + invoiceitems.Tax    ItemAmount" & vbCrLf
    s = s & "      ,invoiceitems.Tax               ItemTax" & vbCrLf
    s = s & "      ,round(invoiceitems.Retainage*(1+isnull(t.RetainageRate,0)/100),2)     ItemRetainage" & vbCrLf
    s = s & "      ,left(invoiceitems.Description,30)       ItemDescription" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "  left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "  left outer join tblvendors   on invoices.DivisionID=tblvendors.DivisionID and invoices.vendor=tblvendors.vendor_id" & vbCrLf
    s = s & "  left outer join taxgroups  t on invoiceitems.divisionid=t.divisionid and invoiceitems.taxgroup=t.taxgroup" & vbCrLf
    s = s & " where invoices.DivisionID = " & HFApp.DivisionID & " and 1=1" & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "order by Invoices.InvoiceID,invoiceitems.ItemID"
    
    Set rs = HFApp.SqlExec(s)
    lastID = -999
    
    While Not rs.EOF
        If lastID <> rs("InvoiceID") Then
            lastID = rs("InvoiceID")
            
            'write invoice header
            s = ""
            s = s & "APIF,"
            s = s & Quote("" & rs("Vendor")) & ","
            s = s & Quote("" & rs("Invoice")) & ","
            s = s & Quote("" & rs("InvoiceDescription"), , 30) & ","
            s = s & Format(rs("InvoiceAmount"), "###.00") & ","
            s = s & Format(rs("InvoiceTax"), "###.00") & ","
            s = s & Format(rs("InvoiceDiscount"), "###.00") & ","
            s = s & ","
            s = s & Format(rs("InvoiceDate"), App.Options(TimberlineDateFormat)) & ","
            s = s & Format(rs("InvoiceReceivedDate"), App.Options(TimberlineDateFormat)) & ","
            s = s & Format(rs("InvoiceDiscountDate"), App.Options(TimberlineDateFormat)) & ","
            s = s & Format(rs("InvoicePaymentDate"), App.Options(TimberlineDateFormat)) & ","
            s = s & Format(rs("InvoiceAccountingDate"), App.Options(TimberlineDateFormat)) & ","
            s = s & Quote("" & rs("InvoiceCode1")) & ","
            s = s & Quote("" & rs("InvoiceCode2"))
            If "" & rs("VendorType") = "Summary" Then
                s = s & "," & Quote("" & rs("InvoiceVendorName"))
                s = s & "," & Quote("" & rs("InvoiceVendorAddr1"))
                s = s & "," & Quote("" & rs("InvoiceVendorAddr2"))
                s = s & "," & Quote("" & rs("InvoiceVendorCity"))
                s = s & "," & Quote("" & rs("InvoiceVendorProv"))
                s = s & "," & Quote("" & rs("InvoiceVendorPostal"))
            End If
            
            Print #invImp, s
            Print #invHst, s
            
            
            'cant do this here... on invoices with lots of distributions the previous select will block these updates. do them after.
            postedInvoiceIDs = postedInvoiceIDs & "," & lastID
            ''reset accounting date
            'Call HFApp.SqlExec("update invoices set AccountingDate=curdate() where DivisionID =" & HFApp.DivisionID & " and status='Hold' and InvoiceID=" & lastID)
            ''change status & set date
            'Call HFApp.SqlExec("update invoices set postingdate=" & DbQuote(Date, Now()) & ", status='Exported' where DivisionID =" & HFApp.DivisionID & " and InvoiceID=" & lastID)
            
        End If
        
        'write estimate item
        If "" & rs("ItemJob") <> "" Then
            s = ""
            If prevJob <> "" & rs("ItemJob") Then
                s = s & "*," & Quote("" & rs("ItemJob")) & vbCrLf
            End If
            If prevExtra <> "" & rs("ItemExtra") Then
                If "" & rs("ItemExtra") <> "" Then
                    s = s & "E," & Quote("" & rs("ItemExtra")) & vbCrLf
                End If
            End If
            s = s & "P," & Quote("" & rs("ItemPhase")) & ",,," & Format(Now(), "mmddyyyy") & vbCrLf
            s = s & "C," & Quote("" & rs("ItemPhase")) & ",," & Quote("" & rs("ItemCategory")) & "," & Format(Now(), "mmddyyyy") & vbCrLf
            Print #estImp, s
            Print #estHst, s
            prevJob = "" & rs("ItemJob")
            prevExtra = "" & rs("ItemExtra")
        End If
        
        'write invoiceitem
        If Val("" & rs("ItemID")) <> 0 And Val("" & rs("ItemAmount")) <> 0 Then
            
            s = ""
            s = s & "APDF"
            If UCase(Trim("" & rs("ItemCommitmentVendor"))) = UCase(Trim("" & rs("Vendor"))) Then
                s = s & "," & Quote("" & rs("ItemCommitment"))
                s = s & "," & rs("ItemCommitmentItem")
            Else
                s = s & "," 'commitment
                s = s & "," 'commitment line
            End If
            s = s & "," 'equipment
            s = s & "," 'eq cost code
            s = s & "," & Quote("" & rs("ItemJob"))
            s = s & "," & Quote("" & rs("ItemExtra"))
            s = s & "," & Quote("" & rs("ItemPhase"))
            s = s & "," & Quote("" & rs("ItemCategory"))
            s = s & "," 'bl std item
            s = s & "," 'reserved
            s = s & "," & Quote("" & rs("ItemDebitAccount"))
            s = s & "," 'creditaccount
            s = s & "," 'misc deduction percent
            s = s & "," & Quote("" & rs("ItemTaxGroup"))
            s = s & "," & Val("" & rs("ItemQuantity"))
            s = s & "," & Val("" & rs("ItemUnitPrice"))
            s = s & "," & Format(Val("" & rs("ItemAmount")), "###.00")
            s = s & "," & Format(Val("" & rs("ItemTax")), "###.00")
            s = s & "," 'tax liability
            s = s & "," 'discount offered
            s = s & "," & Format(Val("" & rs("ItemRetainage")), "###.00")
            s = s & "," & Format(rs("MiscDeductionAmt"), "###.00")
            s = s & "," '1099 exempt
            s = s & "," 'dist code
            s = s & "," 'draw
            s = s & "," 'misc1
            s = s & "," 'misc1 units
            s = s & "," 'misc2
            s = s & "," 'misc2 units
            s = s & "," 'eq meter
            s = s & "," & Quote("" & rs("ItemDescription"), , 30)
            s = s & "," 'authorization
            s = s & "," & Quote("" & rs("JointPayee"), , 30)
            Print #invImp, s
            Print #invHst, s
        End If
        
        rs.MoveNext
    Wend
    Close
    
    
    
    rs.Close
    'reset accounting date
    s = "update invoices set AccountingDate=curdate() where DivisionID =" & HFApp.DivisionID & " and status='Hold' and InvoiceID in(" & Mid(postedInvoiceIDs, 2) & ")"
    Call HFApp.SqlExec(s)
    'change status & set date
    s = "update invoices set postingdate=" & DbQuote(Date, Now()) & ", status='Exported' where DivisionID =" & HFApp.DivisionID & " and InvoiceID in(" & Mid(postedInvoiceIDs, 2) & ")"
    Call HFApp.SqlExec(s)

    
    Screen.MousePointer = vbDefault
    
    'run macro
    Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), HFApp.LoginID, HFApp.LoginPswd, App.Options(ImportMacroFile), App.Options(ImportPrintFile))
    
    Exit Sub
eh: Select Case err.Number
        Case 70
            MsgBox "Unable to post. The export file is in use by another user." & vbCrLf & vbCrLf & "Wait 10 seconds and try the operation again.", vbExclamation, App.ProductName
        Case Else
            Call ErrHandler(SRCFILE & "PostToTimberline")
    End Select
    Close
    Screen.MousePointer = vbDefault
End Sub
Private Sub PostToSage300(WhereClause As String)
'On Error GoTo eh
'
'    Dim X As New HyphenSage300Macaroni.APInvoices
'
'    Dim s        As String
'    Dim rs       As Recordset
'    Dim lastID   As Long
'    Dim MiscDeductionRate  As Double
'    Dim MiscDeductionAmt  As Double
'
'    Dim printFile As String
'    Dim rejectFile As String
'    Dim TLDateFormat As String
'
'    TLDateFormat = HFApp.Options.ValueByName("TimberlineDateFormat")
'    If TLDateFormat = "" Then TLDateFormat = "mm\/dd\/yyyy"
'
'    Screen.MousePointer = vbHourglass
'
'    'open invoices query
'    s = ""
'    s = s & "select Invoices.InvoiceID" & vbCrLf
'    s = s & "      ,Invoices.Vendor" & vbCrLf
'    s = s & "      ,tblvendors.tradetype VendorType" & vbCrLf
'    s = s & "      ,Invoices.Invoice" & vbCrLf
'    s = s & "      ,left(Invoices.Description,30) InvoiceDescription" & vbCrLf
'    s = s & "      ,Invoices.InvoiceCode1" & vbCrLf
'    s = s & "      ,Invoices.InvoiceCode2" & vbCrLf
'    s = s & "      ,Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
'    s = s & "      ,Invoices.Tax                   InvoiceTax" & vbCrLf
'    s = s & "      ,Invoices.Discount              InvoiceDiscount" & vbCrLf
'    s = s & "      ,Invoices.InvoiceDate           InvoiceDate" & vbCrLf
'    s = s & "      ,Invoices.ReceivedDate          InvoiceReceivedDate" & vbCrLf
'    s = s & "      ,Invoices.DiscountDate          InvoiceDiscountDate" & vbCrLf
'    s = s & "      ,Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
'    s = s & "      ,Invoices.AccountingDate        InvoiceAccountingDate" & vbCrLf
'    s = s & "      ,Invoices.VendorName            InvoiceVendorName" & vbCrLf
'    s = s & "      ,Invoices.VendorAddr1           InvoiceVendorAddr1" & vbCrLf
'    s = s & "      ,Invoices.VendorAddr2           InvoiceVendorAddr2" & vbCrLf
'    s = s & "      ,Invoices.VendorCity            InvoiceVendorCity" & vbCrLf
'    s = s & "      ,Invoices.VendorProv            InvoiceVendorProv" & vbCrLf
'    s = s & "      ,Invoices.VendorPostal          InvoiceVendorPostal" & vbCrLf
'    s = s & "      ,invoiceitems.ItemID            ItemID" & vbCrLf
'    s = s & "      ,invoiceitems.CommitmentVendor  ItemCommitmentVendor" & vbCrLf
'    s = s & "      ,invoiceitems.Commitment        ItemCommitment" & vbCrLf
'    s = s & "      ,invoiceitems.CommitmentItem    ItemCommitmentItem" & vbCrLf
'    s = s & "      ,invoiceitems.Job               ItemJob" & vbCrLf
'    s = s & "      ,invoiceitems.Extra             ItemExtra" & vbCrLf
'    s = s & "      ,invoiceitems.Phase             ItemPhase" & vbCrLf
'    s = s & "      ,invoiceitems.Category          ItemCategory" & vbCrLf
'    s = s & "      ,invoiceitems.DebitAccount      ItemDebitAccount" & vbCrLf
'    s = s & "      ,invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
'    s = s & "      ,invoiceitems.InvoicedQuantity  ItemQuantity" & vbCrLf
'    s = s & "      ,invoiceitems.InvoicedUnitPrice ItemUnitPrice" & vbCrLf
'    s = s & "      ,invoiceitems.PreTax + invoiceitems.Tax    ItemAmount" & vbCrLf
'    s = s & "      ,invoiceitems.Tax               ItemTax" & vbCrLf
'    s = s & "      ,round(invoiceitems.Retainage*(1+isnull(t.RetainageRate,0)/100),2)     ItemRetainage" & vbCrLf
'    s = s & "      ,left(invoiceitems.Description,30)       ItemDescription" & vbCrLf
'    s = s & "  from invoices" & vbCrLf
'    s = s & "  left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
'    s = s & "  left outer join tblvendors   on invoices.DivisionID=tblvendors.DivisionID and invoices.vendor=tblvendors.vendor_id" & vbCrLf
'    s = s & "  left outer join taxgroups  t on invoiceitems.divisionid=t.divisionid and invoiceitems.taxgroup=t.taxgroup" & vbCrLf
'    s = s & " where invoices.DivisionID = " & HFApp.DivisionID & " and 1=1" & vbCrLf
'    s = s & WhereClause & vbCrLf
'    s = s & "order by Invoices.InvoiceID,invoiceitems.ItemID"
'    Set rs = HFApp.SqlExec(s)
'    lastID = -999
'
'    While Not rs.EOF
'        If lastID <> rs("InvoiceID") Then
'
'            lastID = rs("InvoiceID")
'
'            MiscDeductionRate = Val("" & HFApp.SqlExec("SELECT MiscDeductionRate FROM tblVendors WHERE DivisionID =" & HFApp.DivisionID & " and vendor_ID=" & DbQuote(Str, "" & rs("Vendor")), dbHomefront)(0))
'            MiscDeductionAmt = Val("" & rs("InvoiceAmount")) * MiscDeductionRate / 100
'
'            X.AddInvoice _
'                  "" & rs("Vendor") _
'                , "" & rs("Invoice") _
'                , "" & rs("InvoiceDescription") _
'                , Round("" & rs("InvoiceDiscount"), 2) _
'                , Round(MiscDeductionAmt, 2) _
'                , rs("InvoiceDate") _
'                , Format("" & rs("InvoiceReceivedDate"), TLDateFormat) _
'                , Format("" & rs("InvoiceDiscountDate"), TLDateFormat) _
'                , Format("" & rs("InvoicePaymentDate"), TLDateFormat) _
'                , Format("" & rs("InvoiceAccountingDate"), TLDateFormat) _
'                , "" & rs("InvoiceCode1") _
'                , "" & rs("InvoiceCode2") _
'                , "" & rs("InvoiceVendorName") _
'                , "" & rs("InvoiceVendorAddr1") _
'                , "" & rs("InvoiceVendorAddr2") _
'                , "" & rs("InvoiceVendorCity") _
'                , "" & rs("InvoiceVendorProv") _
'                , "" & rs("InvoiceVendorPostal")
'
'
'            'change status & set date
'            Call HFApp.SqlExec("update invoices set postingdate=" & DbQuote(Date, Now()) & ", status='Exported' where DivisionID =" & HFApp.DivisionID & " and InvoiceID=" & lastID)
'
'        End If
'
'
'        'write invoiceitem
'        If Val("" & rs("ItemID")) <> 0 Then
'            X.AddInvoiceItem _
'                  "" & rs("Vendor") _
'                , "" & rs("Invoice") _
'                , IIf(UCase(Trim("" & rs("ItemCommitmentVendor"))) = UCase(Trim("" & rs("Vendor"))), "" & rs("ItemCommitment"), "") _
'                , IIf(UCase(Trim("" & rs("ItemCommitmentVendor"))) = UCase(Trim("" & rs("Vendor"))), "" & rs("ItemCommitmentItem"), "") _
'                , "" & rs("ItemJob") _
'                , "" & rs("ItemExtra") _
'                , "" & rs("ItemPhase") _
'                , "" & rs("ItemCategory") _
'                , "" & rs("ItemDebitAccount") _
'                , "" & rs("ItemTaxGroup") _
'                , Val("" & rs("ItemQuantity")) _
'                , Val("" & rs("ItemUnitPrice")) _
'                , Round("" & rs("ItemAmount"), 2) _
'                , Round("" & rs("ItemTax"), 2) _
'                , Round("" & rs("ItemRetainage"), 2) _
'                , left("" & rs("ItemDescription"), 30)
'        End If
'
'        rs.MoveNext
'    Wend
'
'
'    Dim troubleshoot As Boolean
'    troubleshoot = HFApp.Options.ValueByName("TroubleshootSageMacaroni") = "True"
'    Call X.Post(HFApp.Options(Timberline_Data_Path), HFApp.LoginID, HFApp.LoginPswd, True, troubleshoot, printFile, rejectFile)
'
'    'write print where the old posting used to be
'    On Error Resume Next
'    Dim i As Integer
'    Dim Filename As String
'    Filename = ForceExt(App.Options(ImportPrintFile), "txt")
'    i = FreeFile()
'    Call CreatePath("", FilePath(Filename))
'    Open Filename For Output As #i
'    Print #i, printFile
'    Close #i
'    On Error GoTo eh
'
'
'
'    Screen.MousePointer = vbDefault
'
'
'    Exit Sub
'eh: Call ErrHandler(SRCFILE & "PostToSage300")
'    Screen.MousePointer = vbDefault
End Sub

Private Sub PostToSimply(WhereClause As String)
On Error GoTo eh

    Dim es As String
    Dim s        As String
    Dim rs       As Recordset
    Dim lastID   As Long
    Dim Qty As Double, Pretax As Double
    Dim Rate As Double, Tax As Double, tax2 As Double
    Dim X As SimplyWrapper.APInvoice
    
    Screen.MousePointer = vbHourglass
    
es = "opening simply database"
    Set X = New SimplyWrapper.APInvoice
    If Not X.OpenDB(HFApp.Options(SimplyDataFile), HFApp.Options.ValueByName("SimplyUID"), HFApp.Options.ValueByName("SimplyPWD")) Then
        Screen.MousePointer = vbDefault
        MsgBox "Unable to connect to Simply Accounting" & vbCrLf & vbCrLf & "Check that the Simply Accounting security context and password in system settings is correct.", vbCritical, App.ProductName
        Exit Sub
    End If
    
    'open invoices query
es = "querying homefront db"
    s = ""
    s = s & "select Invoices.InvoiceID" & vbCrLf
    s = s & "      ,Invoices.Vendor" & vbCrLf
    s = s & "      ,Invoices.VendorType" & vbCrLf
    s = s & "      ,Invoices.Invoice" & vbCrLf
    s = s & "      ,Invoices.Description           InvoiceDescription" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode1" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode2" & vbCrLf
    s = s & "      ,Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & "      ,Invoices.Tax                   InvoiceTax" & vbCrLf
    s = s & "      ,Invoices.Discount              InvoiceDiscount" & vbCrLf
    s = s & "      ,Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & "      ,Invoices.ReceivedDate          InvoiceReceivedDate" & vbCrLf
    s = s & "      ,Invoices.DiscountDate          InvoiceDiscountDate" & vbCrLf
    s = s & "      ,Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & "      ,Invoices.AccountingDate        InvoiceAccountingDate" & vbCrLf
    s = s & "      ,Invoices.VendorName            InvoiceVendorName" & vbCrLf
    s = s & "      ,invoiceitems.ItemID            ItemID" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentVendor  ItemCommitmentVendor" & vbCrLf
    s = s & "      ,invoiceitems.Commitment        ItemCommitment" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentItem    ItemCommitmentItem" & vbCrLf
    s = s & "      ,invoiceitems.Job               ItemJob" & vbCrLf
    s = s & "      ,tbljobs.ExternalJobID          ItemProject" & vbCrLf
    s = s & "      ,invoiceitems.Extra             ItemExtra" & vbCrLf
    s = s & "      ,c.externalid                   ItemPhase" & vbCrLf
    s = s & "      ,sc.externalid                  ItemCategory" & vbCrLf
    s = s & "      ,a.account                      ItemDebitAccount" & vbCrLf
    s = s & "      ,invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedQuantity  ItemQuantity" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedUnitPrice ItemUnitPrice" & vbCrLf
    s = s & "      ,invoiceitems.PreTax            ItemPretax" & vbCrLf
    s = s & "      ,invoiceitems.Tax               ItemTax" & vbCrLf
    s = s & "      ,invoiceitems.TaxRate " & vbCrLf
    s = s & "      ,invoiceitems.Description       ItemDescription" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "       left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "       left outer join tbljobs on invoiceitems.DivisionID = tbljobs.DivisionID and invoiceitems.job=tbljobs.job_no" & vbCrLf
    s = s & "       LEFT OUTER JOIN standardcostcodes c ON(invoiceitems.DivisionID = c.DivisionID and invoiceitems.phase=c.costcode)" & vbCrLf
    s = s & "       LEFT OUTER JOIN standardcategories sc ON(invoiceitems.DivisionID = sc.DivisionID and invoiceitems.category=sc.category)" & vbCrLf
    s = s & "       left outer join glaccounts a on invoiceitems.DivisionID = a.DivisionID and invoiceitems.debitaccount=a.account" & vbCrLf
    s = s & " where invoices.DivisionID =" & HFApp.DivisionID & " and 1=1" & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "order by Invoices.InvoiceID,invoiceitems.ItemID"
    Set rs = HFApp.SqlExec(s)
    
    lastID = -999
    While Not rs.EOF
        
        If lastID <> rs("InvoiceID") Then
        
            If lastID <> -999 Then
                es = "saving to simply"
                If X.SaveInvoice Then
                    es = "changing hf invoice status"
                    Call SetInvoiceStatus(lastID, "Posted")
                End If
            End If
            
            lastID = rs("InvoiceID")
            es = "creating simply invoice"
            Call X.CreateInvoice("" & rs("Vendor"), "" & rs("Invoice"), CDate("" & rs("InvoiceDate")), "" & rs("InvoiceDescription"), False)
               
        End If
        
        'write invoiceitem
        If Val("" & rs("ItemID")) <> 0 Then
            Pretax = Val("" & rs("ItemPretax"))
            Qty = Round(Val("" & rs("ItemQuantity")), 4)
            If Qty = 0 Then Qty = 1
            Tax = Val("" & rs("ItemTax"))
            Rate = Pretax / Qty
            
            es = "adding line to simply invoice"
            Call X.AddLine("" & rs("ItemDescription"), Qty, "", Rate, FormatAccount("" & rs("ItemDebitAccount")), "" & rs("ItemTaxGroup"), "" & rs("ItemProject"), Tax)
            
        End If
        
        rs.MoveNext
    Wend
    If lastID <> -999 Then
        es = "saving to simply"
        If X.SaveInvoice Then
            es = "changing hf invoice status"
            Call SetInvoiceStatus(lastID, "Posted")
        End If
    End If
    
    es = "closing simply"
    Call X.CloseDB
    Screen.MousePointer = vbDefault
    
    
Exit Sub
eh:
'Call X.CloseDB
Screen.MousePointer = vbDefault
Select Case True
    
    'translate these errors ---------------------------------------------
    Case err.Description = "The invoice number has already been used for this Vendor."
        MsgBox "Unable to post" & vbCrLf & vbCrLf & err.Description, vbInformation, App.ProductName
    
    Case err.Description = "Other users are currently working with this company. You may open the company in multi-user mode or wait for the other users to stop using the company and try again."
        MsgBox "Unable to post" & vbCrLf & vbCrLf & "Your Simply Accounting database is locked by another user. They have to" & vbCrLf & "open the company file in multi-user mode, or you have wait until they are" & vbCrLf & "finished.", vbInformation, App.ProductName
    
    Case err.Description = "Vendor record not found. It may have been deactivated."
        MsgBox "Unable to post" & vbCrLf & vbCrLf & "The vendor record could not be found in your Simply Accounting database. It may have been deactivated.", vbInformation, App.ProductName
    
    'warnings -----------------------------------------------------------
    Case err.Description = "Note: The purchase will be processed and inventory updated, but there is no entry to reflect this transaction."
        MsgBox err.Description, vbInformation, App.ProductName
        Resume Next
    
    
    'ignorable crap -----------------------------------------------------
    Case err.Description = "Do you want to send an e-mail confirmation?"
        Resume Next
    Case left(err.Description, 43) = "This transaction was processed successfully"
        Resume Next
    Case err.Description = "The current allocation will be applied to the entire transaction.  Therefore, any allocation previously defined will be overwritten.  To undo this action, un-check the checkbox. (field=allocation)"
        Resume Next
    Case left(err.Description, 56) = "This transaction has been processed without Fast Posting"
        Resume Next

   
    'everything else ----------------------------------------------------
    Case Else
        Call ErrHandler(SRCFILE & "PostInvoices", es)
End Select
End Sub

Private Function MYOBInsert(SQLStmt As String) As Boolean
 
    Dim rstADO As ADODB.Recordset
    Dim MYOBInsertError As String
    Dim i As Integer
    
    'log request
    i = FreeFile()
    Open PathAppend(App.path & "\MYOBInsert.txt") For Output As #i
    Print #i, SQLStmt
    Close #i
    On Error GoTo Err_Handler
    

    If (Len(SQLStmt) = 0) Then
        MsgBox "No Insert Statement has been entered"
        Exit Function
    End If
    
    m_ADOConnection.BeginTrans
    
    Call m_ADOConnection.Execute(SQLStmt)
    Call m_ADOConnection.Execute("END TRANSACTION")
    
    m_ADOConnection.CommitTrans
     MYOBInsert = True
    
    
    Exit Function
        
Err_Handler:
    If err.Description = "Cannot start more transactions on this session." Then
        Resume Next
    End If
    ''old method
    If InStr(1, err.Description, "Warning", vbTextCompare) <> 0 Or InStr(1, err.Description, "default substituted", vbTextCompare) <> 0 Or InStr(1, err.Description, "ignored", vbTextCompare) <> 0 Then
        MYOBInsert = True
    
    Else
        MYOBInsertError = GetADO_Error()
    ''new method
        MsgBox err.Description & " " & MYOBInsertError
    'Resume
        m_ADOConnection.RollbackTrans
        MYOBInsert = False
    End If

End Function

Private Function GetADO_Error() As String

    Dim i As Integer
    Dim sQuery As String
    Dim nNativeError As Long
    Dim ADORecordset As ADODB.Recordset

    GetADO_Error = ""
    
    If (m_ADOConnection.Errors.Count = 0) Then Exit Function
    
    'Search for "MYOB ODBC" errors only, if there are no MYOB OBC errors then search for system errors
    For i = 0 To m_ADOConnection.Errors.Count - 1
        If (InStr(m_ADOConnection.Errors(i).Description, "[MYOB ODBC]") > 0) Then
            nNativeError = m_ADOConnection.Errors(i).NativeError
            Exit For
        End If
    Next i
        
    If ((nNativeError >= 10000) And (nNativeError < 20000)) Then 'Error
       GetADO_Error = "Error: " & Trim((nNativeError))
       nNativeError = nNativeError - 10000
       sQuery = "Select Description from ImportErrors where ImportErrorID = " & Trim((nNativeError))
    ElseIf ((nNativeError >= 1) And (nNativeError < 10000)) Then 'Warning
       sQuery = "Select Description from ImportWarnings where ImportWarningID = " & Trim((nNativeError))
       GetADO_Error = "Warning: " & Trim((nNativeError))
    ElseIf (nNativeError >= 20000) Then
       sQuery = "Select Description from InternalODBCErrors where NativeErrorNumber = '" & Trim((nNativeError) & "'")
       GetADO_Error = Trim((nNativeError))
    Else
        GetADO_Error = err.Description
        Exit Function
    End If
    
Set ADORecordset = m_ADOConnection.Execute(sQuery)
    GetADO_Error = GetADO_Error & " - " & ADORecordset.Fields("Description").Value
    ADORecordset.Close
    
    
End Function

Private Sub SetInvoiceStatus(InvoiceID As Long, Status As String)
    Dim s As String
    s = s & "update invoices" & vbCrLf
    s = s & "   set status=" & DbQuote(Str, Status) & vbCrLf
    If Status = "Posted" Then
        s = s & ",PostingDate=" & DbQuote(Date, Now()) & vbCrLf
    End If
    s = s & " where DivisionID =" & HFApp.DivisionID & " and invoiceid=" & DbQuote(num, InvoiceID)
    Call HFApp.SqlExec(s, dbHomeFront)
End Sub

Private Sub PostToQuickBooks(WhereClause As String)
    Dim IsUSVersion As Boolean
    IsUSVersion = HFApp.Options.ValueByName("AccountingVersion") <> "CA"
    If IsUSVersion Then
        Call PostToQuickBooksUS(WhereClause)
    Else
        Call PostToQuickBooksCA(WhereClause)
    End If
End Sub


Private Sub PostToQuickBooksOnlineUS(WhereClause As String)
On Error GoTo eh
    Dim s        As String

    
    Dim rs       As ADODB.Recordset
    Dim lastID   As Long
    Dim IsCredit As Boolean
    Dim docType  As String
    Dim Tax As Double
    Dim Response As String
    Dim BrokerUAT As Boolean
'    Dim POXml    As String
'    Dim PONumber As String
'    Dim POLine   As String
'    Dim POTrnID  As String
'    Dim POLineID As String
    
    'Set value to false when compiling for production
    BrokerUAT = "" & HFApp.Options.ValueByName("QBOBrokerUAT") = "True"
     
    Screen.MousePointer = vbHourglass
            
    
    'open invoices query
    s = ""
    s = s & "select " & vbCrLf
    s = s & " Invoices.InvoiceID" & vbCrLf
    s = s & ",Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & ",isnull(nullif(v.ExternalID,''),v.Vendor_ID) Vendor" & vbCrLf
    s = s & ",Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & ",Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & ",Invoices.Invoice" & vbCrLf
    s = s & ",Invoices.Description           InvoiceDescription" & vbCrLf
    s = s & ",invoiceitems.commitment        PONumber" & vbCrLf
    s = s & ",invoiceitems.commitmentitem    POLine" & vbCrLf
    s = s & ",invoiceitems.ItemID            ItemID" & vbCrLf
    s = s & ",isnull(nullif(tbljobs.ExternalJobID,''),tbljobs.Job_No)          ItemJob" & vbCrLf
    s = s & ",isnull(isnull(nullif(a.externalid,''),nullif(ca.externalid,'')),nullif(sca.externalid,'')) ItemDebitAccountID" & vbCrLf
    s = s & ",invoiceitems.PreTax            ItemAmount" & vbCrLf
    s = s & ",invoiceitems.Description       ItemDescription" & vbCrLf
    s = s & ",invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & ",c.externalid                   ItemPhase" & vbCrLf
    s = s & ",invoiceitems.InvoicedQuantity  ItemQuantity" & vbCrLf
    s = s & ",sc.externalid                  ItemCategory" & vbCrLf
    s = s & ",invoiceitems.Billable          ItemBillable" & vbCrLf
    s = s & ",v.PaymentTermsID" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "join tblvendors v on v.DivisionID = invoices.DivisionID and v.Vendor_ID = invoices.Vendor" & vbCrLf
    s = s & "left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "left outer join glaccounts a on invoiceitems.DivisionID = a.DivisionID and invoiceitems.debitaccount=a.account" & vbCrLf
    s = s & "left outer join tbljobs on invoiceitems.DivisionID = tbljobs.DivisionID and invoiceitems.job=tbljobs.job_no" & vbCrLf
    s = s & "left outer join standardcostcodes c ON invoiceitems.DivisionID = c.DivisionID and invoiceitems.phase=c.costcode" & vbCrLf
    s = s & "left outer join glaccounts ca on c.DivisionID = ca.DivisionID and c.debitaccount=ca.account" & vbCrLf
    s = s & "left outer join standardcategories sc ON invoiceitems.DivisionID = sc.DivisionID and invoiceitems.category=sc.category" & vbCrLf
    s = s & "left outer join glaccounts sca on sc.DivisionID = sca.DivisionID and sc.debitaccount=sca.account" & vbCrLf
    s = s & "where invoices.DivisionID =" & HFApp.DivisionID & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "" & vbCrLf
    s = s & "union all" & vbCrLf
    'US taxes get sent as separate lines
    s = s & "" & vbCrLf
    s = s & "select " & vbCrLf
    s = s & " Invoices.InvoiceID" & vbCrLf
    s = s & ",Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & ",isnull(nullif(v.ExternalID,''),v.Vendor_ID) Vendor" & vbCrLf
    s = s & ",Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & ",Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & ",Invoices.Invoice" & vbCrLf
    s = s & ",Invoices.Description           InvoiceDescription" & vbCrLf
    s = s & ",''                             PONumber" & vbCrLf
    s = s & ",''                             POLine" & vbCrLf
    s = s & ",99                             ItemID" & vbCrLf
    s = s & ",isnull(nullif(tbljobs.ExternalJobID,''),tbljobs.Job_No)          ItemJob" & vbCrLf
    s = s & ",isnull(isnull(nullif(a.externalid,''),nullif(ca.externalid,'')),nullif(sca.externalid,'')) ItemDebitAccountID" & vbCrLf
    s = s & ",sum(invoiceitems.Tax)          ItemAmount" & vbCrLf
    s = s & ",tg.Description                 ItemDescription" & vbCrLf
    s = s & ",invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & ",c.externalid                   ItemPhase" & vbCrLf
    s = s & ",1                              ItemQuantity" & vbCrLf
    s = s & ",sc.externalid                  ItemCategory" & vbCrLf
    s = s & ",invoiceitems.Billable          ItemBillable" & vbCrLf
    s = s & ",v.PaymentTermsID" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "join tblvendors v on v.DivisionID = invoices.DivisionID and v.Vendor_ID = invoices.Vendor" & vbCrLf
    s = s & "left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "left outer join glaccounts a on invoiceitems.DivisionID = a.DivisionID and invoiceitems.debitaccount=a.account" & vbCrLf
    s = s & "left outer join tbljobs on invoiceitems.DivisionID = tbljobs.DivisionID and invoiceitems.job=tbljobs.job_no" & vbCrLf
    s = s & "left outer join standardcostcodes c ON invoiceitems.DivisionID = c.DivisionID and invoiceitems.phase=c.costcode" & vbCrLf
    s = s & "left outer join glaccounts ca on c.DivisionID = ca.DivisionID and c.debitaccount=ca.account" & vbCrLf
    s = s & "left outer join standardcategories sc ON invoiceitems.DivisionID = sc.DivisionID and invoiceitems.category=sc.category" & vbCrLf
    s = s & "left outer join glaccounts sca on sc.DivisionID = sca.DivisionID and sc.debitaccount=sca.account" & vbCrLf
    s = s & "left outer join taxgroups tg on invoiceitems.taxgroup=tg.taxgroup and invoiceitems.divisionid=tg.divisionid" & vbCrLf
    s = s & "where invoices.DivisionID =" & HFApp.DivisionID & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "" & vbCrLf
    s = s & "group by" & vbCrLf
    s = s & " Invoices.InvoiceID" & vbCrLf
    s = s & ",Invoices.PreTax + Invoices.Tax" & vbCrLf
    s = s & ",isnull(nullif(v.ExternalID,''),v.Vendor_ID)" & vbCrLf
    s = s & ",Invoices.InvoiceDate" & vbCrLf
    s = s & ",Invoices.PaymentDate" & vbCrLf
    s = s & ",Invoices.Invoice" & vbCrLf
    s = s & ",Invoices.Description" & vbCrLf
    s = s & ",isnull(nullif(tbljobs.ExternalJobID,''),tbljobs.Job_No)" & vbCrLf
    s = s & ",isnull(isnull(nullif(a.externalid,''),nullif(ca.externalid,'')),nullif(sca.externalid,''))" & vbCrLf
    s = s & ",tg.Description" & vbCrLf
    s = s & ",invoiceitems.TaxGroup" & vbCrLf
    s = s & ",c.externalid" & vbCrLf
    s = s & ",sc.externalid,invoiceitems.Billable" & vbCrLf
    s = s & ",v.PaymentTermsID" & vbCrLf
    s = s & "having sum(invoiceitems.Tax)<>0" & vbCrLf
    s = s & "order by Invoices.InvoiceID,isnull(nullif(tbljobs.ExternalJobID,''),tbljobs.Job_No)"
    Set rs = HFApp.SqlExec(s)
    
    
    s = ""
    lastID = -999
    While Not rs.EOF
        If lastID <> rs("InvoiceID") Then
        
            If lastID <> -999 Then
                s = s & "</Bill>" & vbCrLf
                Response = SubmitBrokerXml("QBOnline", "PostBill", s)
                If Response <> "" And Not Response Like "*<Error>*" Then
                    Call SetInvoiceStatus(lastID, "Posted")
                End If
            End If
            
            lastID = rs("InvoiceID")
            IsCredit = Val("" & rs("InvoiceAmount")) < 0
            docType = IIf(IsCredit, "VendorCredit", "Bill")
            
            s = ""
            s = s & "<?xml version=""" & "1.0""" & " encoding=""" & "utf-8""" & "?>" & vbCrLf
            s = "<Bill xmlns:xsi=""" & "http://www.w3.org/2001/XMLSchema-instance""" & " xmlns:xsd=""" & "http://www.w3.org/2001/XMLSchema""" & ">" & vbCrLf
            s = s & HFApp.XmlQBAdd(d, 0, "TransactionDate", rs("InvoiceDate"))
            s = s & HFApp.XmlQBAdd(m, 0, "Vendor", rs("Vendor"))
            s = s & HFApp.XmlQBAdd(m, 0, "Invoice", rs("Invoice"))
            s = s & HFApp.XmlQBAdd(m, 0, "TransactionType", docType)
            If Not IsCredit Then
                s = s & "<SalesTermRef>"
                s = s & HFApp.XmlQBAdd(st, 0, "value", "" & rs("PaymentTermsID"))
                s = s & "</SalesTermRef>"
                If "" & rs("InvoicePaymentDate") <> "" Then
                    s = s & HFApp.XmlQBAdd(d, 0, "DueDate", "" & rs("InvoicePaymentDate"))
                End If
            End If
        End If
            
'        'DONT DO THIS.  THERE IS NO VALUE IN POSTING POS TO QB SINCE PAYMENTS ARENT ALLOCATED TO PO
'        'get po transaction ids from QB
'        If PONumber <> "" & rs("PONumber") Then
'            PONumber = "" & rs("PONumber")
'            POXml = QBOGetPODetails(PONumber)
'            POTrnID = Parse(Parse(POXml, 1, "</Id>"), 2, "<Id>")
'        End If
'        POLine = Val("" & rs("POLine"))
'        POLineID = Parse(Parse(Parse(POXml, POLine + 1, "<Line>"), 1, "</Id>"), 2, "<Id>")
            
        s = s & "<Line>" & vbCrLf
        s = s & vbTab & HFApp.XmlQBAdd(m, 0, "Description", fixxml("" & rs("ItemDescription")))
        s = s & vbTab & HFApp.XmlQBAdd(m, 0, "GLAccount", "" & rs("ItemDebitAccountID"))
        s = s & vbTab & HFApp.XmlQBAdd(m, 0, "JCCostCode", "" & rs("ItemPhase"))
        s = s & vbTab & HFApp.XmlQBAdd(m, 0, "Description", "" & fixxml("" & rs("ItemDescription")))
        s = s & vbTab & HFApp.XmlQBAdd(n, 999.5, "Qty", IIf(Val("" & rs("ItemQuantity")) = 0, 1, Val("" & rs("ItemQuantity"))))
        s = s & vbTab & HFApp.XmlQBAdd(m, 0, "Amount", Format(IIf(IsCredit, -1, 1) * (Val("" & rs("ItemAmount"))), "###.00"))
'        If PONumber <> "" Then
'            s = s & "<LinkToTxn>" & vbCrLf
'            s = s & HFApp.XmlQBAdd(m, 0, "TxnID", POTrnID)
'            s = s & HFApp.XmlQBAdd(m, 0, "TxnLineID", POLineID)
'            s = s & "</LinkToTxn>" & vbCrLf
'        End If
        s = s & vbTab & HFApp.XmlQBAdd(m, 0, "Job", "" & rs("ItemJob"))
        s = s & vbTab & HFApp.XmlQBAdd(m, 0, "JCCategory", "" & rs("ItemCategory"))
        s = s & vbTab & HFApp.XmlQBAdd(m, 0, "BillableStatus", IIf("" & rs("ItemBillable") = "True", "Billable", "NotBillable"))
        s = s & "</Line>" & vbCrLf


        rs.MoveNext
    Wend
    If lastID <> -999 Then
        s = s & "</Bill>" & vbCrLf
        Response = SubmitBrokerXml("QBOnline", "PostBill", s)
        If Response <> "" And Not Response Like "*<Error>*" Then
            Call SetInvoiceStatus(lastID, "Posted")
        End If
    End If
    Screen.MousePointer = vbDefault
    
Exit Sub
eh:
Call ErrHandler(SRCFILE & "PostToQuickBooksOnlineUS", s)
End Sub


Private Sub PostToQuickBooksUS(WhereClause As String)
On Error GoTo eh
    Dim s        As String
    Dim rs       As Recordset
    Dim lastID   As Long
    Dim IsCredit As Boolean
    Dim docType  As String
    Dim Tax As Double
    
    
    Screen.MousePointer = vbHourglass
            
    
    'open invoices query
    s = ""
    s = s & "select " & vbCrLf
    s = s & " Invoices.InvoiceID" & vbCrLf
    s = s & ",Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & ",Invoices.Vendor" & vbCrLf
    s = s & ",Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & ",Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & ",Invoices.Invoice" & vbCrLf
    s = s & ",Invoices.Description           InvoiceDescription" & vbCrLf
    s = s & ",invoiceitems.ItemID            ItemID" & vbCrLf
    s = s & ",tbljobs.ExternalJobID          ItemJob" & vbCrLf
    s = s & ",isnull(isnull(nullif(a.externalid,''),nullif(ca.externalid,'')),nullif(sca.externalid,'')) ItemDebitAccountID" & vbCrLf
    s = s & ",invoiceitems.PreTax            ItemAmount" & vbCrLf
    s = s & ",invoiceitems.Description       ItemDescription" & vbCrLf
    s = s & ",invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & ",c.externalid                   ItemPhase" & vbCrLf
    s = s & ",invoiceitems.invoicedquantity  ItemQuantity" & vbCrLf
    s = s & ",sc.externalid                  ItemCategory" & vbCrLf
    s = s & ",invoiceitems.Billable          ItemBillable" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "left outer join glaccounts a on invoiceitems.DivisionID = a.DivisionID and invoiceitems.debitaccount=a.account" & vbCrLf
    s = s & "left outer join tbljobs on invoiceitems.DivisionID = tbljobs.DivisionID and invoiceitems.job=tbljobs.job_no" & vbCrLf
    s = s & "left outer join standardcostcodes c ON invoiceitems.DivisionID = c.DivisionID and invoiceitems.phase=c.costcode" & vbCrLf
    s = s & "left outer join glaccounts ca on c.DivisionID = ca.DivisionID and c.debitaccount=ca.account" & vbCrLf
    s = s & "left outer join standardcategories sc ON invoiceitems.DivisionID = sc.DivisionID and invoiceitems.category=sc.category" & vbCrLf
    s = s & "left outer join glaccounts sca on sc.DivisionID = sca.DivisionID and sc.debitaccount=sca.account" & vbCrLf
    s = s & "where invoices.DivisionID =" & HFApp.DivisionID & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "" & vbCrLf
    s = s & "union all" & vbCrLf
    'US taxes get sent as separate lines
    s = s & "" & vbCrLf
    s = s & "select " & vbCrLf
    s = s & " Invoices.InvoiceID" & vbCrLf
    s = s & ",Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & ",Invoices.Vendor" & vbCrLf
    s = s & ",Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & ",Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & ",Invoices.Invoice" & vbCrLf
    s = s & ",Invoices.Description           InvoiceDescription" & vbCrLf
    s = s & ",99                             ItemID" & vbCrLf
    s = s & ",tbljobs.ExternalJobID          ItemJob" & vbCrLf
    s = s & ",isnull(isnull(nullif(a.externalid,''),nullif(ca.externalid,'')),nullif(sca.externalid,'')) ItemDebitAccountID" & vbCrLf
    s = s & ",sum(invoiceitems.Tax)          ItemAmount" & vbCrLf
    s = s & ",tg.Description                 ItemDescription" & vbCrLf
    s = s & ",invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & ",c.externalid                   ItemPhase" & vbCrLf
    s = s & ",1                              ItemQuantity" & vbCrLf
    s = s & ",sc.externalid                  ItemCategory" & vbCrLf
    s = s & ",invoiceitems.Billable          ItemBillable" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "left outer join glaccounts a on invoiceitems.DivisionID = a.DivisionID and invoiceitems.debitaccount=a.account" & vbCrLf
    s = s & "left outer join tbljobs on invoiceitems.DivisionID = tbljobs.DivisionID and invoiceitems.job=tbljobs.job_no" & vbCrLf
    s = s & "left outer join standardcostcodes c ON invoiceitems.DivisionID = c.DivisionID and invoiceitems.phase=c.costcode" & vbCrLf
    s = s & "left outer join glaccounts ca on c.DivisionID = ca.DivisionID and c.debitaccount=ca.account" & vbCrLf
    s = s & "left outer join standardcategories sc ON invoiceitems.DivisionID = sc.DivisionID and invoiceitems.category=sc.category" & vbCrLf
    s = s & "left outer join glaccounts sca on sc.DivisionID = sca.DivisionID and sc.debitaccount=sca.account" & vbCrLf
    s = s & "left outer join taxgroups tg on invoiceitems.taxgroup=tg.taxgroup and invoiceitems.divisionid=tg.divisionid" & vbCrLf
    s = s & "where invoices.DivisionID =" & HFApp.DivisionID & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "" & vbCrLf
    s = s & "group by" & vbCrLf
    s = s & " Invoices.InvoiceID" & vbCrLf
    s = s & ",Invoices.PreTax + Invoices.Tax" & vbCrLf
    s = s & ",Invoices.Vendor" & vbCrLf
    s = s & ",Invoices.InvoiceDate" & vbCrLf
    s = s & ",Invoices.PaymentDate" & vbCrLf
    s = s & ",Invoices.Invoice" & vbCrLf
    s = s & ",Invoices.Description" & vbCrLf
    s = s & ",tbljobs.ExternalJobID" & vbCrLf
    s = s & ",isnull(isnull(nullif(a.externalid,''),nullif(ca.externalid,'')),nullif(sca.externalid,''))" & vbCrLf
    s = s & ",tg.Description" & vbCrLf
    s = s & ",invoiceitems.TaxGroup" & vbCrLf
    s = s & ",c.externalid" & vbCrLf
    s = s & ",sc.externalid,invoiceitems.Billable" & vbCrLf
    s = s & "having sum(invoiceitems.Tax)<>0" & vbCrLf
    s = s & "order by Invoices.InvoiceID,tbljobs.ExternalJobID"
    Set rs = HFApp.SqlExec(s)
    
    
    s = ""
    lastID = -999
    While Not rs.EOF
        If lastID <> rs("InvoiceID") Then
        
            If lastID <> -999 Then
                s = s & "</" & docType & "Add>" & vbCrLf
                s = s & "</" & docType & "AddRq>" & vbCrLf
                s = s & HFApp.XmlQBEnd()
                Call QBSubmit(s, lastID)
                Call SetInvoiceStatus(lastID, "Posted")
            End If
            
            lastID = rs("InvoiceID")
            IsCredit = Val("" & rs("InvoiceAmount")) < 0
            docType = IIf(IsCredit, "VendorCredit", "Bill")
            
            s = HFApp.XmlQBStart()
            s = s & "<" & docType & "AddRq>" & vbCrLf
            s = s & "<" & docType & "Add>" & vbCrLf
            s = s & HFApp.XmlQBAdd(m, 0, "VendorRef", "" & rs("Vendor"))
            s = s & HFApp.XmlQBAdd(d, 0, "TxnDate", "" & rs("InvoiceDate"))
            If Not IsCredit Then
                s = s & HFApp.XmlQBAdd(d, 0, "DueDate", "" & rs("InvoicePaymentDate"))
            End If
            s = s & HFApp.XmlQBAdd(m, 0, "RefNumber", "" & rs("Invoice"))
            s = s & HFApp.XmlQBAdd(m, 0, "Memo", "" & fixxml("" & rs("InvoiceDescription")))
            
            
        End If
        
        If Val("" & rs("ItemID")) <> 0 Then
            If "" & rs("ItemJob") = "" Then
                'expense
                s = s & "<ExpenseLineAdd>" & vbCrLf
                s = s & HFApp.XmlQBAdd(m, 0, "AccountRef", "" & rs("ItemDebitAccountID"))
                s = s & HFApp.XmlQBAdd(m, 0, "Amount", Format(IIf(IsCredit, -1, 1) * (Val("" & rs("ItemAmount"))), "###.00"))
                s = s & HFApp.XmlQBAdd(m, 0, "Memo", "" & fixxml("" & rs("ItemDescription")))
                s = s & "</ExpenseLineAdd>" & vbCrLf
            Else
                'item
                s = s & "<ItemLineAdd>" & vbCrLf
                s = s & HFApp.XmlQBAdd(m, 0, "ItemRef", "" & rs("ItemPhase"))
                s = s & HFApp.XmlQBAdd(m, 0, "Desc", "" & fixxml("" & rs("ItemDescription")))
                s = s & HFApp.XmlQBAdd(n, 999.5, "Quantity", IIf(Val("" & rs("ItemQuantity")) = 0, 1, Val("" & rs("ItemQuantity"))))
                s = s & HFApp.XmlQBAdd(m, 0, "Amount", Format(IIf(IsCredit, -1, 1) * (Val("" & rs("ItemAmount"))), "###.00"))
                s = s & HFApp.XmlQBAdd(m, 0, "CustomerRef", "" & rs("ItemJob"))
                s = s & HFApp.XmlQBAdd(m, 0, "ClassRef", "" & rs("ItemCategory"))
                s = s & HFApp.XmlQBAdd(m, 0, "BillableStatus", IIf("" & rs("ItemBillable") = "True", "Billable", "NotBillable"))
                If "" & rs("ItemDebitAccountID") <> "" Then
                    s = s & HFApp.XmlQBAdd(m, 0, "OverrideItemAccountRef", "" & rs("ItemDebitAccountID"))
                End If
                s = s & "</ItemLineAdd>" & vbCrLf
            End If
        End If
        
        rs.MoveNext
    Wend
    If lastID <> -999 Then
        s = s & "</" & docType & "Add>" & vbCrLf
        s = s & "</" & docType & "AddRq>" & vbCrLf
        s = s & HFApp.XmlQBEnd()
        Call QBSubmit(s, lastID)
        Call SetInvoiceStatus(lastID, "Posted")
    End If
    Screen.MousePointer = vbDefault
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "PostToQuickBooksUS", s)
End Sub


Private Function QuickbooksGetPODetails(PONumber As String) As String
    Dim s As String
    Dim xml As String
    
    s = HFApp.XmlQBStart()
    s = s & "<PurchaseOrderQueryRq>" & vbCrLf
    s = s & HFApp.XmlQBAdd(m, 0, "RefNumber", PONumber)
    s = s & "<IncludeLineItems>true</IncludeLineItems>" & vbCrLf
    s = s & "</PurchaseOrderQueryRq>" & vbCrLf
    s = s & HFApp.XmlQBEnd()
    
    
    xml = HFApp.XmlQBSubmit(s)
    If xml = "" Then err.Raise 999, "QuickbooksGetPODetails", "PO is not found in Quickbooks."
    
    QuickbooksGetPODetails = xml

End Function



Private Sub PostToQuickBooksCA(WhereClause As String)
On Error GoTo eh
    Dim s        As String
    Dim rs       As Recordset
    Dim lastID   As Long
    Dim IsCredit As Boolean
    Dim docType  As String
    Dim Tax As Double
    
    Screen.MousePointer = vbHourglass
       
    
    'open invoices query
    s = ""
    s = s & "select Invoices.InvoiceID" & vbCrLf
    s = s & "      ,Invoices.Vendor" & vbCrLf
    s = s & "      ,Invoices.VendorType" & vbCrLf
    s = s & "      ,Invoices.Invoice" & vbCrLf
    s = s & "      ,Invoices.Description           InvoiceDescription" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode1" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode2" & vbCrLf
    s = s & "      ,Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & "      ,Invoices.Tax                   InvoiceTax" & vbCrLf
    s = s & "      ,Invoices.Discount              InvoiceDiscount" & vbCrLf
    s = s & "      ,Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & "      ,Invoices.ReceivedDate          InvoiceReceivedDate" & vbCrLf
    s = s & "      ,Invoices.DiscountDate          InvoiceDiscountDate" & vbCrLf
    s = s & "      ,Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & "      ,Invoices.AccountingDate        InvoiceAccountingDate" & vbCrLf
    s = s & "      ,Invoices.VendorName            InvoiceVendorName" & vbCrLf
    s = s & "      ,invoiceitems.ItemID            ItemID" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentVendor  ItemCommitmentVendor" & vbCrLf
    s = s & "      ,invoiceitems.Commitment        ItemCommitment" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentItem    ItemCommitmentItem" & vbCrLf
    s = s & "      ,tbljobs.ExternalJobID          ItemJob" & vbCrLf
    s = s & "      ,invoiceitems.Extra             ItemExtra" & vbCrLf
    s = s & "      ,c.externalid                   ItemPhase" & vbCrLf
    s = s & "      ,sc.externalid                  ItemCategory" & vbCrLf
    s = s & "      ,isnull(isnull(nullif(a.account,''),nullif(ca.account,'')),nullif(sca.account,'')) ItemDebitAccount" & vbCrLf
    s = s & "      ,isnull(isnull(nullif(a.externalid,''),nullif(ca.externalid,'')),nullif(sca.externalid,'')) ItemDebitAccountID" & vbCrLf
    s = s & "      ,invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedQuantity  ItemQuantity" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedUnitPrice ItemUnitPrice" & vbCrLf
    s = s & "      ,invoiceitems.PreTax            ItemPretax" & vbCrLf
    s = s & "      ,invoiceitems.Tax               ItemTax" & vbCrLf
    s = s & "      ,invoiceitems.TaxRate           TaxRate" & vbCrLf
    s = s & "      ,invoiceitems.Description       ItemDescription" & vbCrLf
    s = s & "      ,invoiceitems.Billable          ItemBillable" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "       left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "       left outer join glaccounts a on invoiceitems.DivisionID = a.DivisionID and invoiceitems.debitaccount=a.account" & vbCrLf
    s = s & "       left outer join tbljobs on invoiceitems.DivisionID = tbljobs.DivisionID and invoiceitems.job=tbljobs.job_no" & vbCrLf
    s = s & "       LEFT OUTER JOIN standardcostcodes c ON invoiceitems.DivisionID = c.DivisionID and invoiceitems.phase=c.costcode" & vbCrLf
    s = s & "       left outer join glaccounts ca on c.DivisionID = ca.DivisionID and c.debitaccount=ca.account" & vbCrLf
    s = s & "       LEFT OUTER JOIN standardcategories sc ON invoiceitems.DivisionID = sc.DivisionID and invoiceitems.category=sc.category" & vbCrLf
    s = s & "       left outer join glaccounts sca on sc.DivisionID = sca.DivisionID and sc.debitaccount=sca.account" & vbCrLf
    s = s & " where invoices.DivisionID =" & HFApp.DivisionID & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "order by Invoices.InvoiceID,tbljobs.ExternalJobID"
    Set rs = HFApp.SqlExec(s)
    
    
    s = ""
    lastID = -999
    While Not rs.EOF
        If lastID <> rs("InvoiceID") Then
        
            If lastID <> -999 Then
                s = s & "</" & docType & "Add>" & vbCrLf
                s = s & "</" & docType & "AddRq>" & vbCrLf
                s = s & HFApp.XmlQBEnd()
                Call QBSubmit(s, lastID)
                Call SetInvoiceStatus(lastID, "Posted")
            End If
            
            lastID = rs("InvoiceID")
            IsCredit = Val("" & rs("InvoiceAmount")) < 0
            docType = IIf(IsCredit, "VendorCredit", "Bill")
            
            s = HFApp.XmlQBStart()
            s = s & "<" & docType & "AddRq>" & vbCrLf
            s = s & "<" & docType & "Add>" & vbCrLf
            s = s & HFApp.XmlQBAdd(m, 0, "VendorRef", "" & rs("Vendor"))
            s = s & HFApp.XmlQBAdd(d, 0, "TxnDate", "" & rs("InvoiceDate"))
            
            If Not IsCredit Then
                s = s & HFApp.XmlQBAdd(d, 0, "DueDate", "" & rs("InvoicePaymentDate"))
            End If
            s = s & HFApp.XmlQBAdd(m, 0, "RefNumber", "" & rs("Invoice"))
            s = s & HFApp.XmlQBAdd(m, 0, "Memo", "" & fixxml("" & rs("InvoiceDescription")))
        End If
        
        If Val("" & rs("ItemID")) <> 0 Then
            If "" & rs("ItemJob") = "" Then
                'expense
                s = s & "<ExpenseLineAdd>" & vbCrLf
                s = s & HFApp.XmlQBAdd(m, 0, "AccountRef", "" & rs("ItemDebitAccountID"))
                s = s & HFApp.XmlQBAdd(m, 0, "Amount", Format(IIf(IsCredit, -1, 1) * Val("" & rs("ItemPretax")), "###.00"))
                s = s & HFApp.XmlQBAdd(m, 0, "Memo", "" & fixxml("" & rs("ItemDescription")))
                If "" & rs("ItemTaxGroup") <> "" Then
                    s = s & "<SalesTaxCodeRef>" & HFApp.XmlQBAdd(c, 3, "FullName", "" & rs("ItemTaxGroup")) & "</SalesTaxCodeRef>" & vbCrLf
                End If
                s = s & "</ExpenseLineAdd>" & vbCrLf
            Else
                'item
                s = s & "<ItemLineAdd>" & vbCrLf
                s = s & HFApp.XmlQBAdd(m, 0, "ItemRef", "" & rs("ItemPhase"))
                s = s & HFApp.XmlQBAdd(m, 0, "Desc", "" & fixxml("" & rs("ItemDescription")))
                s = s & HFApp.XmlQBAdd(n, 999.5, "Quantity", IIf(Val("" & rs("ItemQuantity")) = 0, 1, Val("" & rs("ItemQuantity"))))
                s = s & HFApp.XmlQBAdd(m, 0, "Amount", Format(IIf(IsCredit, -1, 1) * Val("" & rs("ItemPretax")), "###.00"))
                s = s & HFApp.XmlQBAdd(m, 0, "CustomerRef", "" & rs("ItemJob"))
                s = s & HFApp.XmlQBAdd(m, 0, "ClassRef", "" & rs("ItemCategory"))
                If "" & rs("ItemTaxGroup") <> "" Then
                    s = s & "<SalesTaxCodeRef>" & HFApp.XmlQBAdd(c, 3, "FullName", "" & rs("ItemTaxGroup")) & "</SalesTaxCodeRef>" & vbCrLf
                End If
                s = s & HFApp.XmlQBAdd(m, 0, "BillableStatus", IIf("" & rs("ItemBillable") = "True", "Billable", "NotBillable"))
                If "" & rs("ItemDebitAccountID") <> "" Then
                    s = s & HFApp.XmlQBAdd(m, 0, "OverrideItemAccountRef", "" & rs("ItemDebitAccountID"))
                End If
                s = s & "</ItemLineAdd>" & vbCrLf
            End If
        End If
        
        rs.MoveNext
    Wend
    If lastID <> -999 Then
        s = s & "</" & docType & "Add>" & vbCrLf
        s = s & "</" & docType & "AddRq>" & vbCrLf
        s = s & HFApp.XmlQBEnd()
        Call QBSubmit(s, lastID)
        Call SetInvoiceStatus(lastID, "Posted")
    End If
    Screen.MousePointer = vbDefault
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "PostToQuickBooksCA", s)
End Sub
Private Function fixxml(s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    fixxml = s
End Function
Private Sub QBSubmit(xml As String, InvoiceID As Long)
    Dim s As String
    s = HFApp.XmlQBSubmit(xml)
End Sub



Private Function FormatAccount(Account As String) As String
    Dim s As String
    s = Account
    While Len(s) < 8
        s = s & "0"
    Wend
    FormatAccount = s
End Function



Private Sub DoAccountingSystemConfig()
    Dim i As Long
    Dim bShowDiscount As Boolean
    Dim bShowAcctDate As Boolean
    Dim bShowPayDate As Boolean
    Dim bShowRetainage As Boolean
    
    'set custom descriptions
    With gEnterDist
        .TextMatrix(0, .ColIndex("Phase")) = App.Options(Caption_Phase)
        .TextMatrix(0, .ColIndex("Category")) = App.Options(Caption_Category)
        .TextMatrix(0, .ColIndex("PhaseDesc")) = App.Options(Caption_Phase) & " Desc"
        .TextMatrix(0, .ColIndex("CategoryDesc")) = App.Options(Caption_Category) & " Desc"
    End With
    With gViewDist
        .TextMatrix(0, .ColIndex("Phase")) = App.Options(Caption_Phase)
        .TextMatrix(0, .ColIndex("Category")) = App.Options(Caption_Category)
        .TextMatrix(0, .ColIndex("PhaseDesc")) = App.Options(Caption_Phase) & " Desc"
        .TextMatrix(0, .ColIndex("CategoryDesc")) = App.Options(Caption_Category) & " Desc"
    End With
    
    'hide elements by accounting system
    bShowDiscount = True
    bShowAcctDate = True
    bShowPayDate = True
    bShowRetainage = True
    Select Case HFApp.Options(AccountingSystem)
    
        Case asSimply
            bShowDiscount = False
            bShowAcctDate = False
            bShowPayDate = True
            bShowRetainage = False
            
        Case asXero
            bShowDiscount = False
            bShowAcctDate = False
            bShowPayDate = True
            bShowRetainage = False
            Label1(9).Caption = "Due Date"
        
        Case asQuickBooks, asQuickBooksOnline
            bShowDiscount = False
            bShowAcctDate = False
            bShowPayDate = True
            bShowRetainage = False
            gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Phase")) = ""
            gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Category")) = ""
            gViewDist.TextMatrix(0, gViewDist.ColIndex("Phase")) = ""
            gViewDist.TextMatrix(0, gViewDist.ColIndex("Category")) = ""
            
            
    End Select
    
    'discount
    dteDiscount.Visible = bShowDiscount
    numDiscount.Visible = bShowDiscount
    Label1(6).Visible = bShowDiscount
    Label1(5).Visible = bShowDiscount
    If Not bShowDiscount Then
        gViewInv.TextMatrix(0, gViewInv.ColIndex("Discount")) = ""
        gViewInv.TextMatrix(0, gViewInv.ColIndex("DiscountDate")) = ""
    End If
    
    'accounting date
    dteAccounting.Visible = bShowAcctDate
    Label1(10).Visible = bShowAcctDate
    gViewInv.ColHidden(gViewInv.ColIndex("AccountingDate")) = True
    If Not bShowAcctDate Then App.Options.Value(AccountingDateRequired) = False
    
    'payment date
    dtePayment.Visible = bShowPayDate
    Label1(9).Visible = bShowPayDate
    gViewInv.ColHidden(gViewInv.ColIndex("PaymentDate")) = True
    
    
    'retainage
    If Not bShowRetainage Then
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("Retainage")) = ""
        gEnterDist.TextMatrix(0, gEnterDist.ColIndex("RetainageRate")) = ""
        gViewDist.TextMatrix(0, gViewDist.ColIndex("Retainage")) = ""
    End If
        

        
        
    With gEnterDist
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "" Then .ColHidden(i) = True
        Next
    End With
    With gViewInv
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "" Then .ColHidden(i) = True
        Next
    End With
    With gViewDist
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "" Then .ColHidden(i) = True
        Next
    End With
    
    
    
    


End Sub



Private Function CodeOnInvoice(Job As String, Extra As String, Phase As String, Optional Category As String) As Boolean
    Dim i As Long
    With Me.gEnterDist
        For i = 1 To .Rows - 2
            If i <> .Row Then
                If .TextMatrix(i, .ColIndex("Job")) = Job And _
                   .TextMatrix(i, .ColIndex("Extra")) = Extra And _
                   .TextMatrix(i, .ColIndex("Phase")) = Phase And _
                   (.TextMatrix(i, .ColIndex("Category")) = Category Or Category = "") Then
                    
                    CodeOnInvoice = True
                    Exit Function
                End If
            End If
        Next
    End With
End Function

Private Sub ParseAddress(Address As String, Lin1 As String, Lin2 As String, City As String, Prov As String, Pstl As String)
Dim s As String
    
    'replace end of lines with chr(3)
    Address = Replace(Address, vbCrLf, Chr(3))
    Address = Replace(Address, vbLf, Chr(3))
    Address = Replace(Address, vbCr, Chr(3))
    
    'remove leading and trailing whitespace
    Address = Trim(Address)
    While left(Address, 1) = Chr(3)
        Address = Trim(Mid(Address, 2))
    Wend
    While Right(Address, 1) = Chr(3)
        Address = Trim(Mid(Address, 1, Len(Address) - 1))
    Wend
    
        
    Select Case Parse(Address, , Chr(3))
        Case 1
            '----------------------------
            '-- suite 123, 75th street
            '----------------------------
            Lin1 = Parse(Address, 1, Chr(3))
        
        Case 2
            '----------------------------
            '-- suite 123, 75th street
            '-- red deer, ab
            '----------------------------
            Lin1 = Trim(Parse(Address, 1, Chr(3)))
            City = Trim(Parse(Parse(Address, 2, Chr(3)), 1, ","))
            Prov = Trim(Parse(Parse(Address, 2, Chr(3)), 2, ","))
            
        Case 3
            '----------------------------
            '-- 148, 75th street
            '-- red deer, ab
            '-- T2P X1X
            '---OR-------------------------
            '-- suite 123
            '-- 75th street
            '-- red deer, ab
            '----------------------------
            
            If InStr(1, Parse(Address, 2, Chr(3)), ",", vbTextCompare) Then
                'if line 2 has a comma then it must be city, prov
                Lin1 = Trim(Parse(Address, 1, Chr(3)))
                City = Trim(Parse(Parse(Address, 2, Chr(3)), 1, ","))
                Prov = Trim(Parse(Parse(Address, 2, Chr(3)), 2, ","))
                Pstl = Trim(Parse(Address, 3, Chr(3)))
            Else
                'if not then city, prov will be on line 3
                Lin1 = Trim(Parse(Address, 1, Chr(3)))
                Lin2 = Trim(Parse(Address, 2, Chr(3)))
                City = Trim(Parse(Parse(Address, 3, Chr(3)), 1, ","))
                Prov = Trim(Parse(Parse(Address, 3, Chr(3)), 2, ","))
            End If
        
        
        Case Else
            '----------------------------
            '-- suite 123
            '-- 148, 75th street
            '-- red deer, ab
            '-- T2P X1X
            '----------------------------
            Lin1 = Trim(Parse(Address, 1, Chr(3)))
            Lin2 = Trim(Parse(Address, 2, Chr(3)))
            City = Trim(Parse(Parse(Address, 3, Chr(3)), 1, ","))
            Prov = Trim(Parse(Parse(Address, 3, Chr(3)), 2, ","))
            Pstl = Trim(Parse(Address, 4, Chr(3)))
    
    End Select
    
End Sub
Private Sub PostJobsToAccounting(WhereClause As String)
On Error GoTo eh
    Dim s        As String
    Dim rs       As Recordset
    
    s = ""
    s = s & "select distinct tbljobs.job_no" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "  left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "  left outer join tbljobs on invoiceitems.DivisionID = tbljobs.DivisionID and invoiceitems.job=tbljobs.job_no" & vbCrLf
    s = s & " where tblJobs.ExternalJobID='' and invoices.DivisionID =" & HFApp.DivisionID & " and 1=1" & vbCrLf
    s = s & WhereClause & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    While Not rs.EOF
        s = "" & rs(0)
        If s <> "" Then Call HFApp.WriteJobToAccounting(s)
        rs.MoveNext
    Wend
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "PostJobsToAccounting", s)
End Sub



Private Sub FrameEnter_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Call OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub
Private Sub gEnterDist_OLEDragOver(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
    Call OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub
Private Sub gEnterTail_OLEDragOver(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
    Call OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub
Private Sub Toolbar_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub Toolbar_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Call OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub
Private Sub FrameEnter_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub gEnterDist_OLEDragDrop(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub gEnterTail_OLEDragDrop(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub


Private Sub OLEDragDrop(Data, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo eh
    Dim i As Long
    Dim f As Long
    Dim files() As String
    Dim dda As DragDropAttachment
    Dim abyData() As Byte
    
    
    If Data.GetFormat(vbCFFiles) Then
        'file dragged from desktop
        If TypeOf Data Is DataObject Or TypeOf Data Is MSComctlLib.DataObject Then
            'dropped on normal targets like form or frame
            ReDim files(Data.files.Count)
            For i = 1 To Data.files.Count
                files(i) = Data.files(i)
            Next
        ElseIf TypeOf Data Is VSFlex8Ctl.VSDataObject Then
            'dropped on flexgrid who cant follow the dang specs!
            ReDim files(Data.FileCount)
            For i = 1 To Data.FileCount
                files(i) = Data.files(i - 1)
            Next
        End If
        
    Else
        'file dragged from outlook
        Set dda = New DragDropAttachment
        Set dda.Source = Data
        ReDim files(dda.Count())
        For i = 1 To dda.Count()
            'get filename from outlook, add to list of files
            files(i) = PathAppend(DocumentPath(txtVendor.Text), CleanFileName(txtInvoice.Text), dda.FileName(i - 1))
            'save file to disk
            On Error Resume Next
            Call CreatePath("", FilePath(files(i)))
            Kill files(i)
            On Error GoTo eh
            f = FreeFile()
            Open files(i) For Binary As #f
            
            'dont do this. must save to bytarray first otherwise it corrupts the file?? maybe.
            'Put #f, , dda.Attachment(i - 1)
            
            abyData = dda.Attachment(i - 1)
            Put #f, , abyData
            
            Close #f
            
'            Debug.Print files(i)
'            MsgBox files(i)
            
        Next
    
    End If
    
    

    'attach each file
    For i = 1 To UBound(files)
        Call AddFile(InvoiceObjectID, files(i))
        Toolbar.Buttons("documents").Image = "documentsfull"
    Next
    Effect = vbDropEffectNone
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "OLEDragDrop")
End Sub

Private Sub OLEDragOver(Data, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
    Dim b As Boolean
    
    
    'b = Data.GetFormat(vbCFFiles) Or Data.GetFormat(-16239) Or Data.GetFormat(-16238) Or Data.GetFormat(-15713)

'    Dim i As Long
'    For i = -20000 To 20000
'        If Data.GetFormat(i) Then Debug.Print i
'    Next
    
    'i dont know how to tell if these are outlook files
    
    b = True
    
    Effect = IIf(b And Toolbar.Buttons("scan").Enabled, vbDropEffectCopy, vbDropEffectNone)
    
End Sub


Private Sub AddFile(ObjectID As String, FileName As String)
On Error GoTo eh
    Dim s As String
    Dim sFile As String
    Dim DFILE As String
    Dim FInfo As ClsFileInfo
    Dim i As Long
    
        
    If App.Options(DocumentRoot) = "" Then
        MsgBox "Document path must be configured in app options", vbExclamation
        Exit Sub
    End If
    
    Set FInfo = New ClsFileInfo
    FInfo.FullPathName = FileName
    
    If Not FInfo.FileExists Then
        MsgBox "file not found", vbExclamation, App.ProductName
    Else
        
        sFile = FileName
        DFILE = PathAppend(DocumentPath(txtVendor.Text), CleanFileName(txtInvoice.Text), FileTitle(FileName))
        
        If sFile <> DFILE Then
            On Error Resume Next
            Call CreatePath("", FilePath(DFILE))
            Call FileCopy(sFile, DFILE)
            On Error GoTo eh
        End If
        
        s = ""
        s = s & "insert into attachments(objectid,filename,modifieddate,createddate,attacheddate,embedded,weblocation)" & vbCrLf
        s = s & "values(" & DbQuote(Str, ObjectID) & vbCrLf
        s = s & "      ," & DbQuote(Str, DFILE) & vbCrLf
        s = s & "      ," & DbQuote(DateTime, FInfo.ModifyTime) & vbCrLf
        s = s & "      ," & DbQuote(DateTime, FInfo.CreationTime) & vbCrLf
        s = s & "      ,getdate()" & vbCrLf
        s = s & "      ,0" & vbCrLf
        s = s & "      ,'Documents')"
        Call HFApp.SqlExec(s, dbHomeFront)
        
    End If
    
        
Exit Sub
eh: If InStr(1, err.Description, "duplicate", vbTextCompare) Then
    Else
        Call ErrHandler(SRCFILE & "AddFile")
    End If
End Sub
Private Sub ControlPanel_BarRightClick(bar As vbalExplorerBarLib6.cExplorerBar)
    Select Case bar.Key
        Case "REPORTS", "TOOLS", "HELP"
            mnuControlPanel.tag = bar.Key
            PopupMenu mnuControlPanel
    End Select
End Sub
Private Sub ControlPanel_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
On Error GoTo eh
    Dim DivID As String
    Dim File As String
    Dim s As String

    Dim Invoice As String
    Dim Vendor As String


    Select Case Parse(itm.Key, 1, ":")
        
        Case "SELECTINVOICESTOPAY"
            FSelectInvoicesToPay.Show vbModal
            
        Case "WALLETBATCHSTATUS"
            s = ""
            s = s & "select top 150 WalletBatchID ,WalletBatchID [Batch Number], TStmp [Date], CompanyID, Status, BatchJsonReq [Batch Request], BatchJsonRes [Batch Response], PaymentJsonReq [Payment Request], PaymentJsonRes [Payment Response]" & vbCrLf
            s = s & "from Wallet_Batch" & vbCrLf
            s = s & "order by 1 desc" & vbCrLf
            Call FDBGrid.ShowForm("Wallet Posting Batches", s, "Wallet_Batch", "WalletBatchID", True, "WalletBatchID", , False, False, True, "Wallet_Payments")

        Case "EXPORTWALLETPAYEES"
            s = ""
            s = s & "select p.WalletID PayeeID, p.PayeeID VendorCode, p.Name, p.Address1, p.Address2, p.City, p.State, p.Postal" & vbCrLf
            s = s & "from divisions d" & vbCrLf
            s = s & "join datasources s on d.bookofaccount=s.bookofaccount" & vbCrLf
            s = s & "join wallet_payees p on s.datasourceid=p.datasourceid" & vbCrLf
            s = s & "where d.divisionid=" & HFApp.DivisionID & vbCrLf
            s = s & "and p.inactive=0" & vbCrLf
            s = s & "order by p.Name" & vbCrLf
            Call FDataExport.ExportData(s, "Export Payees")
            
        Case "WALLETBATCHES"
            FWalletBatches.ShowForm

        Case "SWITCHDIV"
            s = "select  d.DivisionID,d.DivisionCode, d.DivisionName from Divisions d left outer join divisionusers u on u.DivisionID = d.DivisionID where u.userID = " & DbQuote(Str, HFApp.LoginID) & " or u.userid is null order by 2"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Select Division", s, HFApp.DivisionID, True, False, , "DivisionID") Then
                Call HFApp.SetDivision(FPickList.SelectedItem("DivisionID"))
                Call HFApp.Options.ReadData
                Call SetWindowTitle
            End If
            

        Case "GENERATEINVOICES"
            FGenerateInvoices.Show vbModal

        Case "HFVENDORS"
            If App.EditVendors Then
                Call HFApp.EditVendor(Me.txtVendor.tag)
            Else
                MsgBox "Your security permissions don't allow you to do that. Sorry.", vbInformation, App.ProductName
            End If
            
        Case "JOURNAL"
            Call HFApp.RunTask("JournalEntries")

        Case "BACKCHARGES"
            If HFApp.Options(AccountingSystem) = asTimberline Then
                Call CheckBackChargePOs
            Else
                MsgBox "not yet implemented for non-Timberline accounting.", vbInformation, App.ProductName
            End If

        Case "IMPORTINVOICE"
            If Dirty Then
                Select Case MsgBox("This invoice has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, App.ProductName)
                    Case vbCancel: Exit Sub
                    Case vbYes:    If Not SaveInvoice() Then Exit Sub
                End Select
            End If
'            If FImportInvoice.ImportInvoice(Vendor, Invoice, dteAccounting.Value) Then
'                Call OpenInvoice(Vendor, Invoice, True)
'                FrameEnter.Visible = True
'                FrameView.Visible = False
'                Call Form_Resize
'            End If
            Call FImport.ShowForm

        Case "POCOMPLETIONS"
            FPOCompletions.Show vbModal

        Case "ENTERINVOICES"
            If Dirty Then
                Select Case MsgBox("This invoice has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, App.ProductName)
                    Case vbCancel: Exit Sub
                    Case vbYes:    If Not SaveInvoice() Then Exit Sub
                End Select
            End If
            Call SetWindowTitle
            txtVendor.Text = ""
            Call OpenInvoice("", "")
            FrameEnter.Visible = True
            FrameView.Visible = False
            Call Form_Resize
            On Error Resume Next

        Case "VIEWPENDING", "VIEWAPPROVED", "VIEWHOLDBACK", "VIEWREJECTED", "VIEWHELD", "VIEWLIENHOLD", "VIEWLIENVOUCHER"
            If Dirty Then
                Select Case MsgBox("This invoice has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, App.ProductName)
                    Case vbCancel: Exit Sub
                    Case vbYes:    If Not SaveInvoice() Then Exit Sub
                End Select
            End If
            Call OpenInvoice("", "")
            gViewInv.Rows = 1
            gViewDist.Rows = 1
            Call SetWindowTitle
            FrameEnter.Visible = False
            FrameView.Visible = True
            FrameSearch.Visible = False
            Call Form_Resize
            Select Case Parse(itm.Key, 1, ":")
                Case "VIEWLIENHOLD":     s = "Status in('Lien Hold')"
                Case "VIEWLIENVOUCHER":  s = "Status in('Lien Voucher')"
                'Case "VIEWPENDING":      s = "Status in('Pending')                  and isnull(RetainageInvoice,0)=0"
                'Case "VIEWAPPROVED":     s = "Status in('Approved','SiteApproved')  and isnull(RetainageInvoice,0)=0"
                'Case "VIEWHOLDBACK":     s = "not Status in('Posted')               and isnull(RetainageInvoice,0)=1"
                
                Case "VIEWPENDING":      s = "Status in('Pending')"
                Case "VIEWAPPROVED":     s = "Status in('Approved','SiteApproved')"
                Case "VIEWHOLDBACK":     s = "Status in('Holdback')"
                
                Case "VIEWHELD":         s = "Status in('Hold')"
                Case "VIEWREJECTED":     s = "Status in('Rejected')"
                                         Call VerifyPostedInvoices
            End Select
            Call FindInvoices(s)


        Case "INVOICESEARCH"
            If Dirty Then
                Select Case MsgBox("This invoice has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, App.ProductName)
                    Case vbCancel: Exit Sub
                    Case vbYes:    If Not SaveInvoice() Then Exit Sub
                End Select
            End If
            Call OpenInvoice("", "")
            gViewInv.Rows = 1
            gViewDist.Rows = 1
            Call SetWindowTitle
            FrameEnter.Visible = False
            FrameView.Visible = True
            FrameSearch.Visible = True
            Call Form_Resize

        Case "VIEWPRINTFILE"
            If App.PostInvoices Then
            If PathExists(App.Options(ImportPrintFile)) Then
                File = PathAppend(HFApp.Options(Timberline_Data_Path), ForceExt(HFApp.LoginID, FileExt(App.Options(ImportPrintFile))))
                On Error Resume Next
                Call FileCopy(App.Options(ImportPrintFile), File)
                On Error GoTo eh
                If LCase(FileExt(App.Options(ImportPrintFile))) = "prn" Then
                    Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), HFApp.Options(Timberline_UID), HFApp.Options(Timberline_PWD), File)
                Else
                    Call ShellFile(Me.hwnd, File)
                End If
            ElseIf PathExists(ForceExt(App.Options(ImportPrintFile), "txt")) Then
                Call ShellFile(Me.hwnd, ForceExt(App.Options(ImportPrintFile), "txt"))
            Else
                Call MsgBox("Sage300 print file could not be found." & vbCrLf & vbCrLf & App.Options(ImportPrintFile), vbInformation, App.ProductName)
            End If
            End If

        Case "POSTPENDING":       If App.PostInvoices Then Call ShowPostMenu
        
        Case "APPOPTIONS":
            If App.Admin Then
                Call FOptions.Show(vbModal, Me)
                Call ReadSettings
                Call Form_Resize
            Else
                MsgBox "Your security permissions don't allow you to do that. Sorry.", vbInformation, App.ProductName
            End If
            
        Case "USERSECURITY":
            Call HFApp.RunTask("EditSecurity")
        
        Case "HFOPTIONS":
            Call HFApp.EditOptions
            Call LoadControlPanel
            Call GetFormatMasks
            Call LoadDepartments
            
            
        Case "APPABOUT":          Call HFApp.About

        Case "TOOL", "HELP":
            s = itm.tag
            If InStr(1, s, ".exe ", vbTextCompare) Then
                Call Shell(s, vbNormalFocus)
            Else
                If Not PathExists(s) Then Call CreatePath("the folder", s)
                Call ShellFile(Me.hwnd, s)
            End If

        Case "REPORT"
            'Set crv = New HFPrinter.ReportViewer
            
            'Call crv.ShowReport(HFApp.ConnectionString(dbHomeFront), itm.tag, rvPreview, "", "", "User Name", HFApp.LoginID, "DivisionID", HFApp.divisionid)

            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(itm.tag, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            Call c.ParameterValue("User Name", HFApp.LoginID)
            On Error GoTo eh
            Call c.PrintPreview("Print Preview")
            
            
    End Select
Exit Sub
eh: Call ErrHandler(SRCFILE & "ControlPanel_ItemClick(""" & itm.Text & """)")
End Sub



Private Sub LoadDepartments()
    Call LoadComboBox(cboDepartment, HFApp.Databases(dbHomeFront), "select '','',0 union select name,'',deptid from departments where divisionid=" & DbQuote(num, HFApp.DivisionID))
End Sub

Private Sub GetApprover(Optional Job As String, Optional DepartmentID As Long)
Static bInHere As Boolean
If bInHere Then Exit Sub
bInHere = True
    
    Dim s As String
    Dim rs As Recordset
    Dim DeptID As Long
    Dim pm As String
    
    If Job <> "" Then
        If cboApprover.Text <> "" Then
        
            Exit Sub
        End If
        
        s = ""
        s = s & "select top 1 d.deptid,isnull(nullif(j.pm,''),l.project_manager) pm" & vbCrLf
        s = s & "from tbljobs j" & vbCrLf
        s = s & "left outer join tbllocality l on j.community=l.area" & vbCrLf
        s = s & "left outer join departmentapprovers da on da.pm=isnull(nullif(j.pm,''),l.project_manager)" & vbCrLf
        s = s & "left outer join departments d on da.deptid=d.deptid and d.divisionid=j.divisionid" & vbCrLf
        s = s & "where j.job_no=" & DbQuote(Str, StripFormating(Job)) & vbCrLf
        s = s & "order by da.invoicelimit desc" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        If rs.EOF Then
            cboDepartment.ListIndex = -1
            cboApprover.ListIndex = -1
        Else
            DeptID = Val("" & rs("deptid"))
            pm = "" & rs("pm")
        
            s = ""
            s = s & "select deptid,pm" & vbCrLf
            s = s & "from departmentapprovers " & vbCrLf
            s = s & "where deptid=" & DbQuote(num, DeptID) & vbCrLf
            s = s & "order by sortorder desc" & vbCrLf
            Set rs = HFApp.SqlExec(s)
            cboApprover.Clear
            While Not rs.EOF
                cboApprover.AddItem "" & rs("PM")
                rs.MoveNext
            Wend
            Call SetComboBoxListIndex(cboDepartment, , , DeptID)
            Call SetComboBoxListIndex(cboApprover, pm)
        End If
    Else
        s = ""
        s = s & "select deptid,pm" & vbCrLf
        s = s & "from departmentapprovers " & vbCrLf
        s = s & "where deptid=" & DbQuote(num, DepartmentID) & vbCrLf
        s = s & "order by sortorder desc" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        cboApprover.Clear
        While Not rs.EOF
            cboApprover.AddItem "" & rs("PM")
            rs.MoveNext
        Wend
    End If


On Error Resume Next
cboApprover.ListIndex = 0
bInHere = False
End Sub

Private Sub WriteApprovalLog(ByVal InvoiceIDs_CSV As String, Status As String)
    Dim s As String
    Dim i As Long
    Dim ID As String
    
    For i = 1 To Parse(InvoiceIDs_CSV)
        ID = Val("" & Parse(InvoiceIDs_CSV, i))
        If ID <> 0 Then
            s = "exec Approvals_WriteLog " & DbQuote(num, HFApp.DivisionID) & "," & DbQuote(num, ID) & "," & DbQuote(Str, HFApp.LoginID) & "," & DbQuote(Str, Status)
            Call HFApp.SqlExec(s)
        End If
    Next

End Sub


Private Sub SetDocButtonImage()
On Error GoTo eh
    Dim s As String
    
    s = "select count(*) from attachments where objectid=" & DbQuote(Str, InvoiceObjectID)
    If Val("" & HFApp.SqlExec(s)(0)) = 0 Then
        Toolbar.Buttons("documents").Image = "documents"
    Else
        Toolbar.Buttons("documents").Image = "documentsfull"
    End If

Exit Sub
eh: Call ErrHandler(SRCFILE & "gViewInv_BeforeMouseDown")
End Sub



Private Sub PostToXero(WhereClause As String)
On Error GoTo eh

    Dim es As String
    Dim s        As String
    Dim rs       As Recordset
    Dim lastID   As Long
    Dim Qty As Double, Pretax As Double
    Dim Rate As Double, Tax As Double, tax2 As Double

    Dim XeroJobCatName As String
    Dim XeroCostCodeCatName As String
    XeroJobCatName = HFApp.Options.ValueByName("XeroJobName")
    XeroCostCodeCatName = HFApp.Options.ValueByName("XeroCostCodeName")
    Dim xml As String

    
    Screen.MousePointer = vbHourglass




    'open invoices query
es = "querying homefront db"
    s = ""
    s = s & "select Invoices.InvoiceID" & vbCrLf
    s = s & "      ,Invoices.Vendor" & vbCrLf
    s = s & "      ,Invoices.VendorType" & vbCrLf
    s = s & "      ,Invoices.Invoice" & vbCrLf
    s = s & "      ,Invoices.Description           InvoiceDescription" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode1" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode2" & vbCrLf
    s = s & "      ,Invoices.PreTax                InvoicePretax" & vbCrLf
    s = s & "      ,Invoices.Tax                   InvoiceTax" & vbCrLf
    s = s & "      ,Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & "      ,Invoices.Discount              InvoiceDiscount" & vbCrLf
    s = s & "      ,Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & "      ,Invoices.ReceivedDate          InvoiceReceivedDate" & vbCrLf
    s = s & "      ,Invoices.DiscountDate          InvoiceDiscountDate" & vbCrLf
    s = s & "      ,Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & "      ,Invoices.AccountingDate        InvoiceAccountingDate" & vbCrLf
    s = s & "      ,Invoices.VendorName            InvoiceVendorName" & vbCrLf
    s = s & "      ,invoiceitems.ItemID            ItemID" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentVendor  ItemCommitmentVendor" & vbCrLf
    s = s & "      ,invoiceitems.Commitment        ItemCommitment" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentItem    ItemCommitmentItem" & vbCrLf
    s = s & "      ,invoiceitems.Job               ItemJob" & vbCrLf
    s = s & "      ,tbljobs.ExternalJobID          ItemProject" & vbCrLf
    s = s & "      ,invoiceitems.Extra             ItemExtra" & vbCrLf
    's = s & "      ,c.externalid                   ItemPhase" & vbCrLf
    s = s & "      ,invoiceitems.Phase             ItemPhase" & vbCrLf
    s = s & "      ,sc.externalid                  ItemCategory" & vbCrLf
    
    's = s & "      ,invoiceitems.DebitAccount      ItemDebitAccount" & vbCrLf
    s = s & "      ,isnull(nullif(isnull(nullif(invoiceitems.DebitAccount,''),sc.DebitAccount),''),c.DebitAccount) ItemDebitAccount" & vbCrLf
    
    s = s & "      ,invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedQuantity  ItemQuantity" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedUnitPrice ItemUnitPrice" & vbCrLf
    s = s & "      ,invoiceitems.PreTax            ItemPretax" & vbCrLf
    s = s & "      ,invoiceitems.Tax               ItemTax" & vbCrLf
    s = s & "      ,invoiceitems.TaxRate " & vbCrLf
    s = s & "      ,invoiceitems.Description       ItemDescription" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "       left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "       left outer join tbljobs on invoiceitems.DivisionID = tbljobs.DivisionID and invoiceitems.job=tbljobs.job_no" & vbCrLf
    s = s & "       LEFT OUTER JOIN standardcostcodes c ON(invoiceitems.DivisionID = c.DivisionID and invoiceitems.phase=c.costcode)" & vbCrLf
    s = s & "       LEFT OUTER JOIN standardcategories sc ON(invoiceitems.DivisionID = sc.DivisionID and invoiceitems.category=sc.category)" & vbCrLf
    s = s & " where invoices.DivisionID =" & HFApp.DivisionID & " and 1=1" & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "order by Invoices.InvoiceID,invoiceitems.ItemID"
    Set rs = HFApp.SqlExec(s)

    lastID = -999
    

    
    While Not rs.EOF

        If lastID <> rs("InvoiceID") Then

            If lastID <> -999 Then
                es = "saving to Xero"
                xml = xml & "  </LineItems>" & vbCrLf
                xml = xml & "</Invoice>" & vbCrLf
                s = SubmitBrokerXml("Xero", "PostInvoice", xml)
                
                'TODO: check for error response
                If s = "" Then
                    es = "changing hf invoice status"
                    Call SetInvoiceStatus(lastID, "Posted")
                End If
            End If

            lastID = rs("InvoiceID")
            es = "creating Xero invoice"
            xml = ""
            xml = xml & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
            xml = xml & "<Invoice>" & vbCrLf
            xml = xml & "  <ContactID>" & rs("Vendor") & "</ContactID>" & vbCrLf
            xml = xml & "  <InvoiceNumber>" & rs("Invoice") & "</InvoiceNumber>" & vbCrLf
            xml = xml & "  <Date>" & Format("" & rs("invoiceDate"), "yyyy-mm-dd") & "</Date>" & vbCrLf
            If "" & rs("InvoicePaymentDate") <> "" Then
                xml = xml & "  <DueDate>" & Format("" & rs("InvoicePaymentDate"), "yyyy-mm-dd") & "</DueDate>" & vbCrLf
            End If
            xml = xml & "  <Reference>" & left("" & rs("Invoice"), 255) & "</Reference>" & vbCrLf
            xml = xml & "  <Status>AUTHORISED</Status>" & vbCrLf
            xml = xml & "  <LineItems>" & vbCrLf
            
        End If

        'write invoiceitem
        If Val("" & rs("ItemID")) <> 0 Then
            Pretax = Val("" & rs("ItemPretax"))
            Tax = Val("" & rs("ItemTax"))
            
            es = "adding line to xero invoice"
            xml = xml & "    <LineItem>" & vbCrLf
            xml = xml & "      <Description>" & rs("ItemDescription") & "</Description>" & vbCrLf
            xml = xml & "      <UnitAmount>" & Pretax & "</UnitAmount>" & vbCrLf
            xml = xml & "      <Quantity>" & Val("" & rs("ItemQuantity")) & "</Quantity>" & vbCrLf
            xml = xml & "      <AccountCode>" & rs("ItemDebitAccount") & "</AccountCode>" & vbCrLf
            xml = xml & "      <TaxType>" & rs("ItemTaxGroup") & "</TaxType>" & vbCrLf
            xml = xml & "      <TaxAmount>" & Tax & "</TaxAmount>" & vbCrLf
'            If XeroJobCatName <> "" Then
'                xml = xml & "      <TrackingCategory>" & vbCrLf
'                xml = xml & "        <Name>" & XeroJobCatName & "</Name>" & vbCrLf
'                xml = xml & "        <Option>" & rs("ItemJob") & "</Option>" & vbCrLf
'                xml = xml & "      </TrackingCategory>" & vbCrLf
'            End If
'            If XeroCostCodeCatName <> "" Then
'                xml = xml & "      <TrackingCategory>" & vbCrLf
'                xml = xml & "        <Name>" & XeroCostCodeCatName & "</Name>" & vbCrLf
'                xml = xml & "        <Option>" & rs("ItemPhase") & "</Option>" & vbCrLf
'                xml = xml & "      </TrackingCategory>" & vbCrLf
'            End If
            xml = xml & "    </LineItem>" & vbCrLf

        End If

        rs.MoveNext
    Wend
    If lastID <> -999 Then
        
        es = "saving to Xero"
        xml = xml & "  </LineItems>" & vbCrLf
        xml = xml & "</Invoice>" & vbCrLf
        s = SubmitBrokerXml("Xero", "PostInvoice", xml)
        
        'TODO: check for error response
        If s = "" Then
            es = "changing hf invoice status"
            Call SetInvoiceStatus(lastID, "Posted")
        End If
        
    End If

    Screen.MousePointer = vbDefault


Exit Sub
eh: Call ErrHandler(SRCFILE & "PostXero", es)
End Sub



Private Sub SetWindowTitle()
    Me.Caption = App.ProductName & " (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & " Division: " & "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & HFApp.DivisionID, dbHomeFront)(0)
End Sub


Private Sub dteInvoice_Change()
On Error Resume Next
    Dim d As Date
    
    Dirty = True
    If Not IsDate(dteInvoice) Then Exit Sub
    If App.Options(AcctDateDefaultsToInvDate) And mInvoiceID = 0 Then dteAccounting.Value = dteInvoice.Value
    If HFApp.Options(AccountingSystem) <> asTimberline Then dteAccounting = dteInvoice

    If Not App.Options(CalcDatesFromInvoiceDate) Then Exit Sub
    
    'discount date
    If mDiscPercent > 0 Then
        If mDiscDays = 0 Then
            dteDiscount = ""
        Else
            d = dteInvoice
            Select Case mDiscType
            
                Case "InAGivenNumberOfDays"
                    dteDiscount = DateAdd("d", mDiscDays, d)
                    
                Case "NumberOfDaysAfterEOM"
                    dteDiscount = DateAdd("d", mDiscDays, DateSerial(Year(d), Month(d) + 1, 1))
                
                Case "OnADayOfTheMonth"
                    If d >= DateSerial(Year(d), Month(d), mDiscDays) Then
                        d = DateSerial(Year(d), Month(d) + 1, mDiscDays)
                    Else
                        d = DateSerial(Year(d), Month(d), mDiscDays)
                    End If
                    dteDiscount = d
                    
                Case "DayOfMonthAfterEOM"
                    dteDiscount = DateSerial(Year(d), Month(d) + 1, mDiscDays)
            
            End Select
        End If
    Else
        dteDiscount = ""
    End If
    
    'payment date
    If mNetDays > 0 And Not App.Options(DontSetPaymentDate) Then
        d = dteInvoice
        Select Case mNetType
        
            Case "InAGivenNumberOfDays"
                dtePayment = DateAdd("d", mNetDays, d)
                
            Case "NumberOfDaysAfterEOM"
                dtePayment = DateAdd("d", mNetDays, DateSerial(Year(d), Month(d) + 1, 1))
            
            Case "OnADayOfTheMonth"
                If d >= DateSerial(Year(d), Month(d), mNetDays) Then
                    d = DateSerial(Year(d), Month(d) + 1, mNetDays)
                Else
                    d = DateSerial(Year(d), Month(d), mNetDays)
                End If
                dtePayment = d
                
            Case "DayOfMonthAfterEOM"
                dtePayment = DateSerial(Year(d), Month(d) + 1, mNetDays)
        
        End Select
    End If
End Sub

Private Sub dteReceived_Change()
On Error Resume Next
    Dim d As Date
    
    Dirty = True
    If Not IsDate(dteReceived) Then Exit Sub
    If App.Options(AcctDateDefaultsToRecDate) And mInvoiceID = 0 Then dteAccounting.Value = dteReceived.Value
    If App.Options(CalcDatesFromInvoiceDate) Then Exit Sub
    
    'discount date
    If mDiscPercent > 0 Then
        If mDiscDays = 0 Then
            dteDiscount = ""
        Else
            d = dteReceived
            Select Case mDiscType
            
                Case "InAGivenNumberOfDays"
                    dteDiscount = DateAdd("d", mDiscDays, d)
                    
                Case "NumberOfDaysAfterEOM"
                    dteDiscount = DateAdd("d", mDiscDays, DateSerial(Year(d), Month(d) + 1, 1))
                
                Case "OnADayOfTheMonth"
                    If d >= DateSerial(Year(d), Month(d), mDiscDays) Then
                        d = DateSerial(Year(d), Month(d) + 1, mDiscDays)
                    Else
                        d = DateSerial(Year(d), Month(d), mDiscDays)
                    End If
                    dteDiscount = d
                    
                Case "DayOfMonthAfterEOM"
                    dteDiscount = DateSerial(Year(d), Month(d) + 1, mDiscDays)
            
            End Select
        End If
    Else
        dteDiscount = ""
    End If

    
    'payment date
    If mNetDays > 0 And Not App.Options(DontSetPaymentDate) Then
        d = dteReceived
        Select Case mNetType
        
            Case "InAGivenNumberOfDays"
                dtePayment = DateAdd("d", mNetDays, d)
                
            Case "NumberOfDaysAfterEOM"
                dtePayment = DateAdd("d", mNetDays, DateSerial(Year(d), Month(d) + 1, 1))
            
            Case "OnADayOfTheMonth"
                If d >= DateSerial(Year(d), Month(d), mNetDays) Then
                    d = DateSerial(Year(d), Month(d) + 1, mNetDays)
                Else
                    d = DateSerial(Year(d), Month(d), mNetDays)
                End If
                dtePayment = d
                
            Case "DayOfMonthAfterEOM"
                dtePayment = DateSerial(Year(d), Month(d) + 1, mNetDays)
        
        End Select
    End If
End Sub

Private Function LienReleaseReceived(VoucherIDs As String) As Boolean
    If VoucherIDs = "" Then
        LienReleaseReceived = False
        Exit Function
    End If
    
    Dim s As String
    
    s = ""
    s = s & "update LienVouchers set" & vbCrLf
    s = s & " releaseSignedDate=getdate()" & vbCrLf
    s = s & ",releaseSignedVia='PayablesDesk'" & vbCrLf
    s = s & ",releaseSignedBy=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "where Voucher in(" & VoucherIDs & ")" & vbCrLf
    Call HFApp.SqlExec(s)
        
    s = ""
    s = s & "update Invoices" & vbCrLf
    s = s & "set Status='Approved'" & vbCrLf
    s = s & "where LienVoucher in(" & VoucherIDs & ")" & vbCrLf
    Call HFApp.SqlExec(s)
    
    
    LienReleaseReceived = True
    

End Function

Private Function CreateLienVouchers(InvoiceIDs As String) As Boolean
    
    If InvoiceIDs = "" Then
        CreateLienVouchers = False
        Exit Function
    End If
    
    Dim s As String
    Dim batch As Long
    
    s = "INSERT INTO Batches(BatchType,UStmp,TStmp,DivisionID) VALUES('Lien Vouchers'," & DbQuote(Str, HFApp.LoginID) & ",GETDATE()," & HFApp.DivisionID & ")"
    Call HFApp.SqlExec(s)
    batch = HFApp.SqlIdentity("Batches")
    
    s = ""
    s = s & "insert LienVouchers(batch,Vendor,createddate,createdby)" & vbCrLf
    s = s & "select distinct " & DbQuote(num, batch) & ",i.Vendor,getdate()," & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "from invoices i" & vbCrLf
    s = s & "where isnull(i.lienvoucher,'')=''" & vbCrLf
    s = s & "and i.invoiceid in(" & InvoiceIDs & ")" & vbCrLf
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "update i" & vbCrLf
    s = s & "set LienVoucher=v.Voucher" & vbCrLf
    s = s & ",status='Lien Voucher'" & vbCrLf
    s = s & "from Invoices i" & vbCrLf
    s = s & "join LienVouchers v on i.vendor=v.vendor and v.batch=" & DbQuote(num, batch) & vbCrLf
    s = s & "where isnull(i.lienvoucher,'')=''" & vbCrLf
    s = s & "and i.invoiceid in(" & InvoiceIDs & ")" & vbCrLf
    Call HFApp.SqlExec(s)
    
    
    CreateLienVouchers = True
    
    
End Function


Private Sub PostToSage100(WhereClause As String)
On Error GoTo eh

    Dim s        As String
    Dim rs       As Recordset
    Dim lastID   As Long
    Dim lastInvoiceNumber As String
    Dim MiscDeductionRate  As Double
    Dim MiscDeductionAmt  As Double
    Dim X As String
    Dim POType As String
    Dim subID As Long
    Dim SubRefNumber As String
    Dim TaxLineString As String

    Dim Qty As Double
    Dim Rate As Double
    Dim Amount As Double

    Dim PostingYear As Integer
    Dim PostingPeriod As Integer
    
    Dim PostPOQtyToAccounting As Boolean
    PostPOQtyToAccounting = HFApp.Options.ValueByName("PostPOQtyToAccounting") = "true"

    'invoice is not valid if it includes multiple POs, or if some but not all lines reference a PO
    s = ""
    s = s & "select count(distinct isnull(invoiceitems.commitment,''))" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "where 1=1" & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "group by invoices.invoice" & vbCrLf
    s = s & "having count(distinct isnull(invoiceitems.commitment,''))>1" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then
        MsgBox "Unable to post. An invalid Invoice has been detected." & vbCrLf & vbCrLf & "A Sage 100 invoice is not valid if it includes multiple POs or if some but not all distributions reference a PO.", vbExclamation, App.ProductName
        Exit Sub
    End If


    'get current period
    s = "SELECT [fscyrd],[curprd] From [dbo].[lgrset]"
    Set rs = HFApp.SqlExec(s, dbAccounting)
    PostingYear = Val("" & rs(0))
    PostingPeriod = Val("" & rs(1))
    
    'open invoices query
    s = ""
    s = s & "--PO invoices" & vbCrLf
    s = s & "select Invoices.InvoiceID" & vbCrLf
    s = s & "      ,pi.ponumber,pi.linenumber" & vbCrLf
    s = s & "      ,Invoices.Vendor" & vbCrLf
    s = s & "      ,tblvendors.tradetype VendorType" & vbCrLf
    s = s & "      ,Invoices.Invoice" & vbCrLf
    s = s & "      ,left(Invoices.Description,30) InvoiceDescription" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode1" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode2" & vbCrLf
    s = s & "      ,Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & "      ,Invoices.Tax                   InvoiceTax" & vbCrLf
    s = s & "      ,Invoices.Discount              InvoiceDiscount" & vbCrLf
    s = s & "      ,Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & "      ,Invoices.ReceivedDate          InvoiceReceivedDate" & vbCrLf
    s = s & "      ,Invoices.DiscountDate          InvoiceDiscountDate" & vbCrLf
    s = s & "      ,Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & "      ,Invoices.AccountingDate        InvoiceAccountingDate" & vbCrLf
    s = s & "      ,Invoices.VendorName            InvoiceVendorName" & vbCrLf
    s = s & "      ,Invoices.VendorAddr1           InvoiceVendorAddr1" & vbCrLf
    s = s & "      ,Invoices.VendorAddr2           InvoiceVendorAddr2" & vbCrLf
    s = s & "      ,Invoices.VendorCity            InvoiceVendorCity" & vbCrLf
    s = s & "      ,Invoices.VendorProv            InvoiceVendorProv" & vbCrLf
    s = s & "      ,Invoices.VendorPostal          InvoiceVendorPostal" & vbCrLf
    s = s & "      ,invoiceitems.ItemID            ItemID" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentVendor  ItemCommitmentVendor" & vbCrLf
    s = s & "      ,invoiceitems.Commitment        ItemCommitment" & vbCrLf
    s = s & "      ,pi.linenumber    ItemCommitmentItem" & vbCrLf
    s = s & "      ,pi.Job               ItemJob" & vbCrLf
    s = s & "      ,pi.jcExtra             ItemExtra" & vbCrLf
    s = s & "      ,pi.jccostcode             ItemPhase" & vbCrLf
    s = s & "      ,pi.jcCategory          ItemCategory" & vbCrLf
    s = s & "      ,isnull(nullif(invoiceitems.DebitAccount,''),isnull(nullif(isnull(nullif(cc.debitaccount,''),ca.debitaccount),''),tblvendors.debitaccount)) ItemDebitAccount" & vbCrLf
    s = s & "      ,invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & "      ,isnull(invoiceitems.InvoicedUnitPrice,pi.rate) ItemRate" & vbCrLf
    s = s & "      ,isnull(invoiceitems.InvoicedQuantity,0)  ItemQty" & vbCrLf
    s = s & "      ,isnull(invoiceitems.PreTax,0)            ItemPretax" & vbCrLf
    s = s & "      ,isnull(invoiceitems.PreTax + invoiceitems.Tax,0)    ItemAmount" & vbCrLf
    s = s & "      ,invoiceitems.Tax               ItemTax" & vbCrLf
    s = s & "      ,round(invoiceitems.Retainage*(1+isnull(t.RetainageRate,0)/100),2)     ItemRetainage" & vbCrLf
    s = s & "      ,left(isnull(invoiceitems.Description,pi.description),75)       ItemDescription" & vbCrLf
    s = s & "      , t.JCRate                      ItemTaxRate" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "  left join (select invoiceid,max(commitment) commitment from invoiceitems ii group by invoiceid) x on x.invoiceid = invoices.invoiceid" & vbCrLf
    s = s & "  join pomaster p on x.commitment=p.ponumber " & vbCrLf
    s = s & "  left join poitems pi on p.ponumber=pi.ponumber " & vbCrLf
    s = s & "  left join invoiceitems on invoiceitems.invoiceid=invoices.invoiceid and invoiceitems.commitmentitem=pi.linenumber" & vbCrLf
    s = s & "  left outer join tblvendors   on invoices.DivisionID=tblvendors.DivisionID and invoices.vendor=tblvendors.vendor_id" & vbCrLf
    s = s & "  left outer join taxgroups  t on invoiceitems.divisionid=t.divisionid and invoiceitems.taxgroup=t.taxgroup" & vbCrLf
    s = s & "  left join standardcostcodes cc on invoices.divisionid=cc.divisionid and pi.jccostcode=cc.costcode" & vbCrLf
    s = s & "  left join standardcategories ca on invoices.divisionid=ca.divisionid and pi.jccategory=ca.category" & vbCrLf
    s = s & " where invoices.DivisionID = " & HFApp.DivisionID & " and 1=1" & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "" & vbCrLf
    s = s & "union all" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- non PO invoices" & vbCrLf
    s = s & "select Invoices.InvoiceID" & vbCrLf
    s = s & "      ,'' ponumber,invoiceitems.ItemID" & vbCrLf
    s = s & "      ,Invoices.Vendor" & vbCrLf
    s = s & "      ,tblvendors.tradetype VendorType" & vbCrLf
    s = s & "      ,Invoices.Invoice" & vbCrLf
    s = s & "      ,left(Invoices.Description,30) InvoiceDescription" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode1" & vbCrLf
    s = s & "      ,Invoices.InvoiceCode2" & vbCrLf
    s = s & "      ,Invoices.PreTax + Invoices.Tax InvoiceAmount" & vbCrLf
    s = s & "      ,Invoices.Tax                   InvoiceTax" & vbCrLf
    s = s & "      ,Invoices.Discount              InvoiceDiscount" & vbCrLf
    s = s & "      ,Invoices.InvoiceDate           InvoiceDate" & vbCrLf
    s = s & "      ,Invoices.ReceivedDate          InvoiceReceivedDate" & vbCrLf
    s = s & "      ,Invoices.DiscountDate          InvoiceDiscountDate" & vbCrLf
    s = s & "      ,Invoices.PaymentDate           InvoicePaymentDate" & vbCrLf
    s = s & "      ,Invoices.AccountingDate        InvoiceAccountingDate" & vbCrLf
    s = s & "      ,Invoices.VendorName            InvoiceVendorName" & vbCrLf
    s = s & "      ,Invoices.VendorAddr1           InvoiceVendorAddr1" & vbCrLf
    s = s & "      ,Invoices.VendorAddr2           InvoiceVendorAddr2" & vbCrLf
    s = s & "      ,Invoices.VendorCity            InvoiceVendorCity" & vbCrLf
    s = s & "      ,Invoices.VendorProv            InvoiceVendorProv" & vbCrLf
    s = s & "      ,Invoices.VendorPostal          InvoiceVendorPostal" & vbCrLf
    s = s & "      ,invoiceitems.ItemID            ItemID" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentVendor  ItemCommitmentVendor" & vbCrLf
    s = s & "      ,invoiceitems.Commitment        ItemCommitment" & vbCrLf
    s = s & "      ,invoiceitems.CommitmentItem    ItemCommitmentItem" & vbCrLf
    s = s & "      ,invoiceitems.Job               ItemJob" & vbCrLf
    s = s & "      ,invoiceitems.Extra             ItemExtra" & vbCrLf
    s = s & "      ,invoiceitems.Phase             ItemPhase" & vbCrLf
    s = s & "      ,invoiceitems.Category          ItemCategory" & vbCrLf
    s = s & "      ,invoiceitems.DebitAccount      ItemDebitAccount" & vbCrLf
    s = s & "      ,invoiceitems.TaxGroup          ItemTaxGroup" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedUnitPrice ItemRate" & vbCrLf
    s = s & "      ,invoiceitems.InvoicedQuantity  ItemQty" & vbCrLf
    s = s & "      ,invoiceitems.PreTax            ItemPretax" & vbCrLf
    s = s & "      ,invoiceitems.PreTax + invoiceitems.Tax    ItemAmount" & vbCrLf
    s = s & "      ,invoiceitems.Tax               ItemTax" & vbCrLf
    s = s & "      ,round(invoiceitems.Retainage*(1+isnull(t.RetainageRate,0)/100),2)     ItemRetainage" & vbCrLf
    s = s & "      ,left(invoiceitems.Description,75)       ItemDescription" & vbCrLf
    s = s & "      , t.JCRate                      ItemTaxRate" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "  left outer join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "  left outer join tblvendors   on invoices.DivisionID=tblvendors.DivisionID and invoices.vendor=tblvendors.vendor_id" & vbCrLf
    s = s & "  left outer join taxgroups  t on invoiceitems.divisionid=t.divisionid and invoiceitems.taxgroup=t.taxgroup" & vbCrLf
    s = s & "where isnull(invoiceitems.commitment,'')=''" & vbCrLf
    s = s & "and invoices.DivisionID = " & HFApp.DivisionID & " and 1=1" & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "" & vbCrLf
    s = s & "order by 1,2,3" & vbCrLf
    
    Set rs = HFApp.SqlExec(s)
    lastID = -999
    lastInvoiceNumber = ""
    TaxLineString = ""
    While Not rs.EOF
        If lastID <> Val("" & rs("InvoiceID")) Then
                
            'add header row
            If lastID <> -999 Then
                If TaxLineString <> "" Then
                    X = X & TaxLineString
                End If
                X = X & "</APInvoiceAddRq>" & vbCrLf
                X = X & HFApp.XmlMBEnd()
                On Error Resume Next
                s = HFApp.XmlMbSubmit(X, HFApp.Options(MasterBuilderPWD))
                If err.Number = 0 Then
                    Call HFApp.SqlExec("update Invoices set Status='Posted', postingdate=getdate() where InvoiceID=" & DbQuote(num, lastID), dbHomeFront)
                     
                Else
                    MsgBox "Unable to post invoice " & rs("Invoice") & vbCrLf & vbCrLf & err.Description, vbCritical, App.ProductName
                End If
                On Error GoTo eh
            End If
        
        
            lastID = Val("" & rs("InvoiceID"))
            lastInvoiceNumber = "" & rs("Invoice")
            TaxLineString = ""
            MiscDeductionRate = Val("" & HFApp.SqlExec("SELECT MiscDeductionRate FROM tblVendors WHERE DivisionID =" & HFApp.DivisionID & " and vendor_ID=" & DbQuote(Str, "" & rs("Vendor")), dbHomeFront)(0))
            
            If "" & rs("ItemCommitment") <> "" Then
                s = "Select p.POType from POmaster m join tblPOIndex p on p.DivisionID = m.DivisionID and  p.POIndex = m.POIndex where m.divisionID = " & HFApp.DivisionID & " and m.PONumber = " & DbQuote(Str, "" & rs("ItemCommitment"))
                POType = "" & HFApp.SqlExec(s, dbHomeFront)(0)
                If POType = "" Then POType = "Purchase Order"
            End If
            
            MiscDeductionAmt = Val("" & rs("InvoiceAmount")) * MiscDeductionRate / 100
            X = HFApp.XmlMbStart(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
            lastID = Val("" & rs("InvoiceID"))
            X = X & "<APInvoiceAddRq requestID=""1"">" & vbCrLf
            X = X & vbTab & HFApp.XmlMBAdd(c, 20, "InvoiceNumber", "" & rs("Invoice"))
            If POType = "Purchase Order" Then
                X = X & vbTab & HFApp.XmlMBAdd(c, 20, "PurchaseOrderNumber", "" & rs("ItemCommitment"))
            ElseIf POType = "Subcontract" Then
                X = X & vbTab & HFApp.XmlMBAdd(c, 20, "SubcontractNumber", "" & rs("ItemCommitment"))
                subID = HFApp.SqlExec("Select recnum from subcon where ctcnum = " & DbQuote(Str, "" & rs("ItemCommitment")), dbAccounting)(0)
            End If
            X = X & vbTab & HFApp.XmlMBAdd(n, 10, "VendorRef", "" & rs("Vendor"))
            X = X & vbTab & HFApp.XmlMBAdd(n, 10, "JobRef", "" & rs("ItemJob"))
            If Val("" & HFApp.SqlExec("select case count(distinct jcextra) when 1 then max(jcextra) else '' end from poitems where ponumber=" & DbQuote(Str, "" & rs("ItemCommitment")), dbHomeFront)(0)) <> 0 Then
                X = X & vbTab & HFApp.XmlMBAdd(n, 10, "PhaseRef", Val("" & HFApp.SqlExec("select case count(distinct jcextra) when 1 then max(jcextra) else '' end from poitems where ponumber=" & DbQuote(Str, "" & rs("ItemCommitment")), dbHomeFront)(0)))
            End If
            If "" & rs("InvoiceDescription") = "" Then
                X = X & vbTab & HFApp.XmlMBAdd(c, 50, "Desc", "Invoice for Job " & rs("ItemJob"))
            Else
                X = X & vbTab & HFApp.XmlMBAdd(c, 50, "Desc", "" & rs("InvoiceDescription"))
            End If
            X = X & vbTab & HFApp.XmlMBAdd(d, 0, "InvoiceDate", Format("" & rs("InvoiceDate"), "yyyy-mm-dd"))
        
            If "" & rs("InvoicePaymentDate") = "" Then
                X = X & vbTab & HFApp.XmlMBAdd(d, 0, "DueDate", Format("" & rs("InvoiceDate"), "yyyy-mm-dd"))
            Else
                X = X & vbTab & HFApp.XmlMBAdd(d, 0, "DueDate", Format("" & rs("InvoicePaymentDate"), "yyyy-mm-dd"))
            End If
            X = X & vbTab & HFApp.XmlMBAdd(c, 20, "ReferenceNum", "")
            If "" & rs("InvoiceDiscountDate") <> "" Then
                X = X & vbTab & HFApp.XmlMBAdd(d, 0, "DiscountDate", Format("" & rs("InvoiceDiscountDate"), "yyyy-mm-dd"))
            Else
                X = X & vbTab & HFApp.XmlMBAdd(d, 0, "DiscountDate", Format("" & rs("InvoiceDate"), "yyyy-mm-dd"))
            End If
            X = X & vbTab & HFApp.XmlMBAdd(n, 2, "InvoiceType", "1")
            X = X & vbTab & HFApp.XmlMBAdd(n, 1, "InvoiceStatus", "1")
            If Val("" & rs("ItemRetainage")) <> 0 Then
                X = X & vbTab & HFApp.XmlMBAdd(n, 12.2, "Retention", rs("ItemRetainage"))
            End If
            X = X & vbTab & HFApp.XmlMBAdd(n, 2, "AccountingPeriod", PostingPeriod)
            If Val("" & rs("InvoiceDiscount")) <> 0 Then
                X = X & vbTab & HFApp.XmlMBAdd(n, 12.2, "DiscountAvailable", "" & rs("InvoiceDiscount"))
            End If
            X = X & vbTab & HFApp.XmlMBAdd(n, 12.2, "SetToPay", 0)
            X = X & vbTab & HFApp.XmlMBAdd(c, 20, "ShippingNumber", "" & rs("Invoice"))
            X = X & vbTab & HFApp.XmlMBAdd(n, 4, "PostingYear", PostingYear)

            
                If "" & rs("ItemTaxGroup") <> "" Then
                'Create Sales Tax Line String to be added as the last APInvoiceLine or it fails
                    
                    TaxLineString = TaxLineString & "<APInvoiceLineAdd>" & vbCrLf
                    TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(c, 50, "Desc", "Sales Tax")
                    
                    If Val("" & rs("InvoiceTax")) < 0 Then
                        TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 12.4, "Quantity", "-1")
                        TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 15.6, "UnitPrice", Abs(Val("" & rs("InvoiceTax"))))
                    Else
                        TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 12.4, "Quantity", "1")
                        TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 15.6, "UnitPrice", "" & rs("InvoiceTax"))
                    End If
                    TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(c, 10, "AccountRef", "" & rs("ItemDebitAccount"))
                    If HFApp.Options.ValueByName("UseJobAsSubAcct") = "True" And "" & rs("ItemDebitAccount") <> "" Then
                        TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(c, 10, "SubaccountRef", "" & rs("ItemJob"))
                    End If
                    TaxLineString = TaxLineString & "</APInvoiceLineAdd>" & vbCrLf
                    
                    
                    'Add Job Cost Lines
                    TaxLineString = TaxLineString & "<APInvoiceJobCostLineAdd>" & vbCrLf
                    TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 10, "JobRef", "" & rs("ItemJob"))
                 
                    TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(c, IIf(POType = "Subcontract", 50, 75), "Desc", "Sales Tax")
                    
                    If Val("" & HFApp.SqlExec("select case count(distinct jcextra) when 1 then max(jcextra) else '' end from poitems where ponumber=" & DbQuote(Str, "" & rs("ItemCommitment")), dbHomeFront)(0)) <> 0 Then
                        TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 10, "PhaseRef", Val("" & HFApp.SqlExec("select case count(distinct jcextra) when 1 then max(jcextra) else '' end from poitems where ponumber=" & DbQuote(Str, "" & rs("ItemCommitment")), dbHomeFront)(0)))
                    End If
                    TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 15.3, "CostCodeRef", "" & rs("ItemPhase"))
                    TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 2, "CostTypeRef", "" & rs("ItemCategory"))
                    TaxLineString = TaxLineString & vbTab & HFApp.XmlMBAdd(n, 12.2, "CostAmount", "" & rs("InvoiceTax"))
                    TaxLineString = TaxLineString & "</APInvoiceJobCostLineAdd>" & vbCrLf
                                
                    
                End If
            
            
        End If
        
    
        ' Add AP Lines
        X = X & "<APInvoiceLineAdd>" & vbCrLf
        If "" & rs("ItemDescription") = "" Then
            X = X & vbTab & HFApp.XmlMBAdd(c, 50, "Desc", "Invoice for Job " & rs("ItemJob"))
        Else
            X = X & vbTab & HFApp.XmlMBAdd(c, 50, "Desc", "" & rs("ItemDescription"))
        End If
        
        
        ' ensure pretax is not zero
        Qty = Val("" & rs("ItemQty"))
        Rate = Val("" & rs("ItemRate"))
        Amount = Val("" & rs("ItemPretax"))
        If Not PostPOQtyToAccounting Then
            Qty = 1
            Rate = Amount
        End If
        If Amount < 0 Then
            Qty = -1 * Abs(Qty)
        Else
            Qty = Abs(Qty)
        End If
        Rate = Abs(Rate)
    
        
        X = X & vbTab & HFApp.XmlMBAdd(n, 12.4, "Quantity", Qty)
        X = X & vbTab & HFApp.XmlMBAdd(n, 15.6, "UnitPrice", Rate)
        
        
        X = X & vbTab & HFApp.XmlMBAdd(c, 10, "AccountRef", "" & rs("ItemDebitAccount"))
        If HFApp.Options.ValueByName("UseJobAsSubAcct") = "True" And "" & rs("ItemDebitAccount") <> "" Then
            X = X & vbTab & HFApp.XmlMBAdd(c, 10, "SubaccountRef", "" & rs("ItemJob"))
        End If
        If POType = "Subcontract" And subID <> 0 Then
            SubRefNumber = HFApp.SqlExec("Select linref from sbcnln where recnum = " & DbQuote(num, subID) & " and linnum = " & DbQuote(num, "" & rs("ItemCommitmentItem")), dbAccounting)(0)
            X = X & vbTab & HFApp.XmlMBAdd(c, 32, "SubcontractLineRefNumber", "" & SubRefNumber)
        End If
        X = X & "</APInvoiceLineAdd>" & vbCrLf

        'Add Job Cost Lines
        X = X & "<APInvoiceJobCostLineAdd>" & vbCrLf
        X = X & vbTab & HFApp.XmlMBAdd(n, 10, "JobRef", "" & rs("ItemJob"))
        If "" & rs("ItemDescription") = "" Then
            X = X & vbTab & HFApp.XmlMBAdd(c, IIf(POType = "Subcontract", 50, 75), "Desc", "Invoice for Job " & rs("ItemJob"))
        Else
            X = X & vbTab & HFApp.XmlMBAdd(c, IIf(POType = "Subcontract", 50, 75), "Desc", "" & rs("ItemDescription"))
        End If
        If Val("" & HFApp.SqlExec("select case count(distinct jcextra) when 1 then max(jcextra) else '' end from poitems where ponumber=" & DbQuote(Str, "" & rs("ItemCommitment")), dbHomeFront)(0)) <> 0 Then
            X = X & vbTab & HFApp.XmlMBAdd(n, 10, "PhaseRef", Val("" & HFApp.SqlExec("select case count(distinct jcextra) when 1 then max(jcextra) else '' end from poitems where ponumber=" & DbQuote(Str, "" & rs("ItemCommitment")), dbHomeFront)(0)))
        End If
        X = X & vbTab & HFApp.XmlMBAdd(n, 15.3, "CostCodeRef", "" & rs("ItemPhase"))
        X = X & vbTab & HFApp.XmlMBAdd(n, 2, "CostTypeRef", "" & rs("ItemCategory"))
        X = X & vbTab & HFApp.XmlMBAdd(n, 12.2, "CostAmount", Amount)
        X = X & "</APInvoiceJobCostLineAdd>" & vbCrLf
    
        
        rs.MoveNext
    Wend
    If lastID <> -999 Then
        If TaxLineString <> "" Then
            X = X & TaxLineString
        End If
        X = X & "</APInvoiceAddRq>" & vbCrLf
        X = X & HFApp.XmlMBEnd()
        On Error Resume Next
        s = HFApp.XmlMbSubmit(X, HFApp.Options(MasterBuilderPWD))
        If err.Number = 0 Then
            Call HFApp.SqlExec("update Invoices set Status='Posted', postingdate=getdate() where InvoiceID=" & DbQuote(num, lastID), dbHomeFront)
        Else
            MsgBox "Unable to post invoice " & lastInvoiceNumber & vbCrLf & vbCrLf & err.Description, vbCritical, App.ProductName
        End If
        On Error GoTo eh
    End If
    Screen.MousePointer = vbDefault
    
    
    Exit Sub
eh: Call ErrHandler(SRCFILE & "PostInvoices")
    Screen.MousePointer = vbDefault
End Sub

Private Sub ReadSettings()
    mAPInvoicesDefaultToBillable = IIf(HFApp.Options.ValueByName("APInvoicesDefaultToBillable") = "true", "true", "false")

    mHBInvoiceSuffix = HFApp.Options.ValueByName("HBInvoiceSuffix")
    mV99 = HFApp.Options.ValueByName("DefaultAPVariance")
    
    If mHBInvoiceSuffix = "" Then mHBInvoiceSuffix = "HB"
    If mV99 = "" Then mV99 = "V99"
    
End Sub
