VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{B6C8B132-5973-4983-AD46-8F3F10B04531}#1.0#0"; "vbalCbEx6.ocx"
Begin VB.Form FOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application Options"
   ClientHeight    =   6315
   ClientLeft      =   5970
   ClientTop       =   3375
   ClientWidth     =   9180
   Icon            =   "FOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Page 
      BorderStyle     =   0  'None
      Caption         =   "Digital Documents"
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
      Height          =   5505
      Index           =   6
      Left            =   21930
      TabIndex        =   13
      Top             =   8970
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtDocumentRoot 
         Height          =   315
         Left            =   1965
         TabIndex        =   15
         Top             =   615
         Width           =   3315
      End
      Begin VB.CommandButton cmdDocumentRoot 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5310
         Picture         =   "FOptions.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Browse folders"
         Top             =   555
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Digital Documents "
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
         Index           =   7
         Left            =   480
         TabIndex        =   31
         Top             =   180
         Width           =   1620
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   15
         X1              =   480
         X2              =   6240
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   14
         X1              =   480
         X2              =   6240
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblDocument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Document Root"
         Height          =   195
         Index           =   3
         Left            =   780
         TabIndex        =   16
         Top             =   675
         Width           =   1125
      End
   End
   Begin VB.Frame Page 
      BorderStyle     =   0  'None
      Caption         =   "Integration Settings"
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
      Height          =   5505
      Index           =   1
      Left            =   3645
      TabIndex        =   8
      Top             =   555
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame frmTLOptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2670
         Left            =   15
         TabIndex        =   50
         Top             =   1230
         Width           =   6435
         Begin VB.CheckBox chkUseTimberlineAsPOSource 
            Caption         =   "Read PO's from Sage's database. "
            Height          =   255
            Left            =   1440
            TabIndex        =   53
            Top             =   720
            Width           =   4395
         End
         Begin VB.CheckBox chkUsePervasiveDSN 
            Caption         =   "Use a Pervasive DSN"
            Height          =   255
            Left            =   1440
            TabIndex        =   52
            Top             =   1845
            Width           =   4395
         End
         Begin VB.ComboBox cboPervasiveDSN 
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "FOptions.frx":0596
            Left            =   1695
            List            =   "FOptions.frx":0598
            TabIndex        =   51
            Text            =   "cboPervasiveDSN"
            Top             =   2115
            Width           =   3915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Sage 300 Integration "
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
            Left            =   480
            TabIndex        =   55
            Top             =   270
            Width           =   1860
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            Index           =   16
            X1              =   480
            X2              =   6240
            Y1              =   375
            Y2              =   375
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            Index           =   17
            X1              =   480
            X2              =   6240
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   480
            Picture         =   "FOptions.frx":059A
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label7 
            Caption         =   $"FOptions.frx":0E64
            ForeColor       =   &H000000C0&
            Height          =   885
            Left            =   1740
            TabIndex        =   54
            Top             =   990
            Width           =   4545
         End
      End
      Begin VB.ComboBox cboSchedulingUsage 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "FOptions.frx":0F0B
         Left            =   1995
         List            =   "FOptions.frx":0F1B
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   630
         Width           =   3915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Scheduling Integration "
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
         Index           =   1
         Left            =   480
         TabIndex        =   27
         Top             =   180
         Width           =   1995
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   480
         X2              =   6240
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   480
         X2              =   6240
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Usage"
         Height          =   195
         Index           =   0
         Left            =   1425
         TabIndex        =   25
         Top             =   690
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   480
         Picture         =   "FOptions.frx":0FA4
         Top             =   540
         Width           =   480
      End
   End
   Begin VB.Frame Page 
      BorderStyle     =   0  'None
      Caption         =   "Descriptions && Codes"
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
      Height          =   5505
      Index           =   3
      Left            =   13785
      TabIndex        =   10
      Top             =   1215
      Visible         =   0   'False
      Width           =   6615
      Begin VSFlex8Ctl.VSFlexGrid gFormats 
         Height          =   4140
         Left            =   480
         TabIndex        =   19
         Top             =   480
         Width           =   5775
         _cx             =   1987323850
         _cy             =   1987320966
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   12
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":186E
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descriptions && Codes "
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
         Index           =   6
         Left            =   480
         TabIndex        =   30
         Top             =   180
         Width           =   1875
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   13
         X1              =   480
         X2              =   6240
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   12
         X1              =   480
         X2              =   6240
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.Frame Page 
      BorderStyle     =   0  'None
      Caption         =   "Email Options"
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
      Height          =   5505
      Index           =   8
      Left            =   8625
      TabIndex        =   41
      Top             =   -135
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtEmailMessage 
         Height          =   2775
         Left            =   1290
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox txtEmailSubject 
         Height          =   315
         Left            =   1290
         TabIndex        =   43
         Top             =   1080
         Width           =   4935
      End
      Begin VB.CheckBox chkAutoEmail 
         Caption         =   "Always email held invoices."
         Height          =   315
         Left            =   1320
         TabIndex        =   42
         Top             =   660
         Width           =   3675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Email Options "
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
         Index           =   3
         Left            =   480
         TabIndex        =   45
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label lblDocument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Message Body"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   49
         Top             =   1500
         Width           =   1050
      End
      Begin VB.Label lblDocument 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Subject"
         Height          =   195
         Index           =   4
         Left            =   690
         TabIndex        =   48
         Top             =   1140
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   480
         Picture         =   "FOptions.frx":193C
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblDocument 
         AutoSize        =   -1  'True
         Caption         =   "Use these variables in the subject or message body."
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   1260
         TabIndex        =   47
         Top             =   4320
         Width           =   3675
      End
      Begin VB.Label lblDocument 
         Caption         =   "<%Job%> <%Vendor%> <%VendorName%> <%Invoice%> <%Reason%>"
         ForeColor       =   &H00000080&
         Height          =   855
         Index           =   7
         Left            =   1440
         TabIndex        =   46
         Top             =   4560
         Width           =   3195
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   6
         X1              =   480
         X2              =   6240
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   7
         X1              =   480
         X2              =   6240
         Y1              =   285
         Y2              =   285
      End
   End
   Begin VB.Frame Page 
      BorderStyle     =   0  'None
      Caption         =   "Options && Settings"
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
      Height          =   5505
      Index           =   2
      Left            =   750
      TabIndex        =   9
      Top             =   6270
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox lstOptions 
         Height          =   3930
         IntegralHeight  =   0   'False
         ItemData        =   "FOptions.frx":2206
         Left            =   480
         List            =   "FOptions.frx":225B
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   360
         Width           =   5775
      End
      Begin VB.ComboBox cboDateFormat 
         Height          =   315
         ItemData        =   "FOptions.frx":2730
         Left            =   1830
         List            =   "FOptions.frx":274C
         TabIndex        =   20
         Text            =   "cboDateFormat"
         Top             =   4380
         Width           =   1575
      End
      Begin vbalComboEx6.vbalCboEx cboTaxGroup 
         Height          =   330
         Left            =   1830
         TabIndex        =   22
         Top             =   4710
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
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
         AutoCompleteListItemsOnly=   -1  'True
         DoAutoComplete  =   -1  'True
      End
      Begin vbalComboEx6.vbalCboEx cboGLAccount 
         Height          =   330
         Left            =   1830
         TabIndex        =   39
         Top             =   5070
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   582
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
         AutoCompleteListItemsOnly=   -1  'True
         DoAutoComplete  =   -1  'True
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Default Debit Account"
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   40
         Top             =   5130
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Options && Settings "
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
         Index           =   0
         Left            =   480
         TabIndex        =   26
         Top             =   150
         Width           =   1635
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   480
         X2              =   6240
         Y1              =   255
         Y2              =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   480
         X2              =   6240
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Format"
         Height          =   315
         Index           =   4
         Left            =   840
         TabIndex        =   24
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Default Tax Group"
         Height          =   315
         Index           =   9
         Left            =   300
         TabIndex        =   23
         Top             =   4770
         Width           =   1455
      End
   End
   Begin VB.Frame Page 
      BorderStyle     =   0  'None
      Caption         =   "Back Charges"
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
      Height          =   5505
      Index           =   7
      Left            =   16170
      TabIndex        =   32
      Top             =   6510
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cboBCCheckListItem 
         Height          =   315
         ItemData        =   "FOptions.frx":27A8
         Left            =   2760
         List            =   "FOptions.frx":27B8
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   4020
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CheckBox chkLineBackcharge 
         Caption         =   "Allow line by line invoice back charges."
         Height          =   315
         Left            =   1410
         TabIndex        =   1
         Top             =   600
         Width           =   3195
      End
      Begin VB.TextBox txtBackChargeMarkup 
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Text            =   "10%"
         Top             =   1710
         Width           =   495
      End
      Begin VB.CheckBox chkBatchBackCharge 
         Caption         =   "Allow purchase order back charge processing."
         Height          =   315
         Left            =   1410
         TabIndex        =   2
         Top             =   1380
         Width           =   4005
      End
      Begin VB.ListBox lstBackChargeVendorTypes 
         Height          =   960
         IntegralHeight  =   0   'False
         ItemData        =   "FOptions.frx":27C8
         Left            =   2760
         List            =   "FOptions.frx":27D8
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   2100
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   $"FOptions.frx":280C
         Height          =   795
         Index           =   6
         Left            =   1620
         TabIndex        =   36
         Top             =   3360
         Visible         =   0   'False
         Width           =   3945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendor Types"
         Height          =   195
         Index           =   4
         Left            =   1710
         TabIndex        =   35
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Markup Rate"
         Height          =   195
         Index           =   2
         Left            =   1755
         TabIndex        =   34
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Back Charges"
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
         Index           =   2
         Left            =   480
         TabIndex        =   33
         Top             =   180
         Width           =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   5
         X1              =   480
         X2              =   6240
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   4
         X1              =   480
         X2              =   6240
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   480
         Picture         =   "FOptions.frx":289E
         Top             =   540
         Width           =   480
      End
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   5355
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   423
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   0
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   6420
      TabIndex        =   6
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7740
      TabIndex        =   7
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Frame Page 
      BorderStyle     =   0  'None
      Caption         =   "Other Job Values"
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
      Height          =   5505
      Index           =   5
      Left            =   8655
      TabIndex        =   12
      Top             =   3180
      Visible         =   0   'False
      Width           =   6615
      Begin VSFlex8Ctl.VSFlexGrid gJobValues 
         Height          =   4140
         Left            =   480
         TabIndex        =   17
         Top             =   480
         Width           =   5775
         _cx             =   1987323850
         _cy             =   1987320966
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":3168
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Other Job Values "
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
         Left            =   480
         TabIndex        =   28
         Top             =   180
         Width           =   1530
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   9
         X1              =   480
         X2              =   6240
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   8
         X1              =   480
         X2              =   6240
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.Frame Page 
      BorderStyle     =   0  'None
      Caption         =   "Invoice Codes"
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
      Height          =   5505
      Index           =   4
      Left            =   2400
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VSFlex8Ctl.VSFlexGrid gCodes 
         Height          =   4140
         Left            =   480
         TabIndex        =   18
         Top             =   480
         Width           =   5775
         _cx             =   1987323850
         _cy             =   1987320966
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":3216
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Invoice Codes "
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
         Index           =   5
         Left            =   480
         TabIndex        =   29
         Top             =   180
         Width           =   1290
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   11
         X1              =   480
         X2              =   6240
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   10
         X1              =   480
         X2              =   6240
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000003&
      Caption         =   "   Select a page"
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   2355
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      Index           =   1
      X1              =   2355
      X2              =   2355
      Y1              =   0
      Y2              =   5595
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      Index           =   0
      X1              =   -960
      X2              =   4.80000e5
      Y1              =   5595
      Y2              =   5595
   End
End
Attribute VB_Name = "FOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FOptions::"
Public Cancel As Boolean



Private Sub cboPervasiveDSN_Change()
    chkUsePervasiveDSN.Value = IIf(cboPervasiveDSN.Text = "", vbUnchecked, vbChecked)
End Sub

Private Sub cboPervasiveDSN_Click()
    chkUsePervasiveDSN.Value = IIf(cboPervasiveDSN.Text = "", vbUnchecked, vbChecked)
End Sub

Private Sub chkUsePervasiveDSN_Click()
    If chkUsePervasiveDSN.Value = vbUnchecked Then cboPervasiveDSN.Text = ""
End Sub

Private Sub chkUseTimberlineAsPOSource_Click()
    cboPervasiveDSN.Enabled = chkUseTimberlineAsPOSource.Value = vbChecked
    chkUsePervasiveDSN.Enabled = chkUseTimberlineAsPOSource.Value = vbChecked
End Sub

Private Sub Form_Load()
On Error GoTo eh
    Dim i As Long
    Dim rs As Recordset
    
    For i = Page.LBound To Page.UBound
        Page(i).Move Label6.left + Label6.Width + 225, 0
        Page(i).Visible = False
    Next
    
    
    With TreeView
        .Nodes.Add(, , "P1", "Integration Settings").Expanded = True
        
        
        .Nodes.Add(, , "P2", "Options & Settings").Expanded = True
        .Nodes.Add(, , "P3", "Descriptions & Codes").Expanded = True
            .Nodes.Add("P3", tvwChild, "P4", "Invoice Codes").Expanded = True
            .Nodes.Add("P3", tvwChild, "P5", "Other Job Values").Expanded = True
        .Nodes.Add(, , "P6", "Digital Documents").Expanded = True
        .Nodes.Add(, , "P7", "Back Charges").Expanded = True
        .Nodes.Add(, , "P8", "Email Options").Expanded = True
    End With
    
    
    TreeView.Nodes.Item("P1").selected = True
    Call TreeView_NodeClick(TreeView.Nodes.Item("P1"))
    
    Call IniGetForm(Me)
    
    cboTaxGroup.Clear
    Set rs = HFApp.SqlExec("SELECT taxgroup,description FROM taxgroups where DivisionID =" & HFApp.DivisionID, dbHomeFront)
    While Not rs.EOF
        cboTaxGroup.AddItem Trim("" & rs(0)) & " - " & Trim("" & rs(1))
        rs.MoveNext
    Wend
    
    cboGLAccount.Clear
    Set rs = HFApp.SqlExec("SELECT account,description FROM glaccounts where DivisionID =" & HFApp.DivisionID, dbHomeFront)
    While Not rs.EOF
        cboGLAccount.AddItem Trim("" & rs(0)) & " - " & Trim("" & rs(1))
        rs.MoveNext
    Wend
    
    Call LoadDSNs(cboPervasiveDSN, , "Pervasive")
        
    Call ReadData
Exit Sub
eh: Call ErrHandler(SRCFILE & "Form_Load")
End Sub

Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
    Dim i As Long
    For i = Page.LBound To Page.UBound
        Page(i).Visible = Replace(Page(i).Caption, "&", "") = Replace(Node.Text, "&", "")
    Next
End Sub

Private Sub cmdDocumentRoot_Click()
    Dim s As String
    txtDocumentRoot.SetFocus
    s = txtDocumentRoot.Text
    If VBChooseFolder(0, BIF_RETURNONLYFSDIRS, s, "Select Document Root Folder", Me.hwnd, s) Then
        txtDocumentRoot.Text = s
    End If
End Sub





Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub gCodes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col < 1 Or Row < 1
End Sub

Private Sub gFormats_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = 0
End Sub

Private Sub gJobValues_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col < 1 Or Row < 1
End Sub

Private Sub lstOptions_GotFocus()
    lstOptions.ListIndex = Max(0, lstOptions.ListIndex)
End Sub

Private Sub lstOptions_ItemCheck(Item As Integer)
    If Item = 3 Then lstOptions.selected(4) = Not lstOptions.selected(3)
    If Item = 4 Then lstOptions.selected(3) = Not lstOptions.selected(4)
    
    If Item = 12 And lstOptions.selected(12) Then lstOptions.selected(13) = False
    If Item = 13 And lstOptions.selected(13) Then lstOptions.selected(12) = False
End Sub

Private Sub lstOptions_LostFocus()
    lstOptions.ListIndex = -1
End Sub

Private Sub txtBackChargeMarkup_Validate(Cancel As Boolean)
    txtBackChargeMarkup.Text = Val(txtBackChargeMarkup.Text) & "%"
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'OK
            If WriteData Then
                Unload Me
            End If
            
        Case 1 'cancel
            Cancel = True
            Unload Me
    End Select
End Sub

Private Sub ReadData()
On Error GoTo eh
    Dim i As Long
    Dim rs As Long
    Dim s As String
    
    With App.Options
        
        chkLineBackcharge = vbUnchecked ' IIf(.Value(EnableLineBackCharge), vbChecked, vbUnchecked)
        chkLineBackcharge.Enabled = False
        
        chkBatchBackCharge = IIf(.Value(EnableBackChargePOs), vbChecked, vbUnchecked)
        txtBackChargeMarkup = Val("" & .Value(BackChargeMarkup)) & "%"
        lstBackChargeVendorTypes.selected(0) = InStr(1, .Value(BackChargeVendorTypes), "1") <> 0
        lstBackChargeVendorTypes.selected(1) = InStr(1, .Value(BackChargeVendorTypes), "2") <> 0
        lstBackChargeVendorTypes.selected(2) = InStr(1, .Value(BackChargeVendorTypes), "3") <> 0
        lstBackChargeVendorTypes.selected(3) = InStr(1, .Value(BackChargeVendorTypes), "4") <> 0
        cboBCCheckListItem.ListIndex = Val(.Value(BackChargeStatusItem)) - 1
        
        
        i = -1
        
        i = i + 1: lstOptions.selected(i) = .Value(DefaultDescription)
        i = i + 1: lstOptions.selected(i) = .Value(AllowCrossPayingPOs)
        i = i + 1: lstOptions.selected(i) = .Value(AutoVarianceRemaining)
        i = i + 1: lstOptions.selected(i) = .Value(CalcDatesFromInvoiceDate)
        i = i + 1: lstOptions.selected(i) = Not lstOptions.selected(i - 1) 'calc dates from received date
        
        i = i + 1: lstOptions.selected(i) = .Value(ReceivedDateRequired)
        i = i + 1: lstOptions.selected(i) = .Value(PaymentDateRequired)
        i = i + 1: lstOptions.selected(i) = .Value(VerifyBudgets)
        i = i + 1: lstOptions.selected(i) = .Value(DontSetPaymentDate)
        i = i + 1: lstOptions.selected(i) = .Value(ShowQtyAndUnitPrice)
        i = i + 1: lstOptions.selected(i) = .Value(AutoCalcHeaderTotals)
        i = i + 1: lstOptions.selected(i) = .Value(WarnIfAcctDateNotCurrent)
        i = i + 1: lstOptions.selected(i) = .Value(AcctDateDefaultsToRecDate)
        i = i + 1: lstOptions.selected(i) = .Value(AcctDateDefaultsToInvDate)
        
        i = i + 1: lstOptions.selected(i) = HFApp.Options.ValueByName("AcctDateDefaultsToLienVoucherApprovalDate") = "True"
        
        i = i + 1: lstOptions.selected(i) = .Value(InvDescDefaultsToJobAddr)
        i = i + 1: lstOptions.selected(i) = .Value(JobBasedPONumbers)
        i = i + 1: lstOptions.selected(i) = .Value(RecalcPOTaxAtCurrentRate)
        i = i + 1: lstOptions.selected(i) = .Value(AutomaticallyClosePO)
        i = i + 1: lstOptions.selected(i) = .Value(AllowJobDrOverride)
        i = i + 1: lstOptions.selected(i) = .Value(AutoPayApprovedPOs)
        i = i + 1: lstOptions.selected(i) = .Value(DepartmentIsRequired)
        i = i + 1: lstOptions.selected(i) = .Value(AutoApprovePOInvoices)
        i = i + 1: lstOptions.selected(i) = .Value(AutoReleaseHeldPOInvoices)
        i = i + 1: lstOptions.selected(i) = .Value(SaveIncompleteInvoices)
        
        i = i + 1: lstOptions.selected(i) = .Value(PayablesDisablePayPOsScreen)
        
        i = i + 1: lstOptions.selected(i) = HFApp.Options.ValueByName("APInvoicesDefaultToBillable") = "True"

        
        lstOptions.ListIndex = -1
        
        gFormats.TextMatrix(1, 1) = .Value(Caption_Category)
        gFormats.TextMatrix(2, 1) = .Value(Caption_Commitment)
        gFormats.TextMatrix(3, 1) = .Value(Caption_DebitAccount)
        gFormats.TextMatrix(4, 1) = .Value(Caption_Extra)
        gFormats.TextMatrix(5, 1) = .Value(Caption_Job)
        gFormats.TextMatrix(6, 1) = .Value(Caption_Phase)
        gFormats.TextMatrix(7, 1) = .Value(Caption_Retainage)
        gFormats.TextMatrix(8, 1) = .Value(Caption_RetainagePct)
        gFormats.TextMatrix(9, 1) = .Value(Caption_TaxGroup)
        gFormats.TextMatrix(10, 1) = .Value(Caption_Vendor)
        gFormats.TextMatrix(11, 1) = .Value(Caption_WrapInsurance)
        
        
        gCodes.TextMatrix(1, 1) = .Value(InvoiceCode1Usage)
        gCodes.TextMatrix(2, 1) = .Value(InvoiceCode2Usage)
        gCodes.TextMatrix(1, 2) = .Value(InvoiceCode1Validate)
        gCodes.TextMatrix(2, 2) = .Value(InvoiceCode2Validate)
        gCodes.TextMatrix(1, 3) = .Value(InvoiceCode1Label)
        gCodes.TextMatrix(2, 3) = .Value(InvoiceCode2Label)
        
        
        gJobValues.Cell(flexcpChecked, 1, 1) = IIf(.Value(ShowJobPM), flexChecked, flexUnchecked)
        gJobValues.Cell(flexcpChecked, 2, 1) = IIf(.Value(ShowJobTitle1), flexChecked, flexUnchecked)
        gJobValues.Cell(flexcpChecked, 3, 1) = IIf(.Value(ShowJobTitle2), flexChecked, flexUnchecked)
        gJobValues.Cell(flexcpChecked, 4, 1) = IIf(.Value(ShowJobTitle3), flexChecked, flexUnchecked)
        gJobValues.Cell(flexcpChecked, 5, 1) = IIf(.Value(ShowJobTitle4), flexChecked, flexUnchecked)
       
        gJobValues.TextMatrix(1, 2) = .Value(Caption_JobPM)
        gJobValues.TextMatrix(2, 2) = .Value(Caption_JobTitle1)
        gJobValues.TextMatrix(3, 2) = .Value(Caption_JobTitle2)
        gJobValues.TextMatrix(4, 2) = .Value(Caption_JobTitle3)
        gJobValues.TextMatrix(5, 2) = .Value(Caption_JobTitle4)
        
        
        
        
        cboDateFormat = .Value(TimberlineDateFormat)
        cboTaxGroup.Text = .Value(DefaultTaxGroup) & " - " & .Value(DefaultTaxGroupDesc)
        cboGLAccount.Text = .Value(DefaultGLAccount) & " - " & .Value(DefaultGLAccoundDesc)
        
            
        cboSchedulingUsage.ListIndex = Val("" & .Value(SchedulingUsage))
            
                    
        txtDocumentRoot.Text = .Value(DocumentRoot)
        
            
    
        frmTLOptions.Visible = HFApp.Options(AccountingSystem) = asTimberline
        chkUseTimberlineAsPOSource.Value = IIf(.Value(UseTimberlineAsPOSource) = True, vbChecked, vbUnchecked)
        Call chkUseTimberlineAsPOSource_Click
        s = "" & HFApp.Options.ValueByName("TLDSN")
        chkUsePervasiveDSN.Value = IIf(s <> "", vbChecked, vbUnchecked)
        cboPervasiveDSN.Text = s
    
    
        txtEmailSubject.Text = .Value(EmailSubject)
        txtEmailMessage.Text = .Value(EmailMessage)
        chkAutoEmail.Value = IIf(.Value(AutoEmail), vbChecked, vbUnchecked)
    
    
    End With
    
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "ReadData")
End Sub

Private Function WriteData() As Boolean
    Dim i As Integer
    With App.Options
    
    
        .Value(EnableLineBackCharge) = chkLineBackcharge = vbChecked
        .Value(EnableBackChargePOs) = chkBatchBackCharge = vbChecked
        .Value(BackChargeMarkup) = Val(txtBackChargeMarkup)
        .Value(BackChargeVendorTypes) = Mid(IIf(lstBackChargeVendorTypes.selected(0), ",1", "") & _
                                            IIf(lstBackChargeVendorTypes.selected(1), ",2", "") & _
                                            IIf(lstBackChargeVendorTypes.selected(2), ",3", "") & _
                                            IIf(lstBackChargeVendorTypes.selected(3), ",4", ""), 2)
        .Value(BackChargeStatusItem) = cboBCCheckListItem.ListIndex + 1
    
        i = -1
        i = i + 1: .Value(DefaultDescription) = lstOptions.selected(i)
        i = i + 1: .Value(AllowCrossPayingPOs) = lstOptions.selected(i)
        i = i + 1: .Value(AutoVarianceRemaining) = lstOptions.selected(i)
        i = i + 1: .Value(CalcDatesFromInvoiceDate) = lstOptions.selected(i)
        i = i + 1: 'this is correct. do not remove. 'calc dates from received date
        i = i + 1: .Value(ReceivedDateRequired) = lstOptions.selected(i)
        i = i + 1: .Value(PaymentDateRequired) = lstOptions.selected(i)
        i = i + 1: .Value(VerifyBudgets) = lstOptions.selected(i)
        i = i + 1: .Value(DontSetPaymentDate) = lstOptions.selected(i)
        i = i + 1: .Value(ShowQtyAndUnitPrice) = lstOptions.selected(i)
        i = i + 1: .Value(AutoCalcHeaderTotals) = lstOptions.selected(i)
        i = i + 1: .Value(WarnIfAcctDateNotCurrent) = lstOptions.selected(i)
        i = i + 1: .Value(AcctDateDefaultsToRecDate) = lstOptions.selected(i)
        i = i + 1: .Value(AcctDateDefaultsToInvDate) = lstOptions.selected(i)
        i = i + 1: HFApp.Options.ValueByName("AcctDateDefaultsToLienVoucherApprovalDate") = lstOptions.selected(i)
        i = i + 1: .Value(InvDescDefaultsToJobAddr) = lstOptions.selected(i)
        i = i + 1: .Value(JobBasedPONumbers) = lstOptions.selected(i)
        i = i + 1: .Value(RecalcPOTaxAtCurrentRate) = lstOptions.selected(i)
        i = i + 1: .Value(AutomaticallyClosePO) = lstOptions.selected(i)
        i = i + 1: .Value(AllowJobDrOverride) = lstOptions.selected(i)
        i = i + 1: .Value(AutoPayApprovedPOs) = lstOptions.selected(i)
        i = i + 1: .Value(DepartmentIsRequired) = lstOptions.selected(i)
        i = i + 1: .Value(AutoApprovePOInvoices) = lstOptions.selected(i)
        i = i + 1: .Value(AutoReleaseHeldPOInvoices) = lstOptions.selected(i)
        i = i + 1: .Value(SaveIncompleteInvoices) = lstOptions.selected(i)
        i = i + 1: .Value(PayablesDisablePayPOsScreen) = lstOptions.selected(i)
        i = i + 1: HFApp.Options.ValueByName("APInvoicesDefaultToBillable") = lstOptions.selected(i)
        
        .Value(InvoiceDateRequired) = True
        .Value(AccountingDateRequired) = True
        
        
        .Value(Caption_Category) = gFormats.TextMatrix(1, 1)
        .Value(Caption_Commitment) = gFormats.TextMatrix(2, 1)
        .Value(Caption_DebitAccount) = gFormats.TextMatrix(3, 1)
        .Value(Caption_Extra) = gFormats.TextMatrix(4, 1)
        .Value(Caption_Job) = gFormats.TextMatrix(5, 1)
        .Value(Caption_Phase) = gFormats.TextMatrix(6, 1)
        .Value(Caption_Retainage) = gFormats.TextMatrix(7, 1)
        .Value(Caption_RetainagePct) = gFormats.TextMatrix(8, 1)
        .Value(Caption_TaxGroup) = gFormats.TextMatrix(9, 1)
        .Value(Caption_Vendor) = gFormats.TextMatrix(10, 1)
        .Value(Caption_WrapInsurance) = gFormats.TextMatrix(11, 1)
        
        .Value(InvoiceCode1Usage) = gCodes.TextMatrix(1, 1)
        .Value(InvoiceCode2Usage) = gCodes.TextMatrix(2, 1)
        .Value(InvoiceCode1Validate) = gCodes.TextMatrix(1, 2)
        .Value(InvoiceCode2Validate) = gCodes.TextMatrix(2, 2)
        .Value(InvoiceCode1Label) = gCodes.TextMatrix(1, 3)
        .Value(InvoiceCode2Label) = gCodes.TextMatrix(2, 3)
        
        .Value(ShowJobPM) = gJobValues.Cell(flexcpChecked, 1, 1) = flexChecked
        .Value(ShowJobTitle1) = gJobValues.Cell(flexcpChecked, 2, 1) = flexChecked
        .Value(ShowJobTitle2) = gJobValues.Cell(flexcpChecked, 3, 1) = flexChecked
        .Value(ShowJobTitle3) = gJobValues.Cell(flexcpChecked, 4, 1) = flexChecked
        .Value(ShowJobTitle4) = gJobValues.Cell(flexcpChecked, 5, 1) = flexChecked
        
              
        
        
        .Value(Caption_JobPM) = gJobValues.TextMatrix(1, 2)
        .Value(Caption_JobTitle1) = gJobValues.TextMatrix(2, 2)
        .Value(Caption_JobTitle2) = gJobValues.TextMatrix(3, 2)
        .Value(Caption_JobTitle3) = gJobValues.TextMatrix(4, 2)
        .Value(Caption_JobTitle4) = gJobValues.TextMatrix(5, 2)
        
        
        .Value(TimberlineDateFormat) = cboDateFormat

        .Value(DefaultTaxGroup) = Parse(cboTaxGroup.Text, 1, " - ")
        .Value(DefaultTaxGroupDesc) = Parse(cboTaxGroup.Text, 2, " - ")
        On Error Resume Next
        App.Options.Value(DefaultTaxGroupRate) = Val("" & HFApp.SqlExec("select grouprate from taxgroups where DivisionID =" & HFApp.DivisionID & " and taxgroup=" & DbQuote(Str, App.Options(DefaultTaxGroup)), dbHomeFront)(0))
        On Error GoTo 0
        .Value(DefaultGLAccount) = Parse(cboGLAccount.Text, 1, " - ")
        .Value(DefaultGLAccoundDesc) = Parse(cboGLAccount.Text, 2, " - ")

        .Value(SchedulingUsage) = cboSchedulingUsage.ListIndex
        
        
        .Value(DocumentRoot) = txtDocumentRoot.Text
        
        .Value(UseTimberlineAsPOSource) = chkUseTimberlineAsPOSource.Value = vbChecked
        
        .Value(AutoEmail) = chkAutoEmail.Value = vbChecked
        .Value(EmailSubject) = txtEmailSubject.Text
        .Value(EmailMessage) = txtEmailMessage.Text
        
        
        
        HFApp.Options.ValueByName("TLDSN") = cboPervasiveDSN.Text
        .DBWrite
        
    End With
    
    WriteData = True
    
    

    

End Function

















