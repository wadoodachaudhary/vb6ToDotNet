VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmuser 
   Caption         =   "HomeFront - User Update"
   ClientHeight    =   6990
   ClientLeft      =   7500
   ClientTop       =   3870
   ClientWidth     =   13620
   Icon            =   "frmuser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   13620
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   150
      TabIndex        =   17
      Top             =   120
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "General Information"
      TabPicture(0)   =   "frmuser.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CommonDialog1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Permissions / Privileges"
      TabPicture(1)   =   "frmuser.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).ControlCount=   1
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Height          =   5175
         Left            =   120
         TabIndex        =   141
         Top             =   360
         Width           =   10815
         Begin VB.CommandButton cmdDivisions 
            Caption         =   "Assign Divisions"
            Height          =   375
            Left            =   3240
            TabIndex        =   211
            Top             =   3360
            Width           =   3525
         End
         Begin VB.CommandButton cmdCopyUser 
            Caption         =   "Copy User Permissions"
            Height          =   735
            Left            =   1080
            TabIndex        =   185
            Top             =   3840
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Height          =   300
            Left            =   6840
            Picture         =   "frmuser.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   179
            Top             =   2475
            Width           =   255
         End
         Begin VB.CommandButton cmdAssignCommunity 
            Caption         =   "Assign Communities"
            Height          =   375
            Left            =   3240
            TabIndex        =   157
            Top             =   2880
            Width           =   3525
         End
         Begin VB.OptionButton Option6 
            Caption         =   "6 - Read Only Access"
            Height          =   255
            Left            =   7320
            TabIndex        =   150
            Top             =   2775
            Width           =   2535
         End
         Begin VB.OptionButton Option5 
            Caption         =   "5 - Design Center Sales Person"
            Height          =   255
            Left            =   7320
            TabIndex        =   149
            Top             =   2475
            Width           =   2535
         End
         Begin VB.OptionButton Option4 
            Caption         =   "4 - Sales Person"
            Height          =   255
            Left            =   7320
            TabIndex        =   148
            Top             =   2175
            Width           =   2055
         End
         Begin VB.OptionButton Option3 
            Caption         =   "3 - Accounting"
            Height          =   255
            Left            =   7320
            TabIndex        =   147
            Top             =   1875
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "2 -Sales Manager"
            Height          =   255
            Left            =   7320
            TabIndex        =   146
            Top             =   1575
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1 - Administrator"
            Height          =   255
            Left            =   7320
            TabIndex        =   145
            Top             =   1275
            Width           =   2055
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Set Default Permissions/Privileges for Selected HomeFront User Type"
            Height          =   615
            Left            =   7320
            TabIndex        =   144
            Top             =   3480
            Width           =   2685
         End
         Begin VB.OptionButton Option7 
            Caption         =   "7 - Estimating"
            Height          =   255
            Left            =   7320
            TabIndex        =   143
            Top             =   3120
            Width           =   2415
         End
         Begin VB.CommandButton cmdButtons 
            Height          =   315
            Index           =   9
            Left            =   6480
            Picture         =   "frmuser.frx":0236
            Style           =   1  'Graphical
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   960
            Width           =   285
         End
         Begin TDBText6Ctl.TDBText txtUserId 
            Height          =   285
            Left            =   3240
            TabIndex        =   0
            Top             =   975
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5644
            _ExtentY        =   503
            Caption         =   "frmuser.frx":07C0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmuser.frx":082C
            Key             =   "frmuser.frx":084A
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "A9"
            FormatMode      =   0
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   10
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TDBText6Ctl.TDBText txtUserName 
            DataField       =   "User_Name"
            DataSource      =   "curuser"
            Height          =   285
            Left            =   3240
            TabIndex        =   1
            Top             =   1275
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   503
            Caption         =   "frmuser.frx":08A4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmuser.frx":0910
            Key             =   "frmuser.frx":092E
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   0
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TDBText6Ctl.TDBText txtdesignation 
            DataField       =   "Designation"
            DataSource      =   "curuser"
            Height          =   285
            Left            =   3240
            TabIndex        =   2
            Top             =   1575
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   503
            Caption         =   "frmuser.frx":0988
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmuser.frx":09F4
            Key             =   "frmuser.frx":0A12
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   0
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TDBText6Ctl.TDBText txtOffice 
            DataField       =   "Office"
            DataSource      =   "curuser"
            Height          =   285
            Left            =   3240
            TabIndex        =   3
            Top             =   1875
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   503
            Caption         =   "frmuser.frx":0A6C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmuser.frx":0AD8
            Key             =   "frmuser.frx":0AF6
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   0
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TDBText6Ctl.TDBText txtusrpass 
            Height          =   285
            Left            =   3240
            TabIndex        =   4
            Top             =   2175
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   503
            Caption         =   "frmuser.frx":0B50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmuser.frx":0BBC
            Key             =   "frmuser.frx":0BDA
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   "*"
            AllowSpace      =   -1
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   0
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   3
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TDBText6Ctl.TDBText txtHomePage 
            DataField       =   "HomePage"
            DataSource      =   "curuser"
            Height          =   285
            Left            =   3240
            TabIndex        =   5
            Top             =   2475
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   503
            Caption         =   "frmuser.frx":0C34
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmuser.frx":0CA0
            Key             =   "frmuser.frx":0CBE
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   250
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin VB.Label lblLabels 
            Caption         =   "Home Page Report"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   178
            Top             =   2475
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "User Type:"
            Height          =   255
            Index           =   5
            Left            =   7320
            TabIndex        =   156
            Top             =   975
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Password"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   155
            Top             =   2175
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Office"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   154
            Top             =   1875
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Designation"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   153
            Top             =   1575
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "User Name"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   152
            Top             =   1275
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Login"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   151
            Top             =   975
            Width           =   1815
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5355
         Left            =   -74940
         TabIndex        =   18
         Top             =   450
         Width           =   13185
         _ExtentX        =   23257
         _ExtentY        =   9446
         _Version        =   393216
         Style           =   1
         Tabs            =   10
         Tab             =   7
         TabsPerRow      =   10
         TabHeight       =   520
         TabCaption(0)   =   "File Menu"
         TabPicture(0)   =   "frmuser.frx":0D18
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tasks Menu"
         TabPicture(1)   =   "frmuser.frx":0D34
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Setup Menu"
         TabPicture(2)   =   "frmuser.frx":0D50
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Inquiry Menu"
         TabPicture(3)   =   "frmuser.frx":0D6C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame2(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Reports Menu"
         TabPicture(4)   =   "frmuser.frx":0D88
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame2(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Tools Menu"
         TabPicture(5)   =   "frmuser.frx":0DA4
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame2(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Main Menu Bar"
         TabPicture(6)   =   "frmuser.frx":0DC0
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame2(6)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Customer Information"
         TabPicture(7)   =   "frmuser.frx":0DDC
         Tab(7).ControlEnabled=   -1  'True
         Tab(7).Control(0)=   "SSTab3"
         Tab(7).Control(0).Enabled=   0   'False
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Precision Builder"
         TabPicture(8)   =   "frmuser.frx":0DF8
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Frame5"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "Workticket"
         TabPicture(9)   =   "frmuser.frx":0E14
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Frame6"
         Tab(9).ControlCount=   1
         Begin VB.Frame Frame6 
            Height          =   4455
            Left            =   -74760
            TabIndex        =   214
            Top             =   690
            Width           =   10215
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Office Approve Tickets"
               DataField       =   "WTOfficeApprove"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   74
               Left            =   180
               TabIndex        =   218
               Top             =   1110
               Value           =   1  'Checked
               Width           =   2340
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer Approve Tickets"
               DataField       =   "WTCustomerApprove"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   73
               Left            =   180
               TabIndex        =   217
               Top             =   900
               Value           =   1  'Checked
               Width           =   2340
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Field Approve Tickets"
               DataField       =   "WTFieldApprove"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   72
               Left            =   180
               TabIndex        =   216
               Top             =   690
               Value           =   1  'Checked
               Width           =   2340
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Change grid layouts"
               DataField       =   "WTEditGridLayouts"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   71
               Left            =   180
               TabIndex        =   215
               Top             =   240
               Value           =   1  'Checked
               Width           =   2340
            End
         End
         Begin VB.Frame Frame5 
            Height          =   4455
            Left            =   -74760
            TabIndex        =   190
            Top             =   780
            Width           =   10215
            Begin VB.TextBox txtMaxPOAmount 
               DataField       =   "maxpoamount"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4105
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "curuser"
               Height          =   285
               Left            =   1845
               TabIndex        =   212
               Top             =   3840
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Limit Attachment Classifications to the List"
               DataField       =   "LimitAttachmentClasses"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   87
               Left            =   120
               TabIndex        =   210
               ToolTipText     =   "Only allow users to select classifications already in the list."
               Top             =   3360
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Access Options Menu"
               DataField       =   "PrecisionBldrOptions"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   86
               Left            =   5160
               TabIndex        =   209
               ToolTipText     =   "Allow users to access the Options menu item"
               Top             =   3360
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Generate Purchase Orders Only (no editing permitted)"
               DataField       =   "GeneratePOs"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   85
               Left            =   5160
               TabIndex        =   208
               ToolTipText     =   "Provide ability to generate PO's for existing items, but no editiing or adding is allowed."
               Top             =   3000
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Create/Edit Purchase Orders"
               DataField       =   "IssuePOs"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   84
               Left            =   5160
               TabIndex        =   207
               ToolTipText     =   "Full access to Purchase order generation."
               Top             =   2640
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Create/Edit Budgets"
               DataField       =   "IssueBudgets"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   83
               Left            =   5160
               TabIndex        =   206
               ToolTipText     =   "Full access to create/edit budget details."
               Top             =   2280
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Publish Worksheets?"
               DataField       =   "PublishWorksheets"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   52
               Left            =   120
               TabIndex        =   205
               Top             =   1320
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Access Precision Builder Estimating Worksheets?"
               DataField       =   "OpenEstimatingWorksheets"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   51
               Left            =   120
               TabIndex        =   204
               Top             =   960
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Access Precision Builder Marketing Worksheets?"
               DataField       =   "OpenMarketingWorksheets"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   50
               Left            =   120
               TabIndex        =   203
               Top             =   600
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Access to Precision Builder?"
               DataField       =   "OpenEstimating"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   49
               Left            =   120
               TabIndex        =   202
               Top             =   240
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Post Budgets to Accounting?"
               DataField       =   "PostBudgets"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   58
               Left            =   120
               TabIndex        =   201
               Top             =   2040
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Post POs to Accounting?"
               DataField       =   "PostCommitments"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   59
               Left            =   120
               TabIndex        =   200
               Top             =   2400
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Send POs?"
               DataField       =   "SendPO"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   60
               Left            =   120
               TabIndex        =   199
               Top             =   2760
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Edit Item Database?"
               DataField       =   "EditItemDb"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   61
               Left            =   120
               TabIndex        =   198
               Top             =   3090
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Edit Assemblies?"
               DataField       =   "EditAssemblies"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   62
               Left            =   5160
               TabIndex        =   197
               Top             =   240
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Setup Default Vendors?"
               DataField       =   "SetupDefaultVendor"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   63
               Left            =   5160
               TabIndex        =   196
               Top             =   600
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Setup PO Indexes?"
               DataField       =   "SetupPoIndex"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   64
               Left            =   5160
               TabIndex        =   195
               Top             =   960
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Syncrhonize from Timberline?"
               DataField       =   "PostToTimberline"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   53
               Left            =   120
               TabIndex        =   194
               Top             =   1680
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Setup Vendors?"
               DataField       =   "SetupVendors"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   65
               Left            =   5160
               TabIndex        =   193
               Top             =   1320
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Setup Vendor Pricing?"
               DataField       =   "SetupVendorPricing"
               DataSource      =   "curuser"
               Height          =   195
               Index           =   66
               Left            =   5160
               TabIndex        =   192
               Top             =   1680
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Create Job Quotes"
               DataField       =   "TaskEntQuote"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   16
               Left            =   5160
               TabIndex        =   191
               Top             =   1920
               Value           =   1  'Checked
               Width           =   4560
            End
            Begin VB.Label Label1 
               Caption         =   "PO Approval Max"
               Height          =   210
               Left            =   450
               TabIndex        =   213
               Top             =   3855
               Width           =   1440
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4695
            Index           =   0
            Left            =   -74880
            TabIndex        =   137
            Top             =   660
            Width           =   10575
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Open Database"
               DataField       =   "FileOpenData"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   0
               Left            =   3600
               TabIndex        =   140
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Database Utilities"
               DataField       =   "FileDataUtil"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   1
               Left            =   3600
               TabIndex        =   139
               Top             =   600
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 1"
               Height          =   375
               Index           =   2
               Left            =   3600
               TabIndex        =   138
               Top             =   960
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4575
            Index           =   1
            Left            =   -74880
            TabIndex        =   115
            Top             =   660
            Width           =   10575
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Allow User to Create Schedules?"
               DataField       =   "CreateSchedule"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   57
               Left            =   6960
               TabIndex        =   184
               Top             =   1440
               Value           =   1  'Checked
               Width           =   3360
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Allow User to Modify Actual End Dates Previously Saved?"
               DataField       =   "EditActualEndDate"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   56
               Left            =   6960
               TabIndex        =   183
               Top             =   1920
               Value           =   1  'Checked
               Width           =   3360
            End
            Begin VB.CheckBox Check0 
               Alignment       =   1  'Right Justify
               Caption         =   "Allow Users to View all Schedules"
               DataField       =   "ViewAllSchedules"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   55
               Left            =   6960
               TabIndex        =   182
               Top             =   2520
               Value           =   1  'Checked
               Width           =   3360
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Export to WMS XML files"
               DataField       =   "TaskWMSExport"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   82
               Left            =   7200
               TabIndex        =   176
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Enter Deposits"
               DataField       =   "TaskEntDep"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   3
               Left            =   360
               TabIndex        =   136
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Enter Mortgage"
               DataField       =   "TaskEntMort"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   4
               Left            =   360
               TabIndex        =   135
               Top             =   600
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Post Deposits"
               DataField       =   "TaskPostDep"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   5
               Left            =   360
               TabIndex        =   134
               Top             =   960
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Post Mortgage"
               DataField       =   "TaskPostMort"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   6
               Left            =   360
               TabIndex        =   133
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Process Lot Inventory"
               DataField       =   "TaskProcessLot"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   7
               Left            =   360
               TabIndex        =   132
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "View/Update Lot Inventory"
               DataField       =   "TaskUpdLot"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   8
               Left            =   360
               TabIndex        =   131
               Top             =   2040
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Post Lot Inventory"
               DataField       =   "TaskPostLot"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   9
               Left            =   360
               TabIndex        =   130
               Top             =   2400
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Process Sales Closing"
               DataField       =   "TaskProcessSales"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   10
               Left            =   360
               TabIndex        =   129
               Top             =   2760
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Process Commission"
               DataField       =   "TaskProcessComm"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   11
               Left            =   360
               TabIndex        =   128
               Top             =   3120
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "View/Update Unposted Commission"
               DataField       =   "TaskUpdComm"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   12
               Left            =   360
               TabIndex        =   127
               Top             =   3480
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Post Commission"
               DataField       =   "TaskPostComm"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   13
               Left            =   360
               TabIndex        =   126
               Top             =   3840
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "View/Update Unposted Adjustments"
               DataField       =   "TaskUpdAdj"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   14
               Left            =   3840
               TabIndex        =   125
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Post Adjustments"
               DataField       =   "TaskPostAdj"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   15
               Left            =   3840
               TabIndex        =   124
               Top             =   600
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Update Totals to Job Cost"
               DataField       =   "TaskUpdTotalJobCost"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   17
               Left            =   3840
               TabIndex        =   123
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Update Address Info to Job Cost"
               DataField       =   "TaskUpdAddrJobCost"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   18
               Left            =   3840
               TabIndex        =   122
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Update Date Info to Timberline"
               DataField       =   "TaskUpdDatesTL"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   19
               Left            =   3840
               TabIndex        =   121
               Top             =   2040
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Import Timberline Custom Fields"
               DataField       =   "TaskImportTLCustom"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   20
               Left            =   3840
               TabIndex        =   120
               Top             =   2400
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Synchronize Customer Date List"
               DataField       =   "TaskSyncDate"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   21
               Left            =   3840
               TabIndex        =   119
               Top             =   2760
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Synchronize Color Selection Items"
               DataField       =   "TaskSyncGroup"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   22
               Left            =   3840
               TabIndex        =   118
               Top             =   3120
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Import WebQuotes"
               DataField       =   "TaskImportWebQuotes"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   23
               Left            =   3840
               TabIndex        =   117
               Top             =   3480
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "WMS Synch Schedules"
               DataField       =   "TaskWMS_Synch"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   24
               Left            =   3840
               TabIndex        =   116
               Top             =   3840
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.Line Line2 
               X1              =   7080
               X2              =   10080
               Y1              =   1320
               Y2              =   1320
            End
            Begin VB.Label Label2 
               Caption         =   "Project Builder Settings"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7320
               TabIndex        =   181
               Top             =   960
               Width           =   2775
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4575
            Index           =   2
            Left            =   -74880
            TabIndex        =   92
            Top             =   660
            Width           =   10575
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Parking Stall"
               DataField       =   "ParkingStallMenu"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   81
               Left            =   7320
               TabIndex        =   172
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer Information"
               DataField       =   "SetupCustomer"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   25
               Left            =   240
               TabIndex        =   114
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Community"
               DataField       =   "SetupCommunity"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   26
               Left            =   240
               TabIndex        =   113
               Top             =   600
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Model"
               DataField       =   "SetupModel"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   27
               Left            =   240
               TabIndex        =   112
               Top             =   960
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Series"
               DataField       =   "SetupSeries"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   28
               Left            =   240
               TabIndex        =   111
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Major Groups"
               DataField       =   "SetupMajorGroup"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   29
               Left            =   240
               TabIndex        =   110
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Category"
               DataField       =   "SetupCategory"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   30
               Left            =   240
               TabIndex        =   109
               Top             =   2040
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Model Specific Option"
               DataField       =   "SetupModelOptions"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   31
               Left            =   240
               TabIndex        =   108
               Top             =   2400
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Global Option"
               DataField       =   "SetupGlobalOptions"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   32
               Left            =   240
               TabIndex        =   107
               Top             =   2760
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Design Center Area"
               DataField       =   "SetupDCArea"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   33
               Left            =   240
               TabIndex        =   106
               Top             =   3120
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Design Center Option"
               DataField       =   "SetupDCOption"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   34
               Left            =   240
               TabIndex        =   105
               Top             =   3480
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Jobs"
               DataField       =   "SetupJobs"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   35
               Left            =   240
               TabIndex        =   104
               Top             =   3840
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Construction Cut Off"
               DataField       =   "SetupConstStatus"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   36
               Left            =   3720
               TabIndex        =   103
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Project Managers"
               DataField       =   "SetupProjectManager"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   37
               Left            =   3720
               TabIndex        =   102
               Top             =   600
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Sales Persons"
               DataField       =   "SetupSalesperson"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   38
               Left            =   3720
               TabIndex        =   101
               Top             =   960
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer Service Person"
               DataField       =   "SetupCustSrvPerson"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   39
               Left            =   3720
               TabIndex        =   100
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Law Firms"
               DataField       =   "SetupLawFirms"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   40
               Left            =   3720
               TabIndex        =   99
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Lending Companies"
               DataField       =   "SetupLendingCompany"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   41
               Left            =   3720
               TabIndex        =   98
               Top             =   2040
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Developers"
               DataField       =   "SetupDevelopers"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   42
               Left            =   3720
               TabIndex        =   97
               Top             =   2400
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Lot Inventory"
               DataField       =   "SetupLotInventory"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   43
               Left            =   3720
               TabIndex        =   96
               Top             =   2760
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "TL Date List"
               DataField       =   "SetupTLDateList"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   44
               Left            =   3720
               TabIndex        =   95
               Top             =   3120
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Group Template/Color Sheet"
               DataField       =   "SetupGroupTemplate"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   45
               Left            =   3720
               TabIndex        =   94
               Top             =   3480
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Milestone"
               DataField       =   "MileStoneMenu"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   46
               Left            =   3720
               TabIndex        =   93
               Top             =   3840
               Value           =   1  'Checked
               Width           =   3000
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4575
            Index           =   3
            Left            =   -74880
            TabIndex        =   80
            Top             =   660
            Width           =   10575
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Model/Options"
               DataField       =   "InqModelOptions"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   47
               Left            =   3600
               TabIndex        =   91
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Global Options"
               DataField       =   "InqGlobalOptions"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   48
               Left            =   3600
               TabIndex        =   90
               Top             =   600
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Design Center Options"
               DataField       =   "InqDCOptions"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   49
               Left            =   3600
               TabIndex        =   89
               Top             =   960
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "New Deposits"
               DataField       =   "InqNewDeposits"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   50
               Left            =   3600
               TabIndex        =   88
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Posted Deposits"
               DataField       =   "InqPostDeposits"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   51
               Left            =   3600
               TabIndex        =   87
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "New Mortgage Advance"
               DataField       =   "InqNewMortgage"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   52
               Left            =   3600
               TabIndex        =   86
               Top             =   2040
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Posted Mortgage Advance"
               DataField       =   "InqPostMortgage"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   53
               Left            =   3600
               TabIndex        =   85
               Top             =   2400
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Locked Customer"
               DataField       =   "InqLockedCustomer"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   54
               Left            =   3600
               TabIndex        =   84
               Top             =   2760
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 1"
               Height          =   375
               Index           =   55
               Left            =   3600
               TabIndex        =   83
               Top             =   3120
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 2"
               Height          =   375
               Index           =   56
               Left            =   3600
               TabIndex        =   82
               Top             =   3480
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 3"
               Height          =   375
               Index           =   57
               Left            =   3600
               TabIndex        =   81
               Top             =   3840
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4575
            Index           =   4
            Left            =   -74880
            TabIndex        =   78
            Top             =   660
            Width           =   10575
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Report Manager"
               DataField       =   "ReportsRepMan"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   58
               Left            =   3600
               TabIndex        =   79
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4575
            Index           =   5
            Left            =   -74880
            TabIndex        =   60
            Top             =   660
            Width           =   10575
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Timberline Field Descriptions"
               DataField       =   "ToolsTLFieldDesc"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   69
               Left            =   2040
               TabIndex        =   77
               Top             =   3840
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Custom Descriptions"
               DataField       =   "ToolsCustDesc"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   68
               Left            =   2040
               TabIndex        =   76
               Top             =   3480
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "GST/PST Rebate Settings"
               DataField       =   "ToolsGSTPSTRebate"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   67
               Left            =   2040
               TabIndex        =   75
               Top             =   3120
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Field Validation"
               DataField       =   "ToolsFieldVal"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   66
               Left            =   2040
               TabIndex        =   74
               Top             =   2760
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Lot Status"
               DataField       =   "ToolsLotStatus"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   65
               Left            =   2040
               TabIndex        =   73
               Top             =   2400
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Commission Settings"
               DataField       =   "ToolsCommSettings"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   64
               Left            =   2040
               TabIndex        =   72
               Top             =   2040
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "System Settings"
               DataField       =   "ToolsSystemSetup"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   63
               Left            =   2040
               TabIndex        =   71
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer's Report Package"
               DataField       =   "ToolsCustRepPackage"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   62
               Left            =   2040
               TabIndex        =   70
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Administration and Maintenance"
               DataField       =   "ToolsAdminMaintain"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   61
               Left            =   2040
               TabIndex        =   69
               Top             =   960
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Customer Maintenance"
               DataField       =   "ToolsCustMaintain"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   60
               Left            =   2040
               TabIndex        =   68
               Top             =   600
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "User Administration"
               DataField       =   "ToolsUsrAdmin"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   59
               Left            =   2040
               TabIndex        =   67
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 2"
               Height          =   375
               Index           =   72
               Left            =   5520
               TabIndex        =   66
               Top             =   960
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 1"
               Height          =   375
               Index           =   71
               Left            =   5520
               TabIndex        =   65
               Top             =   600
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Work In Progress  Descriptions"
               DataField       =   "ToolsWorkInProgress"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   70
               Left            =   5520
               TabIndex        =   64
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 2"
               Height          =   375
               Index           =   73
               Left            =   5520
               TabIndex        =   63
               Top             =   1320
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 2"
               Height          =   375
               Index           =   74
               Left            =   5520
               TabIndex        =   62
               Top             =   1680
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Additional 2"
               Height          =   375
               Index           =   75
               Left            =   5520
               TabIndex        =   61
               Top             =   2040
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3000
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4575
            Index           =   6
            Left            =   -74880
            TabIndex        =   54
            Top             =   660
            Width           =   10575
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tasks Menu"
               DataField       =   "TasksMenu"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   76
               Left            =   3600
               TabIndex        =   59
               Top             =   240
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Setup Menu"
               DataField       =   "SetupMenu"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   77
               Left            =   3600
               TabIndex        =   58
               Top             =   600
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Reports Menu"
               DataField       =   "ReportsMenu"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   78
               Left            =   3600
               TabIndex        =   57
               Top             =   960
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Inquiry menu"
               DataField       =   "InquiryMenu"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   79
               Left            =   3600
               TabIndex        =   56
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3000
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tools Menu"
               DataField       =   "ToolsMenu"
               DataSource      =   "curuser"
               Height          =   375
               Index           =   80
               Left            =   3600
               TabIndex        =   55
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3000
            End
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   4695
            Left            =   120
            TabIndex        =   19
            Top             =   660
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   8281
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "General Permissions 1"
            TabPicture(0)   =   "frmuser.frx":0E30
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame2(7)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "General Permissions 2"
            TabPicture(1)   =   "frmuser.frx":0E4C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame4"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Customer Info Tabs"
            TabPicture(2)   =   "frmuser.frx":0E68
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame2(8)"
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame4 
               Height          =   4215
               Left            =   -74880
               TabIndex        =   170
               Top             =   360
               Width           =   10335
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Allow Edits to Quote Complete Custom Options?"
                  DataField       =   "EditCompletedCustom"
                  DataSource      =   "curuser"
                  Height          =   435
                  Index           =   70
                  Left            =   240
                  TabIndex        =   189
                  Top             =   2640
                  Value           =   1  'Checked
                  Width           =   4560
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Reverse Top Tab Order for Dates and Contract Details?"
                  DataField       =   "ReverseCustDateTab"
                  DataSource      =   "curuser"
                  Height          =   795
                  Index           =   69
                  Left            =   240
                  TabIndex        =   188
                  Top             =   2760
                  Value           =   1  'Checked
                  Width           =   4560
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Prompt for Change Order Approval?"
                  DataField       =   "PromptforCOApproval"
                  DataSource      =   "curuser"
                  Height          =   195
                  Index           =   68
                  Left            =   240
                  TabIndex        =   187
                  Top             =   2400
                  Value           =   1  'Checked
                  Width           =   4560
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Modify Custom Option Comments after Quote Complete?"
                  DataField       =   "ModifyCustom"
                  DataSource      =   "curuser"
                  Height          =   195
                  Index           =   67
                  Left            =   240
                  TabIndex        =   186
                  Top             =   2040
                  Value           =   1  'Checked
                  Width           =   4560
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Sales Managers Default Pick List View is Pending Customers?"
                  DataField       =   "PendingCustomerPick"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   54
                  Left            =   5040
                  TabIndex        =   180
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   4800
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Uncheck Purchased on Customers?"
                  DataField       =   "UnPurchase"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   48
                  Left            =   240
                  TabIndex        =   177
                  Top             =   1560
                  Value           =   1  'Checked
                  Width           =   4560
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Un Approve Change Order?"
                  DataField       =   "UnApproveCO"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   47
                  Left            =   240
                  TabIndex        =   175
                  Top             =   1200
                  Value           =   1  'Checked
                  Width           =   4560
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Construction Status?"
                  DataField       =   "UpdateConstStatus"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   46
                  Left            =   240
                  TabIndex        =   174
                  Top             =   840
                  Value           =   1  'Checked
                  Width           =   4560
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Cost Information?"
                  DataField       =   "UpdateCostInfo"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   45
                  Left            =   240
                  TabIndex        =   173
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   4560
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Access to All Menu Items when Customer form is open ?"
                  DataField       =   "AccessMenu"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   44
                  Left            =   240
                  TabIndex        =   171
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   4560
               End
            End
            Begin VB.Frame Frame2 
               Height          =   4215
               Index           =   8
               Left            =   -74760
               TabIndex        =   158
               Top             =   360
               Width           =   10335
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Buyer Info ?"
                  DataField       =   "ShowBuyerInfo"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   33
                  Left            =   240
                  TabIndex        =   169
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Home Info ?"
                  DataField       =   "ShowHomeInfo"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   34
                  Left            =   240
                  TabIndex        =   168
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Selections ?"
                  DataField       =   "ShowSelections"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   35
                  Left            =   240
                  TabIndex        =   167
                  Top             =   840
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Lot Premiums ?"
                  DataField       =   "ShowLotPremium"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   43
                  Left            =   240
                  TabIndex        =   166
                  Top             =   3720
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Deleted Items ?"
                  DataField       =   "ShowDeletedItems"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   42
                  Left            =   240
                  TabIndex        =   165
                  Top             =   3360
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Accounting Info ?"
                  DataField       =   "ShowAccountingInfo"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   41
                  Left            =   240
                  TabIndex        =   164
                  Top             =   3000
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Contract Summary ?"
                  DataField       =   "ShowContractSummary"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   40
                  Left            =   240
                  TabIndex        =   163
                  Top             =   2640
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Design Center ?"
                  DataField       =   "ShowDesignCenter"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   39
                  Left            =   240
                  TabIndex        =   162
                  Top             =   2280
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Change Order ?"
                  DataField       =   "ShowChangeOrder"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   38
                  Left            =   240
                  TabIndex        =   161
                  Top             =   1920
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Addendums ?"
                  DataField       =   "ShowAddendum"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   37
                  Left            =   240
                  TabIndex        =   160
                  Top             =   1560
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Dates ?"
                  DataField       =   "ShowDates"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   36
                  Left            =   240
                  TabIndex        =   159
                  Top             =   1200
                  Value           =   1  'Checked
                  Width           =   3000
               End
            End
            Begin VB.Frame Frame2 
               Height          =   4215
               Index           =   7
               Left            =   0
               TabIndex        =   20
               Top             =   360
               Width           =   10335
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Change Order ?"
                  DataField       =   "ShowCO"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   0
                  Left            =   240
                  TabIndex        =   53
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Show Design Center ?"
                  DataField       =   "ShowDC"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   1
                  Left            =   240
                  TabIndex        =   52
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Allow Deletion of Change Orders?"
                  DataField       =   "AllowDelCO"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   2
                  Left            =   240
                  TabIndex        =   51
                  Top             =   3000
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Change Selection After Purchased?"
                  DataField       =   "ChangeMiscPurchased"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   32
                  Left            =   7200
                  TabIndex        =   50
                  Top             =   3720
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Add CO to Locked Change Orders"
                  DataField       =   "AllowAddCO_LOCKCO"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   31
                  Left            =   7200
                  TabIndex        =   49
                  Top             =   3360
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Custom Pricing?"
                  DataField       =   "CustomPricing"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   30
                  Left            =   7200
                  TabIndex        =   48
                  Top             =   3000
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Sales Override?"
                  DataField       =   "SalesOverride"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   29
                  Left            =   7200
                  TabIndex        =   47
                  Top             =   2640
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Copy Customer?"
                  DataField       =   "CopyCustomer"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   28
                  Left            =   7200
                  TabIndex        =   46
                  Top             =   2280
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Change Dates"
                  DataField       =   "ChangeDate"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   27
                  Left            =   7200
                  TabIndex        =   45
                  Top             =   1920
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Customer"
                  DataField       =   "UPDCustomer"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   26
                  Left            =   7200
                  TabIndex        =   44
                  Top             =   1560
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Design Center"
                  DataField       =   "UPDDC"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   25
                  Left            =   7200
                  TabIndex        =   43
                  Top             =   1200
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Change Order"
                  DataField       =   "UPDCO"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   24
                  Left            =   7200
                  TabIndex        =   42
                  Top             =   840
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Addendum"
                  DataField       =   "UPDAddendum"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   23
                  Left            =   7200
                  TabIndex        =   41
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Customer Dates"
                  DataField       =   "UPDCustomerDates"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   22
                  Left            =   7200
                  TabIndex        =   40
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update TL Job"
                  DataField       =   "UPDTLJob"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   21
                  Left            =   3720
                  TabIndex        =   39
                  Top             =   3720
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Builder Approve?"
                  DataField       =   "BldrApproveCustomer"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   20
                  Left            =   3720
                  TabIndex        =   38
                  Top             =   3360
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Deleted Items?"
                  DataField       =   "UpdDeletedItems"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   19
                  Left            =   3720
                  TabIndex        =   37
                  Top             =   3000
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Purchase?"
                  DataField       =   "PurchaseCustomer"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   18
                  Left            =   3720
                  TabIndex        =   36
                  Top             =   2640
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Approve Customer on CO?"
                  DataField       =   "ApproveCustomer"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   17
                  Left            =   3720
                  TabIndex        =   35
                  Top             =   2280
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Create Jobs?"
                  DataField       =   "CreateJob"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   16
                  Left            =   3720
                  TabIndex        =   34
                  Top             =   1920
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Contract Summary?"
                  DataField       =   "UpdContractSummary"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   15
                  Left            =   3720
                  TabIndex        =   33
                  Top             =   1560
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "BarCode Entry?"
                  DataField       =   "BarcodeEntry"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   14
                  Left            =   3720
                  TabIndex        =   32
                  Top             =   1200
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Allow Addition of Design Center?"
                  DataField       =   "AddDC"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   13
                  Left            =   3720
                  TabIndex        =   31
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Allow Addition of Customer Dates?"
                  DataField       =   "AddCustomerDates"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   12
                  Left            =   3720
                  TabIndex        =   30
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Change Grid Layout?"
                  DataField       =   "ChangeGridLayout"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   11
                  Left            =   3720
                  TabIndex        =   29
                  Top             =   840
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Allow Addition of Change Orders?"
                  DataField       =   "AddCO"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   10
                  Left            =   240
                  TabIndex        =   28
                  Top             =   3720
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Allow Addition of Addendums?"
                  DataField       =   "AddAddendum"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   9
                  Left            =   240
                  TabIndex        =   27
                  Top             =   3360
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Create New Customer?"
                  DataField       =   "CreateNewCustomer"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   8
                  Left            =   240
                  TabIndex        =   26
                  Top             =   2640
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Delete Customer?"
                  DataField       =   "DeleteCustomer"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   7
                  Left            =   240
                  TabIndex        =   25
                  Top             =   2280
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Interior Color Info?"
                  DataField       =   "UPDCustomerGroups"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   6
                  Left            =   240
                  TabIndex        =   24
                  Top             =   1920
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update from TL?"
                  DataField       =   "UpdFromTL"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   5
                  Left            =   240
                  TabIndex        =   23
                  Top             =   1560
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update to TL?"
                  DataField       =   "UpdToTL"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   4
                  Left            =   240
                  TabIndex        =   22
                  Top             =   1200
                  Value           =   1  'Checked
                  Width           =   3000
               End
               Begin VB.CheckBox Check0 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Update Spec/Show Home?"
                  DataField       =   "UpdSpecShow"
                  DataSource      =   "curuser"
                  Height          =   375
                  Index           =   3
                  Left            =   240
                  TabIndex        =   21
                  Top             =   840
                  Value           =   1  'Checked
                  Width           =   3000
               End
            End
         End
      End
   End
   Begin MSAdodcLib.Adodc curuser 
      Height          =   375
      Left            =   5280
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   16
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "curuser"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc CurSearch 
      Height          =   330
      Left            =   2400
      Top             =   7320
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   1
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "CurSearch"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   11535
      Begin VB.CommandButton cmdClose 
         Height          =   375
         Left            =   7800
         Picture         =   "frmuser.frx":0E84
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Close"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   6960
         Picture         =   "frmuser.frx":140E
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Refresh"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         Picture         =   "frmuser.frx":1998
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Undo"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         Picture         =   "frmuser.frx":1CDA
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   375
         Left            =   5160
         Picture         =   "frmuser.frx":201C
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Delete User"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   4560
         Picture         =   "frmuser.frx":235E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Add New Record"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   3960
         Picture         =   "frmuser.frx":26A0
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Last Record"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   3360
         Picture         =   "frmuser.frx":2812
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Next Record"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   2760
         Picture         =   "frmuser.frx":2984
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Previous Record"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   2160
         Picture         =   "frmuser.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "First Record"
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myBookMark As Variant
Dim myEditFlag As Boolean
Dim myAddNewFlag As Boolean
Dim myDataChanged As Boolean
Dim vSave As Boolean
Dim realpass As String
Dim mySearchFlag As Boolean
Dim vrecnew As Boolean
Dim vStoreUser As String
Dim gSales_Person_ID As String


Public Function ShowForm()
    Me.Show
End Function

Private Sub Check0_Click(Index As Integer)
If Check0(Index).DataChanged = True Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub Check1_Click(Index As Integer)
If Check1(Index).DataChanged = True Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub cmdAdd_Click()
SetFields True
curuser.Recordset.AddNew
myAddNewFlag = True
Call ShowValue
txtusrpass.Text = ""
SetButtons False
SetFields True
txtUserId.SetFocus
Option4.value = True
Check0(0).value = 1
Check0(1).value = 1
Check0(2).value = 0
Me.cmdAssignCommunity.Visible = False
End Sub

Private Sub cmdAssignCommunity_Click()
'gcall = 2
'Me.WindowState = 1
gSales_Person_ID = Me.txtUserId.Text
frmSelectedList.ShowForm (2) ' vbModal
End Sub

Private Sub cmdbuttons_Click(Index As Integer)
gcall = 1
If FPickList.Choose(HFApp.Databases(dbHomefront), "Login Users", "select User_ID,User_Name from User_Manager" & IIf(gUser_UID <> "SYSTEM", " where User_ID<>'SYSTEM'", ""), txtUserId.Text, , , , "") Then
    gString = FPickList.SelectedItem("User_ID")
    Call FindUser
End If
End Sub
Private Sub cmdCancel_Click()
On Error Resume Next
curuser.Recordset.CancelUpdate
myAddNewFlag = False
If myEditFlag = False Then
    If myBookMark > 0 Then
      curuser.Recordset.Bookmark = myBookMark
    Else
      If Not curuser.Recordset.RecordCount = 0 Then curuser.Recordset.MoveFirst
    End If
    If Not IsNull(curuser.Recordset!User_ID) Then txtUserId.Text = curuser.Recordset!User_ID
    'txtusrpass.Text = ""
End If
Call ShowValue
SetButtons True
myEditFlag = False
myDataChanged = False
vSave = True
If Option4.value = True Then
   cmdAssignCommunity.Visible = True
Else
   cmdAssignCommunity.Visible = False
End If
End Sub
Private Sub cmdClose_Click()
Unload Me
Set frmuser = Nothing
End Sub
Private Sub cmdDelete_Click()
On Error GoTo herror
If MsgBox("Really want to Delete the selected user?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub

If gUser_UID <> txtUserId.Text Then
    With curuser.Recordset
      .Delete
      If Not .BOF Then
        .MovePrevious
      Else
        .MoveFirst
      End If
      If Not IsNull(!User_ID) Then txtUserId.Text = !User_ID
      txtusrpass.Text = ""
    End With
Else
    MsgBox "You cannot delete active user", vbOKOnly + vbExclamation, Me.Caption
End If

Exit Sub

herror:
If Err.Number <> 0 Then
   Call errHandler("cmdDelete")
   cmdCancel_Click
End If
End Sub
Private Sub cmdEdit_Click()
SetFields True
With curuser.Recordset
  If Not (.BOF Or .EOF) Then
    myBookMark = .Bookmark
  End If
End With
myEditFlag = True
SetButtons False
txtUserName.SetFocus
End Sub

Private Sub cmdDivisions_Click()
frmSelectedList.ShowForm (-2)
End Sub

Private Sub cmdFirst_Click()
On Error GoTo herror
'Call SaveValue
With curuser.Recordset
  .MoveFirst
  myDataChanged = False
  Call ShowValue
End With

Exit Sub

herror:
If Err.Number <> 0 Then
   Call errHandler("cmdFirst")
   cmdCancel_Click
End If
End Sub
Private Sub cmdLast_Click()
On Error GoTo herror
'Call SaveValue
With curuser.Recordset
  .MoveLast
  myDataChanged = False
  Call ShowValue
End With

Exit Sub

herror:
If Err.Number <> 0 Then
   Call errHandler("cmdLast")
   cmdCancel_Click
End If
End Sub
Private Sub cmdNext_Click()
On Error GoTo herror
'Call SaveValue
With curuser.Recordset
    If Not .EOF Then .MoveNext
    If .EOF And .RecordCount > 0 Then
       Beep
       .MoveLast
    End If
    myDataChanged = False
   Call ShowValue
End With

Exit Sub

herror:
If Err.Number <> 0 Then
   Call errHandler("cmdNext")
   cmdCancel_Click
End If
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo herror
'Call SaveValue
With curuser.Recordset
    If Not .BOF Then .MovePrevious
    If .BOF And .RecordCount > 0 Then
       Beep
       .MoveFirst
    End If
    mbDataChanged = False
   Call ShowValue
End With

Exit Sub

herror:
If Err.Number <> 0 Then
   Call errHandler("cmdPrevious")
   cmdCancel_Click
End If
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo herror
Dim bk As Variant
bk = curuser.Recordset.Bookmark
Call Form_Load
curuser.Recordset.Bookmark = bk
If Not IsNull(curuser.Recordset!User_ID) Then txtUserId.Text = curuser.Recordset!User_ID
Call ShowValue
Exit Sub

herror:
If Err.Number <> 0 Then
   Call errHandler("cmdRefresh")
   Call cmdCancel_Click
End If
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo herror
Dim s As String
With curuser.Recordset
    If myAddNewFlag = True Then
        With curuser.Recordset
            If Not (.BOF Or .EOF) Then
              myBookMark = .Bookmark
            End If
        End With
       !User_ID = txtUserId.Text
       .Update
    Else
       txtUserName.Text = txtUserName.Text
    End If
    !user_password = MyHomeFrontDll.MyCrypt(txtusrpass.Text, "Fazlul")
    !User_Access = IIf(Option1.value = True, 1, IIf(Option2.value = True, 2, IIf(Option3.value = True, 3, IIf(Option4.value = True, 4, IIf(Option5.value = True, 5, IIf(Option6.value = True, 6, 7))))))

   If myAddNewFlag = True And (Option4 = True Or Option5 = True) Then
        If Option4 = True Then
            HFApp.SqlExec "Insert into tblsales_persons (sales_person_id,sales_person_name) values(" & DbQuote(str, txtUserId.Text) & "," & DbQuote(str, txtUserName.Text) & ")"
        ElseIf Option5 = True Then
            HFApp.SqlExec "Insert into tblsales_persons (sales_person_id,sales_person_name,DCSales) values(" & DbQuote(str, txtUserId.Text) & "," & DbQuote(str, txtUserName.Text) & ",1)"
        End If
   End If
   s = "Update user_manager set " & vbCrLf
   s = s & " user_password = " & DbQuote(str, MyHomeFrontDll.MyCrypt(txtusrpass.Text, "Fazlul")) & ", " & vbCrLf
   s = s & " user_access = " & DbQuote(Num, IIf(Option1.value = True, 1, IIf(Option2.value = True, 2, IIf(Option3.value = True, 3, IIf(Option4.value = True, 4, IIf(Option5.value = True, 5, IIf(Option6.value = True, 6, 7))))))) & ", " & vbCrLf
   s = s & " user_name = " & DbQuote(str, txtUserName.Text) & ", " & vbCrLf
   s = s & " Designation = " & DbQuote(str, txtdesignation.Text) & ", " & vbCrLf
   s = s & " office = " & DbQuote(str, txtOffice.Text) & ", " & vbCrLf
   s = s & " MaxPOAmount = " & DbQuote(Cur, txtMaxPOAmount.Text) & ", " & vbCrLf
   s = s & " homepage = " & DbQuote(str, txtHomePage.Text) & ", " & vbCrLf
   For i = 0 To Check0.Count - 1
        If Check0(i).DataField <> "" Then
            s = s & Check0(i).DataField & "=" & IIf(Check0(i).value = 2, 0, Check0(i).value) & ", " & vbCrLf
        End If
   Next
   For i = 0 To Check1.Count - 1
        If Check1(i).DataField <> "" Then
            s = s & Check1(i).DataField & "=" & IIf(Check1(i).value = 2, 0, Check1(i).value) & IIf(i < Check1.Count - 1, ", ", "") & vbCrLf
        End If
   Next
   s = s & " where user_id = " & DbQuote(str, txtUserId.Text)
   HFApp.SqlExec s
   If myAddNewFlag = True Then
     HFApp.SqlExec "update user_manager set openestimating = 1,estimator=1,openmarketingworksheets=1,openestimatingworksheets=1,publishworksheets=1,posttotimberline=1 where user_access=1 or user_access=7"
   End If
   curuser.Refresh
   curuser.Recordset.Find "User_Id=" & DbQuote(str, txtUserId.Text)
End With

myEditFlag = False
myAddNewFlag = False
SetButtons True
myDataChanged = False
vrecnew = False
gsrch = False
vSave = True
Call ShowValue
SetButtons True
SetFields True
If Option4.value = True Then
   cmdAssignCommunity.Visible = True
Else
   cmdAssignCommunity.Visible = False
End If

Exit Sub

herror:
If Err.Number <> 0 Then
If Err.Number = -2147217864 Then
    MsgBox "Problems occurred when updating the user. Please Input your data again, and re-save", vbCritical, App.ProductName
Else
   Call errHandler("cmdUpdate")
End If

   cmdCancel_Click
End If
End Sub



Private Sub Command1_Click()

If Option1.value = True Then
   Call SetAdminPermission
ElseIf Option2.value = True Then
   Call SetSalesManagerPermission
ElseIf Option3.value = True Then
   Call SetAccountingPermission
ElseIf Option4.value = True Then
   Call SetSalesPersonPermission
ElseIf Option5.value = True Then
   Call SetDCSalesPersonPermission
ElseIf Option6.value = True Then
   Call SetReadOnlyPermission
ElseIf Option7.value = True Then
   Call SetEstimatingPermission
End If
'curuser.Recordset.Update
'If myAddNewFlag = False And vrecnew = False Then
'   myEditFlag = True
'   SetButtons False
'End If
End Sub

Private Sub Command2_Click()
CommonDialog1.filter = "*.rpt"
CommonDialog1.DialogTitle = "Please select a users home page"
CommonDialog1.ShowOpen
If CommonDialog1.FileTitle <> "" Then
   txtHomePage.Text = CommonDialog1.FileName
End If
End Sub

Private Sub cmdCopyUser_Click()
Dim s As String
Dim c As Long
Dim CopyUser As String
Dim rs As ADODB.Recordset

If myAddNewFlag = True Then
 Call cmdUpdate_Click
End If
If FPickList.Choose(HFApp.Databases(dbHomefront), "Login Users", "select User_ID,User_Name from User_Manager" & IIf(gUser_UID <> "SYSTEM", " where User_ID<>'SYSTEM' and user_id <>" & DbQuote(str, txtUserId.Text), ""), txtUserId.Text, , , , "") Then
    CopyUser = FPickList.SelectedItem("User_ID")
    
    Set rs = HFApp.SqlExec("Select top 1 * from user_manager", dbHomefront)
    If Not rs Is Nothing Then
        s = ""
        s = s & "Update T1 Set " & vbCrLf
        For c = 0 To rs.fields.Count - 1
           If rs.fields(c).Type = adBoolean Then
                s = s & rs.fields(c).Name & " = c." & rs.fields(c).Name
                If c < rs.fields.Count - 1 Then
                    s = s & "," & vbCrLf
                End If
           End If
        Next
        s = s & "     from user_manager T1"
        s = s & "     join user_manager c on (c.user_id = " & DbQuote(str, CopyUser) & ")"
        s = s & "     where T1.user_id = " & DbQuote(str, txtUserId.Text)

        HFApp.SqlExec s
    End If
'    s = ""
'    s = s & "    Update T1 Set" & vbCrLf
'    s = s & "    ShowCO=c.ShowCO," & vbCrLf
'    s = s & "    ShowDC=c.ShowDC," & vbCrLf
'    s = s & "    ALLOWDELCO=c.ALLOWDELCO," & vbCrLf
'    s = s & "    FileOpenData=c.FileOpenData," & vbCrLf
'    s = s & "    FileDataUtil=c.FileDataUtil," & vbCrLf
'    s = s & "    TaskEntDep=c.TaskEntDep," & vbCrLf
'    s = s & "    TaskEntMort=c.TaskEntMort," & vbCrLf
'    s = s & "    TaskPostDep=c.TaskPostDep," & vbCrLf
'    s = s & "    TaskPostMort=c.TaskPostMort," & vbCrLf
'    s = s & "    TaskProcessLot=c.TaskProcessLot," & vbCrLf
'    s = s & "    TaskUpdLot=c.TaskUpdLot," & vbCrLf
'    s = s & "    TaskPostLot=c.TaskPostLot," & vbCrLf
'    s = s & "    TaskProcessSales=c.TaskProcessSales," & vbCrLf
'    s = s & "    TaskProcessComm=c.TaskProcessComm," & vbCrLf
'    s = s & "    TaskUpdComm=c.TaskUpdComm," & vbCrLf
'    s = s & "    TaskPostComm=c.TaskPostComm," & vbCrLf
'    s = s & "    TaskUpdAdj=c.TaskUpdAdj," & vbCrLf
'    s = s & "    TaskPostAdj=c.TaskPostAdj," & vbCrLf
'    s = s & "    TaskEntQuote=c.TaskEntQuote," & vbCrLf
'    s = s & "    TaskUpdTotalJobCost=c.TaskUpdTotalJobCost," & vbCrLf
'    s = s & "    TaskUpdAddrJobCost=c.TaskUpdAddrJobCost," & vbCrLf
'    s = s & "    TaskUpdDatesTL=c.TaskUpdDatesTL," & vbCrLf
'    s = s & "    TaskImportTLCustom=c.TaskImportTLCustom," & vbCrLf
'    s = s & "    TaskSyncDate=c.TaskSyncDate," & vbCrLf
'    s = s & "    TaskSyncGroup=c.TaskSyncGroup," & vbCrLf
'    s = s & "    SetupCustomer=c.SetupCustomer," & vbCrLf
'    s = s & "    SetupCommunity=c.SetupCommunity," & vbCrLf
'    s = s & "    SetupModel=c.SetupModel," & vbCrLf
'    s = s & "    SetupSeries=c.SetupSeries," & vbCrLf
'    s = s & "    SetupMajorGroup=c.SetupMajorGroup," & vbCrLf
'    s = s & "    SetupCategory=c.SetupCategory," & vbCrLf
'    s = s & "    SetupModelOptions=c.SetupModelOptions," & vbCrLf
'    s = s & "    SetupGlobalOptions=c.SetupGlobalOptions," & vbCrLf
'    s = s & "    SetupDCArea=c.SetupDCArea," & vbCrLf
'    s = s & "    SetupDCOption=c.SetupDCOption," & vbCrLf
'    s = s & "    SetupJobs=c.SetupJobs," & vbCrLf
'    s = s & "    SetupConstStatus=c.SetupConstStatus," & vbCrLf
'    s = s & "    SetupProjectManager=c.SetupProjectManager," & vbCrLf
'    s = s & "    SetupSalesPerson=c.SetupSalesPerson," & vbCrLf
'    s = s & "    SetupCustSrvPerson=c.SetupCustSrvPerson," & vbCrLf
'    s = s & "    SetupLawFirms=c.SetupLawFirms," & vbCrLf
'    s = s & "    SetupLendingCompany=c.SetupLendingCompany," & vbCrLf
'    s = s & "    SetupDevelopers=c.SetupDevelopers," & vbCrLf
'    s = s & "    SetupLotInventory=c.SetupLotInventory," & vbCrLf
'    s = s & "    SetupTLDateList=c.SetupTLDateList," & vbCrLf
'    s = s & "    SetupGroupTemplate=c.SetupGroupTemplate," & vbCrLf
'    s = s & "    ReportsRepMan=c.ReportsRepMan," & vbCrLf
'    s = s & "    InqModelOptions=c.InqModelOptions," & vbCrLf
'    s = s & "    InqGlobalOptions=c.InqGlobalOptions," & vbCrLf
'    s = s & "    InqDCOptions=c.InqDCOptions," & vbCrLf
'    s = s & "    InqNewDeposits=c.InqNewDeposits," & vbCrLf
'    s = s & "    InqPostDeposits=c.InqPostDeposits," & vbCrLf
'    s = s & "    InqNewMortgage=c.InqNewMortgage," & vbCrLf
'    s = s & "    InqPostMortgage=c.InqPostMortgage," & vbCrLf
'    s = s & "    InqLockedCustomer=c.InqLockedCustomer," & vbCrLf
'    s = s & "    ToolsUsrAdmin=c.ToolsUsrAdmin," & vbCrLf
'    s = s & "    ToolsCustMaintain=c.ToolsCustMaintain," & vbCrLf
'    s = s & "    ToolsAdminMaintain=c.ToolsAdminMaintain," & vbCrLf
'    s = s & "    ToolsCustRepPackage=c.ToolsCustRepPackage," & vbCrLf
'    s = s & "    ToolsSystemSetup=c.ToolsSystemSetup," & vbCrLf
'    s = s & "    ToolsCommSettings=c.ToolsCommSettings," & vbCrLf
'    s = s & "    ToolsLotStatus=c.ToolsLotStatus," & vbCrLf
'    s = s & "    ToolsFieldVal=c.ToolsFieldVal," & vbCrLf
'    s = s & "    ToolsGSTPSTRebate=c.ToolsGSTPSTRebate," & vbCrLf
'    s = s & "    ToolsCustDesc=c.ToolsCustDesc," & vbCrLf
'    s = s & "    ToolsTLFieldDesc=c.ToolsTLFieldDesc," & vbCrLf
'    s = s & "    ToolsWorkInProgress=c.ToolsWorkInProgress," & vbCrLf
'    s = s & "    TasksMenu=c.TasksMenu," & vbCrLf
'    s = s & "    SetupMenu=c.SetupMenu," & vbCrLf
'    s = s & "    ReportsMenu=c.ReportsMenu," & vbCrLf
'    s = s & "    InquiryMenu=c.InquiryMenu," & vbCrLf
'    s = s & "    ToolsMenu=c.ToolsMenu," & vbCrLf
'    s = s & "    UPDSpecShow=c.UPDSpecShow," & vbCrLf
'    s = s & "    UPDToTL=c.UPDToTL," & vbCrLf
'    s = s & "    UPDFromTL=c.UPDFromTL," & vbCrLf
'    s = s & "    UPDCustomerGroups=c.UPDCustomerGroups," & vbCrLf
'    s = s & "    UPDContractSummary=c.UPDContractSummary," & vbCrLf
'    s = s & "    UPDDeletedItems=c.UPDDeletedItems," & vbCrLf
'    s = s & "    UPDTLJob=c.UPDTLJob," & vbCrLf
'    s = s & "    UPDCustomerDates=c.UPDCustomerDates," & vbCrLf
'    s = s & "    UPDAddendum=c.UPDAddendum," & vbCrLf
'    s = s & "    UPDCO=c.UPDCO," & vbCrLf
'    s = s & "    UPDDC=c.UPDDC," & vbCrLf
'    s = s & "    UPDCustomer=c.UPDCustomer," & vbCrLf
'    s = s & "    AddCustomerDates=c.AddCustomerDates," & vbCrLf
'    s = s & "    AddAddendum=c.AddAddendum," & vbCrLf
'    s = s & "    AddCO=c.AddCO," & vbCrLf
'    s = s & "    AddDC=c.AddDC," & vbCrLf
'    s = s & "    DeleteCustomer=c.DeleteCustomer," & vbCrLf
'    s = s & "    CreateNewCustomer=c.CreateNewCustomer," & vbCrLf
'    s = s & "    ChangeGridLayout=c.ChangeGridLayout," & vbCrLf
'    s = s & "    CreateJob=c.CreateJob," & vbCrLf
'    s = s & "    ChangeDate=c.ChangeDate," & vbCrLf
'    s = s & "    ApproveCustomer=c.ApproveCustomer," & vbCrLf
'    s = s & "    PurchaseCustomer=c.PurchaseCustomer," & vbCrLf
'    s = s & "    BldrApproveCustomer=c.BldrApproveCustomer," & vbCrLf
'    s = s & "    BarcodeEntry=c.BarcodeEntry," & vbCrLf
'    s = s & "    CopyCustomer=c.CopyCustomer," & vbCrLf
'    s = s & "    SalesOverride=c.SalesOverride," & vbCrLf
'    s = s & "    CustomPricing=c.CustomPricing," & vbCrLf
'    s = s & "    ShowBuyerInfo=c.ShowBuyerInfo," & vbCrLf
'    s = s & "    ShowHomeInfo=c.ShowHomeInfo," & vbCrLf
'    s = s & "    ShowDates=c.ShowDates," & vbCrLf
'    s = s & "    ShowSelections=c.ShowSelections," & vbCrLf
'    s = s & "    ShowAddendum=c.ShowAddendum," & vbCrLf
'    s = s & "    ShowChangeOrder=c.ShowChangeOrder," & vbCrLf
'    s = s & "    ShowDesignCenter=c.ShowDesignCenter," & vbCrLf
'    s = s & "    ShowContractSummary=c.ShowContractSummary," & vbCrLf
'    s = s & "    ShowAccountingInfo=c.ShowAccountingInfo," & vbCrLf
'    s = s & "    ShowDeletedItems=c.ShowDeletedItems," & vbCrLf
'    s = s & "    TaskImportWebQuotes=c.TaskImportWebQuotes," & vbCrLf
'    s = s & "    AllowAddCO=c.AllowAddCO," & vbCrLf
'    s = s & "    ChangeMiscPurchased=c.ChangeMiscPurchased," & vbCrLf
'    s = s & "    AllowAddCO_LOCKCO=c.AllowAddCO_LOCKCO," & vbCrLf
'    s = s & "    AccessMenu=c.AccessMenu," & vbCrLf
'    s = s & "    MileStoneMenu=c.MileStoneMenu," & vbCrLf
'    s = s & "    ParkingStallMenu=c.ParkingStallMenu," & vbCrLf
'    s = s & "    UpdateCostInfo=c.UpdateCostInfo," & vbCrLf
'    s = s & "    UpdateConstStatus=c.UpdateConstStatus," & vbCrLf
'    s = s & "    UnApproveCO=c.UnApproveCO," & vbCrLf
'    s = s & "    TaskWMS_Synch=c.TaskWMS_Synch," & vbCrLf
'    s = s & "    TaskWMSExport=c.TaskWMSExport," & vbCrLf
'    s = s & "    UnPurchase=c.UnPurchase," & vbCrLf
'    s = s & "    ShowLotPremium=c.ShowLotPremium," & vbCrLf
'    s = s & "    HomePage=c.HomePage," & vbCrLf
'    s = s & "    EditInventoryHomes=c.EditInventoryHomes," & vbCrLf
'    s = s & "    OpenEstimating=c.OpenEstimating," & vbCrLf
'    s = s & "    Estimator=c.Estimator," & vbCrLf
'    s = s & "    OpenMarketingWorksheets=c.OpenMarketingWorksheets," & vbCrLf
'    s = s & "    OpenEstimatingWorksheets=c.OpenEstimatingWorksheets," & vbCrLf
'    s = s & "    PublishWorksheets=c.PublishWorksheets," & vbCrLf
'    s = s & "    PostToTimberline=c.PostToTimberline," & vbCrLf
'    s = s & "    PendingCustomerPick=c.PendingCustomerPick," & vbCrLf
'    s = s & "    CreateSchedule=c.CreateSchedule," & vbCrLf
'    s = s & "    EditActualEndDate=c.EditActualEndDate," & vbCrLf
'    s = s & "    ViewAllSchedules=c.ViewAllSchedules," & vbCrLf
'    s = s & "    EditProjectManager=c.EditProjectManager," & vbCrLf
'    s = s & "    SetupDefaultVendor=c.SetupDefaultVendor," & vbCrLf
'    s = s & "    SetupVendorGroups=c.SetupVendorGroups," & vbCrLf
'    s = s & "    User_Access=c.User_Access," & vbCrLf
'    s = s & "    SetupVendors=c.SetupVendors" & vbCrLf
'    s = s & "     from user_manager T1"
'    s = s & "     join user_manager c on (c.user_id = 'CopyUser')"
'    s = s & "     where T1.user_id = 'UsertoUpdate'"
'    s = Replace(s, "UsertoUpdate", Me.txtUserId)
'    s = Replace(s, "CopyUser", CopyUser)
        
'    HFApp.SqlExec s
    Call cmdRefresh_Click
    
End If
Exit Sub
eh:
If Err.Number <> 0 Then
    Call MsgBox(Err.Description, vbCritical, App.ProductName)
    Resume Next
End If
End Sub

Private Sub curuser_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Not vrecnew Then
        With curuser.Recordset
            If Not (.BOF Or .EOF) Then
              myBookMark = .Bookmark
            End If
        End With
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
SSTab1.Tab = 0

'CenterForm Me, FMain

Me.Caption = App.ProductName & " - " & "User Management"
Call CreateConn

SetButtons True
vSave = True

If curuser.Recordset.RecordCount = 0 Then
   SetFields False
   vrecnew = True
Else
   SetFields True
   vrecnew = False
End If

If Not IsNull(curuser.Recordset!User_ID) Then txtUserId.Text = curuser.Recordset!User_ID
txtusrpass.Text = ""
myAddNewFlag = False
myEditFlag = False
mySearchFlag = False
Call ShowValue

If Option4.value = True Then
   cmdAssignCommunity.Visible = True
Else
   cmdAssignCommunity.Visible = False
End If

Exit Sub

herror:
If Err.Number <> 0 Then
   MsgBox Err.Description, vbOKOnly, Me.Caption
   
End If
End Sub
Public Sub FindUser()
Dim strseek As String
With curuser.Recordset
    If Not (.BOF Or .EOF) Then
      myBookMark = .Bookmark
    End If
    strseek = gString
    If strseek <> "" Then
        strseek = "User_Id=" & "'" & strseek & "'"
            If Not .BOF Then .MoveFirst
            .Find strseek
            If .EOF Then
               MsgBox "Unable to locate [" & strseek & "]", vbExclamation, "ORSS - User Seek"
                  If myBookMark > 0 Then
                    .Bookmark = myBookMark
                  Else
                    .MoveFirst
                  End If
            Else
                If Not IsNull(!User_ID) Then txtUserId.Text = !User_ID
                txtusrpass.Text = ""
            End If
    End If
End With
Call ShowValue
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdRefresh.Enabled = bVal
  cmdAdd.Enabled = bVal
  cmdUpdate.Enabled = Not bVal
  cmdCancel.Enabled = Not bVal
  cmdDelete.Enabled = bVal
  cmdClose.Enabled = bVal
  'cmdFind.Enabled = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub SetFields(bVal As Boolean)
  txtUserId.Enabled = bVal
  txtUserName.Enabled = bVal
  txtdesignation.Enabled = bVal
  txtOffice.Enabled = bVal
  txtusrpass.Enabled = bVal
  'txtAccess.Enabled = bVal
End Sub

Private Sub txtAccess_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReadUserPermissions
End Sub

Private Sub Option1_Click()
If Option1.value = True Then
    myDataChanged = True
   cmdAssignCommunity.Visible = False
End If
End Sub

Private Sub Option1_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub Option2_Click()
If Option2.value = True Then
    myDataChanged = True
   cmdAssignCommunity.Visible = False
End If
End Sub

Private Sub Option2_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub Option3_Click()
If Option3.value = True Then
    myDataChanged = True
   cmdAssignCommunity.Visible = False
End If
End Sub

Private Sub Option3_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub Option4_Click()
If Option4.value = True Then
    myDataChanged = True
   cmdAssignCommunity.Visible = True
End If
End Sub

Private Sub Option4_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub Option5_Click()
If Option5.value = True Then
    myDataChanged = True
   cmdAssignCommunity.Visible = False
End If
End Sub

Private Sub Option5_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub Option6_Click()
If Option6.value = True Then
    myDataChanged = True
   cmdAssignCommunity.Visible = False
End If
End Sub

Private Sub Option7_Click()
If Option7.value = True Then
    myDataChanged = True
   cmdAssignCommunity.Visible = False
End If
End Sub

Private Sub txtdesignation_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub txtHomePage_Change()
cmdUpdate.Enabled = True
End Sub

Private Sub txtMaxPOAmount_Change()
If txtMaxPOAmount.DataChanged = True Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub txtOffice_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub


Private Sub txtUserId_Change()
If myAddNewFlag = False And myEditFlag = False Then mySearchFlag = True
End Sub

Private Sub txtuserid_KeyDown(KeyCode As Integer, Shift As Integer)
If Not vrecnew Then
    If KeyCode = 13 Then
       If myAddNewFlag = False And mySearchFlag = True Then
          Call cmdFindTxt
       Else
          SendKeys "{TAB}"
       End If
    ElseIf KeyCode = vbKeyF4 Then
       gcall = 1
       frmUserList.Show vbModal
       Call FindUser
    End If
Else
    If KeyCode = 13 Then
       If myAddNewFlag = False Then
            gsrch = True
            cmdAdd_Click
            SendKeys "{TAB}"
        Else
            SendKeys "{TAB}"
        End If
    End If
End If
mySearchFlag = False
End Sub

Private Sub txtUserName_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Private Sub txtusrpass_GotFocus()
If myAddNewFlag = False And vrecnew = False Then
   myEditFlag = True
   SetButtons False
End If
End Sub

Public Sub Settings()
With curSettings.Recordset
     .AddNew
     !User_ID = txtUserId.Text
     !Date_Format = "mm/dd/yyyy"
     !Date_Mask = "##/##/####"
     !Job_List = 0
     !Equipment_List = 0
     !Equipment_Type_List = 0
     !Rate_Type_List = 0
     !Extra_List = 0
     !Cost_Code_List = 0
     !Agreement_List = 0
     !Location_List = 0
     !Category_List = 0
     !Customer_List = 0
     .Update
End With
End Sub

Public Sub CreatePass()
Dim lenpass As Integer
Dim a, b, c, d, E, f, g, H As String
lenpass = Len(txtusrpass.Text)
If myAddNewFlag = True Then
    If lenpass >= 4 Then
        a = Mid(txtusrpass.Text, 1, 1)
        b = Mid(txtusrpass.Text, 2, 1)
        c = Mid(txtusrpass.Text, 3, 1)
        d = Mid(txtusrpass.Text, 4, 1)
        realpass = d & "@" & a & "#!" & c & "%" & b & "$"
    Else
        MsgBox "Password must be of 4 Characters long", vbOKOnly + vbExclamation, Me.Caption
        vSave = False
    End If
Else
    If txtusrpass <> "" Then
        If lenpass >= 4 Then
            a = Mid(txtusrpass.Text, 1, 1)
            b = Mid(txtusrpass.Text, 2, 1)
            c = Mid(txtusrpass.Text, 3, 1)
            d = Mid(txtusrpass.Text, 4, 1)
            realpass = d & "@" & a & "#!" & c & "%" & b & "$"
        Else
            MsgBox "Password must be of 4 Characters long", vbOKOnly + vbExclamation, Me.Caption
            vSave = False
        End If
    End If
End If
End Sub




Public Sub SaveValue()
If txtusrpass.Text <> "" Then
    Call CreatePass
    curuser.Recordset!user_password = realpass
    curuser.Recordset.Update
End If
txtusrpass.Text = ""
End Sub

Public Sub cmdFindTxt()
Dim strseek As String
CurSearch.Refresh
With CurSearch.Recordset
    If Not (curuser.Recordset.BOF And curuser.Recordset.EOF) Then
      myBookMark = curuser.Recordset.Bookmark
    End If
    strseek = "'" & txtUserId.Text & "'"
    If strseek <> "" Then
            If Not .BOF Then .MoveFirst
            .Find "User_Id=" & strseek
            If .EOF Then
                Dim msg
                msg = "User Not Found. Do you really want to enter new User ?"
                If MsgBox(msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                    Cancel = True
                    If myBookMark > 0 Then
                      curuser.Recordset.Bookmark = myBookMark
                    Else
                      curuser.Recordset.MoveFirst
                    End If
                    myDataChanged = False
                Else
                    gsrch = True
                    Call cmdAdd_Click
                    SendKeys "{TAB}"
                End If
            Else
                curuser.Recordset.MoveFirst
                curuser.Recordset.Find "User_Id=" & strseek
            End If
            Call ShowValue
    End If
End With
End Sub
Public Sub ShowValue()
If myAddNewFlag = False Then
    With curuser.Recordset
         If Not IsNull(!User_ID) Then
            txtUserId.Text = !User_ID
         Else
            txtUserId.Text = ""
         End If
         If Not IsNull(!user_password) Then
            txtusrpass.Text = MyHomeFrontDll.MyDeCrypt(!user_password, "Fazlul")
         Else
            txtusrpass.Text = ""
         End If
         Select Case !User_Access
         Case 1
            Option1.value = True
         Case 2
            Option2.value = True
         Case 3
            Option3.value = True
         Case 4
            Option4.value = True
         Case 5
            Option5.value = True
         Case 6
            Option6.value = True
         Case 7
            Option7.value = True
         End Select
    End With
Else
    If gsrch = False Then txtUserId.Text = ""
End If
End Sub
Public Sub CreateConn()

'If gCentralizedSecurityReports = True Then
'    curuser.ConnectionString = gPrimaryDBConnString
'Else
curuser.ConnectionString = HFApp.ConnectionString(dbHomefront)
'End If
curuser.CommandType = adCmdText
curuser.RecordSource = "select * from User_Manager" & IIf(gUser_UID <> "SYSTEM", " where User_ID<>'SYSTEM'", "")
curuser.Refresh

'If gCentralizedSecurityReports = True Then
'    CurSearch.ConnectionString = gPrimaryDBConnString
'Else
CurSearch.ConnectionString = HFApp.ConnectionString(dbHomefront)
'End If
CurSearch.CommandType = adCmdText
CurSearch.RecordSource = "select * from User_Manager" & IIf(gUser_UID <> "SYSTEM", " where User_ID<>'SYSTEM'", "")
CurSearch.Refresh

End Sub

Public Sub SetAdminPermission()
Dim i As Integer
For i = 0 To 57
    If i <> 54 Then
        Check0(i).value = 1
    End If
Next i
Check0(3).value = 0
For i = 0 To 81
    Check1(i).value = 1
Next i
End Sub

Public Sub SetSalesManagerPermission()
Dim i As Integer
For i = 0 To 43
    Check0(i).value = 1
Next i
Check0(4).value = 0
Check0(5).value = 0
Check0(6).value = 0
Check0(15).value = 0
Check0(21).value = 0
Check0(12).value = 0
Check0(7).value = 0
Check0(13).value = 0
Check0(16).value = 0
Check0(33).value = 1
Check1(0).value = 1
Check1(1).value = 0
For i = 3 To 22
    Check1(i).value = 0
Next i
Check1(16).value = 1
For i = 25 To 39
    Check1(i).value = 1
Next i
Check1(40).value = 0
Check1(41).value = 0
Check1(42).value = 1
Check1(43).value = 0
Check1(44).value = 0
Check1(45).value = 0
Check1(47).value = 1
Check1(48).value = 1
Check1(49).value = 1
Check1(45).value = 0
Check1(47).value = 1
Check1(48).value = 1
Check1(49).value = 1
Check1(54).value = 1
For i = 59 To 72
    Check1(i).value = 0
Next i
Check1(62).value = 1
For i = 76 To 80
    Check1(i).value = 1
Next i
End Sub

Public Sub SetAccountingPermission()
Dim i As Integer
For i = 0 To 43
    Check0(i).value = 1
Next i
Check0(3).value = 0
Check0(6).value = 0
Check0(11).value = 0
Check0(7).value = 0
Check0(17).value = 0

Check1(0).value = 1
Check1(1).value = 0
For i = 3 To 20
    Check1(i).value = 1
Next i
Check1(16).value = 1
Check1(21).value = 0
Check1(22).value = 0
Check1(25).value = 1
For i = 26 To 39
    Check1(i).value = 0
Next i
Check1(35).value = 1
Check1(40).value = 1
Check1(41).value = 1
Check1(42).value = 1
Check1(43).value = 1
Check1(44).value = 0
For i = 47 To 54
    Check1(i).value = 1
Next i
Check1(58).value = 0
For i = 59 To 72
    Check1(i).value = 0
Next i
Check1(64).value = 1
For i = 76 To 80
    Check1(i).value = 1
Next i
End Sub

Public Sub SetSalesPersonPermission()
Dim i As Integer
For i = 0 To 43
    Check0(i).value = 0
Next i
Check0(0).value = 1
Check0(1).value = 1
Check0(2).value = 1
Check0(19).value = 1
Check0(22).value = 1
Check0(23).value = 1
Check0(24).value = 1
Check0(26).value = 1
Check0(9).value = 1
Check0(10).value = 1
Check0(8).value = 1
Check0(13).value = 1
Check0(18).value = 1
Check0(27).value = 1
Check0(33).value = 1
Check0(34).value = 1
Check0(35).value = 1
Check0(36).value = 1
Check0(37).value = 1
Check0(38).value = 1
Check0(40).value = 1
Check0(41).value = 1
Check0(42).value = 1

Check1(0).value = 1
Check1(1).value = 0
For i = 3 To 22
    Check1(i).value = 0
Next i
Check1(25).value = 1
For i = 26 To 45
    Check1(i).value = 0
Next i
Check1(47).value = 1
Check1(48).value = 1
Check1(49).value = 1
Check1(50).value = 0
Check1(51).value = 0
Check1(52).value = 0
Check1(53).value = 0
Check1(54).value = 1
For i = 59 To 70
    Check1(i).value = 0
Next i
For i = 76 To 80
    Check1(i).value = 1
Next i
Check1(76).value = 0
Check1(80).value = 0
End Sub

Public Sub SetDCSalesPersonPermission()
Dim i As Integer
For i = 0 To 43
    Check0(i).value = 0
Next i
Check0(0).value = 1
Check0(1).value = 1
Check0(2).value = 1
Check0(25).value = 1
Check0(26).value = 1
Check0(11).value = 1
Check0(13).value = 1
Check0(33).value = 1
Check0(34).value = 1
Check0(35).value = 1
Check0(36).value = 1
Check0(37).value = 1
Check0(38).value = 1
Check0(39).value = 1
Check0(40).value = 1
Check0(41).value = 1
Check0(42).value = 1

Check1(0).value = 1
Check1(1).value = 0
For i = 3 To 22
    Check1(i).value = 0
Next i
Check1(25).value = 1
For i = 26 To 45
    Check1(i).value = 0
Next i
Check1(47).value = 1
Check1(48).value = 1
Check1(49).value = 1
Check1(50).value = 0
Check1(51).value = 0
Check1(52).value = 0
Check1(53).value = 0
Check1(54).value = 1
For i = 59 To 70
    Check1(i).value = 0
Next i
For i = 76 To 80
    Check1(i).value = 1
Next i
Check1(76).value = 0
Check1(80).value = 0
End Sub

Public Sub SetReadOnlyPermission()
Dim i As Integer
For i = 0 To 27
    Check0(i).value = 0
Next i
Check0(0).value = 1
Check0(1).value = 1

Check1(0).value = 1
Check1(1).value = 0
For i = 3 To 22
    Check1(i).value = 0
Next i
Check1(25).value = 1
For i = 26 To 45
    Check1(i).value = 0
Next i
Check1(47).value = 1
Check1(48).value = 1
Check1(49).value = 1
Check1(50).value = 0
Check1(51).value = 0
Check1(52).value = 0
Check1(53).value = 0
Check1(54).value = 1
For i = 59 To 70
    Check1(i).value = 0
Next i
For i = 76 To 80
    Check1(i).value = 1
Next i
Check1(76).value = 0
Check1(80).value = 0
End Sub

Public Sub SetEstimatingPermission()
Dim i As Integer
For i = 0 To 27
    Check0(i).value = 0
Next i
Check0(0).value = 1
Check0(1).value = 1
Check0(2).value = 1
Check0(3).value = 1
Check0(26).value = 1
Check0(13).value = 1
Check0(55).value = 1
Check0(56).value = 1
Check0(57).value = 1
Check1(0).value = 1
Check1(1).value = 0
For i = 3 To 22
    Check1(i).value = 0
Next i
Check1(16).value = 1
Check1(25).value = 1
For i = 26 To 45
    Check1(i).value = 0
Next i
Check1(36).value = 1
Check1(47).value = 1
Check1(48).value = 1
Check1(49).value = 1
Check1(50).value = 0
Check1(51).value = 0
Check1(52).value = 0
Check1(53).value = 0
Check1(54).value = 1
For i = 59 To 70
    Check1(i).value = 0
Next i
Check1(80).value = 0
Check1(76).value = 1
Check1(83).value = 1
Check1(84).value = 1
Check1(85).value = 1
Check1(86).value = 1
End Sub

