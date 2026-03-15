VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{302C5C1A-C2E2-4302-9AF9-BAC87EEFDECE}#1.0#0"; "Panels.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6C83CF2C-BE8D-4EE2-9B07-7FF5E27AB2FD}#1.0#0"; "zybCombo.ocx"
Begin VB.Form FUserPermissions 
   Caption         =   "User Permissions"
   ClientHeight    =   7770
   ClientLeft      =   2625
   ClientTop       =   1260
   ClientWidth     =   14370
   Icon            =   "FUserPermissions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15255
   ScaleWidth      =   28800
   Begin VB.Frame UserPage 
      Caption         =   "User"
      Height          =   7605
      Left            =   9240
      TabIndex        =   21
      Top             =   3510
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox txtUserFirstName 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4530
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   2475
      End
      Begin VB.CheckBox chkPMUser 
         Caption         =   "Project Manager"
         Height          =   255
         Left            =   4530
         TabIndex        =   8
         Top             =   3000
         Width           =   1725
      End
      Begin zybCombo.zybCombobox cboSecurityGroups 
         Height          =   240
         Left            =   4530
         TabIndex        =   7
         Top             =   2640
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   476
         Style           =   2
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
      End
      Begin VB.CheckBox chkCRMAccess 
         Caption         =   "Sales User"
         Height          =   255
         Left            =   4530
         TabIndex        =   9
         Top             =   3270
         Width           =   1215
      End
      Begin VB.TextBox txtUserPswd 
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4530
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1260
         Width           =   2475
      End
      Begin VB.TextBox txtUserEmail 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4530
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2250
         Width           =   3345
      End
      Begin VB.TextBox txtUserOffice 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4530
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1950
         Width           =   2475
      End
      Begin VB.TextBox txtUserDesignation 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4530
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1650
         Width           =   2475
      End
      Begin VB.TextBox txtUserLastName 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4530
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   2475
      End
      Begin VB.TextBox txtUserID 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4530
         TabIndex        =   0
         Text            =   "abcdEFabcdEF"
         Top             =   360
         Width           =   1605
      End
      Begin VSFlex8Ctl.VSFlexGrid gDivisions 
         Height          =   2445
         Left            =   810
         TabIndex        =   10
         Top             =   3930
         Width           =   3525
         _cx             =   1976834122
         _cy             =   1976832217
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":000C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
      Begin VSFlex8Ctl.VSFlexGrid gSalesCommunities 
         Height          =   3315
         Left            =   4920
         TabIndex        =   11
         Top             =   3870
         Width           =   4815
         _cx             =   1976836397
         _cy             =   1976833751
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
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
         FormatString    =   $"FUserPermissions.frx":00B3
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "First Name"
         Height          =   195
         Index           =   7
         Left            =   3705
         TabIndex        =   65
         Top             =   690
         Width           =   750
      End
      Begin VB.Label lblCRMAccess 
         AutoSize        =   -1  'True
         Caption         =   "2 of 4 seats have been assigned."
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5820
         TabIndex        =   57
         Top             =   3300
         Width           =   2355
      End
      Begin VB.Label lblSalesCommunities 
         AutoSize        =   -1  'True
         Caption         =   "Sales Communities"
         Height          =   195
         Left            =   4920
         TabIndex        =   53
         Top             =   3630
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Security Group"
         Height          =   195
         Index           =   8
         Left            =   3420
         TabIndex        =   52
         Top             =   2670
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Divisions"
         Height          =   195
         Index           =   6
         Left            =   810
         TabIndex        =   51
         Top             =   3690
         Width           =   630
      End
      Begin VB.Image Image1 
         Height          =   1995
         Left            =   810
         Picture         =   "FUserPermissions.frx":0133
         Top             =   660
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   5
         Left            =   3750
         TabIndex        =   29
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email Address"
         Height          =   195
         Index           =   4
         Left            =   3465
         TabIndex        =   28
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Office"
         Height          =   195
         Index           =   3
         Left            =   4035
         TabIndex        =   27
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Designation"
         Height          =   195
         Index           =   2
         Left            =   3615
         TabIndex        =   26
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Last Name"
         Height          =   195
         Index           =   1
         Left            =   3690
         TabIndex        =   25
         Top             =   990
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "User ID"
         Height          =   195
         Index           =   0
         Left            =   3915
         TabIndex        =   24
         Top             =   390
         Width           =   540
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Document Management"
      Height          =   4995
      Index           =   9
      Left            =   5325
      TabIndex        =   63
      Top             =   1665
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   9
         Left            =   270
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   510
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":485D
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Data Portal"
      Height          =   4995
      Index           =   10
      Left            =   6972
      TabIndex        =   59
      Top             =   3264
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   10
         Left            =   150
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   240
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":4A32
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Save"
      Height          =   396
      Index           =   1
      Left            =   16416
      TabIndex        =   58
      Top             =   1956
      Width           =   1104
   End
   Begin Panels.Slider VSlider 
      Height          =   7500
      Left            =   3990
      Top             =   60
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   13229
      Min             =   300
      Max             =   7000
   End
   Begin Panels.Slider HSlider 
      Height          =   60
      Left            =   120
      Top             =   3600
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   106
      Orientation     =   1
      Min             =   300
      Max             =   7000
   End
   Begin VB.Frame TabFrame 
      Caption         =   "CRM"
      Height          =   4995
      Index           =   8
      Left            =   6780
      TabIndex        =   37
      Top             =   2970
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   8
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   360
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":4C07
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Workticket"
      Height          =   4995
      Index           =   7
      Left            =   6450
      TabIndex        =   36
      Top             =   2610
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   7
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   270
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":4DDC
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Warranty"
      Height          =   4995
      Index           =   6
      Left            =   6240
      TabIndex        =   35
      Top             =   2370
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   6
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":4FB1
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Scheduling"
      Height          =   4995
      Index           =   5
      Left            =   5850
      TabIndex        =   34
      Top             =   1920
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   5
         Left            =   90
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   210
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":5186
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
   End
   Begin VSFlex8Ctl.VSFlexGrid gGroups 
      Height          =   3225
      Left            =   150
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   330
      Width           =   3825
      _cx             =   1976834651
      _cy             =   1976833593
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
      FormatString    =   $"FUserPermissions.frx":535B
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
   Begin VSFlex8Ctl.VSFlexGrid gUsers 
      Height          =   3585
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3990
      Width           =   3825
      _cx             =   1976834651
      _cy             =   1976834228
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
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FUserPermissions.frx":53C5
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
   Begin VB.Frame TabFrame 
      Caption         =   "Purchasing"
      Height          =   4995
      Index           =   3
      Left            =   4980
      TabIndex        =   32
      Top             =   1050
      Width           =   7005
      Begin VB.TextBox txtMaxPOAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2130
         TabIndex        =   46
         Top             =   330
         Width           =   1485
      End
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   3
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   780
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":5419
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Maximum PO Amount"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   45
         Top             =   390
         Width           =   1605
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Sales"
      Height          =   4995
      Index           =   2
      Left            =   4800
      TabIndex        =   31
      Top             =   780
      Width           =   7005
      Begin zybCombo.zybCombobox cboUserAccess 
         Height          =   240
         Left            =   1860
         TabIndex        =   50
         Top             =   540
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   423
         Style           =   2
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
      End
      Begin VB.TextBox txtHomePage 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1860
         TabIndex        =   48
         Text            =   "Text7"
         Top             =   285
         Width           =   4545
      End
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   2
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1020
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":55EE
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
      Begin VB.Image cmdBrowse 
         Height          =   240
         Left            =   6420
         Picture         =   "FUserPermissions.frx":57C3
         Top             =   285
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "User Type"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   49
         Top             =   570
         Width           =   1605
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Home page report"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   47
         Top             =   300
         Width           =   1605
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Admin"
      Height          =   4995
      Index           =   1
      Left            =   4440
      TabIndex        =   30
      Top             =   540
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   5655
         Index           =   1
         Left            =   150
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976837879
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":590D
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Payables"
      Height          =   4995
      Index           =   4
      Left            =   5340
      TabIndex        =   33
      Top             =   1395
      Width           =   7005
      Begin VB.TextBox txtMaxInvoicePOOveragePercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   215
         Left            =   5460
         TabIndex        =   15
         Top             =   570
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtMaxInvoicePOOverageAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   215
         Left            =   5445
         TabIndex        =   16
         Text            =   "999999999.99"
         Top             =   795
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtMaxInvoiceAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   215
         Left            =   3570
         TabIndex        =   12
         Text            =   "999999999.99"
         Top             =   230
         Width           =   1095
      End
      Begin VB.TextBox txtMaxInvoiceOverrideAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   215
         Left            =   1845
         TabIndex        =   14
         Text            =   "999999999.99"
         Top             =   795
         Width           =   1095
      End
      Begin VB.TextBox txtMaxInvoiceOverridePercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   215
         Left            =   1845
         TabIndex        =   13
         Top             =   570
         Width           =   405
      End
      Begin VSFlex8Ctl.VSFlexGrid gPermissions 
         Height          =   4815
         Index           =   4
         Left            =   90
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1110
         Width           =   8295
         _cx             =   1976842535
         _cy             =   1976836397
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUserPermissions.frx":5AE2
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   7
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   2
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "PO overpayment %"
         Height          =   195
         Index           =   7
         Left            =   3300
         TabIndex        =   62
         Top             =   585
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "To a maximum of"
         Height          =   195
         Index           =   6
         Left            =   3300
         TabIndex        =   61
         Top             =   810
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice approval limit"
         Height          =   195
         Index           =   5
         Left            =   1845
         TabIndex        =   56
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "to a maximum of"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   55
         Top             =   810
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice variance %"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   54
         Top             =   585
         Width           =   1695
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   7488
      Left            =   4056
      TabIndex        =   19
      Top             =   108
      Width           =   10152
      _ExtentX        =   17912
      _ExtentY        =   13203
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   10
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Admin"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sales"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Purchasing"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Payables"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Scheduling"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Warranty"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Workticket"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CRM"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Document Management"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Data Portal"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image cmdTools 
      Height          =   240
      Index           =   5
      Left            =   660
      Picture         =   "FUserPermissions.frx":5CB7
      ToolTipText     =   "Find"
      Top             =   3720
      Width           =   240
   End
   Begin VB.Image cmdTools 
      Height          =   240
      Index           =   4
      Left            =   390
      Picture         =   "FUserPermissions.frx":6241
      ToolTipText     =   "Create New"
      Top             =   3720
      Width           =   240
   End
   Begin VB.Image cmdTools 
      Height          =   240
      Index           =   3
      Left            =   150
      Picture         =   "FUserPermissions.frx":67CB
      ToolTipText     =   "Delete"
      Top             =   3720
      Width           =   240
   End
   Begin VB.Image cmdTools 
      Height          =   240
      Index           =   2
      Left            =   690
      Picture         =   "FUserPermissions.frx":6D55
      ToolTipText     =   "Save As"
      Top             =   60
      Width           =   240
   End
   Begin VB.Image cmdTools 
      Height          =   240
      Index           =   1
      Left            =   420
      Picture         =   "FUserPermissions.frx":72DF
      ToolTipText     =   "Create New"
      Top             =   60
      Width           =   240
   End
   Begin VB.Image cmdTools 
      Height          =   240
      Index           =   0
      Left            =   180
      Picture         =   "FUserPermissions.frx":7869
      ToolTipText     =   "Delete"
      Top             =   60
      Width           =   240
   End
   Begin VB.Label lblUsers 
      AutoSize        =   -1  'True
      Caption         =   "Members"
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
      Left            =   990
      TabIndex        =   23
      Top             =   3750
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Groups"
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
      Left            =   1020
      TabIndex        =   22
      Top             =   90
      Width           =   615
   End
End
Attribute VB_Name = "FUserPermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "::FUserPermissions"

Private Dirty As EDirty
Private Enum EDirty
    e_none = 0
    e_user
    e_perms
End Enum
Private mUserIndex As Long  ' when you load a user save the original row number here so you can save
Private mSecGroupID As Long ' when you load permissions save the secgroup here so you can save

Private mCRMSeats As Long

'these are frame NOT tab indexes. tabs get removed. frames dont
Const TADMIN = 1
Const TSALES = 2
Const TPURCHASING = 3
Const TPAYABLES = 4
Const TSCHEDULING = 5
Const TWARRANTY = 6
Const TWORKTICKET = 7
Const TCRM = 8
Const TDOCUMENTS = 9
Const TDATAPORTAL = 10




Private Sub cboSecurityGroups_Click()
    Dirty = e_user
End Sub

Private Sub cboUserAccess_Click()
    Dirty = e_perms
End Sub


Private Sub chkCRMAccess_Click()
    Dirty = e_user
End Sub

Private Sub chkPMUser_Click()
    Dirty = e_user
End Sub

Private Sub cmdBrowse_Click()
    Dim s As String
    
    s = txtHomePage.Text
    If VBGetOpenFileName(s, , , , , True, "Crystal Report Files (*.rpt)|*.rpt|All Files (*.*)|*.*", , , , "rpt", Me.hwnd) Then
        txtHomePage.Text = s
        Dirty = e_perms
    End If

End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Dirty = e_perms Then Call SaveGroup
    If Dirty = e_user Then Call SaveUser
End Sub

Private Sub cmdTools_Click(Index As Integer)
    If IsDirty Then Exit Sub
    Select Case Index
        Case 0: Call DeleteGroup
        Case 1: Call CreateGroup
        Case 2: Call CopyGroup
        Case 3: Call DeleteUser
        Case 4: Call CreateUser
        Case 5: Call FindUser
    End Select
End Sub

Private Sub DeleteGroup()
    Dim s As String
    
    If mSecGroupID = 0 Then Exit Sub
    
    If gUsers.Rows > 1 Then
        MsgBox "You must remove all users before you can delete this group.", vbInformation, "Homefront"
        Exit Sub
    End If
    
    
    If MsgBox("Are you sure you want to delete this group?", vbOKCancel + vbQuestion, "Homefront") = vbCancel Then Exit Sub
    
    s = "delete from securitygroups where secgroupid=" & DbQuote(Num, mSecGroupID)
    Call HFApp.SqlExec(s)
    
    Call LoadGroups
End Sub
Private Sub CreateGroup()
    Dim ID As Long
    Dim r As Long
    
    Call HFApp.SqlExec("insert into securitygroups(Description) values('(untitled)')")
    ID = HFApp.SqlIdentity("SecurityGroups", dbHomeFront)
    Call LoadGroups
    
    With gGroups
        r = .FindRow(ID, 0, .ColIndex("SecGroupID"), , True)
        If r = -1 Then Exit Sub
        .Row = r
        Call .ShowCell(r, .ColIndex("Description"))
        Call .EditCell
    End With
    
End Sub
Private Sub CopyGroup()
    Dim s As String
    Dim ID As Long
    Dim r As Long
    
    With gGroups
        If .Row < 0 Then Exit Sub
        ID = .TextMatrix(.Row, .ColIndex("SecGroupID"))
    End With
    
    s = ""
    s = s & "declare @column varchar(100)" & vbCrLf
    s = s & "declare @sql nvarchar(max)" & vbCrLf
    s = s & "set @sql=''" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "declare c cursor for" & vbCrLf
    s = s & "select column_name " & vbCrLf
    s = s & "from information_schema.columns " & vbCrLf
    s = s & "where table_name='SecurityGroups' " & vbCrLf
    s = s & "and column_name not in('SecGroupID','Description')" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "open c" & vbCrLf
    s = s & "fetch next from c into @column" & vbCrLf
    s = s & "while @@fetch_status = 0" & vbCrLf
    s = s & "begin" & vbCrLf
    s = s & "    set @sql = @sql + ',[' + @column + ']'" & vbCrLf
    s = s & "    fetch next from c into @column" & vbCrLf
    s = s & "end" & vbCrLf
    s = s & "close c" & vbCrLf
    s = s & "deallocate c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "--select @sql" & vbCrLf
    s = s & "set @sql = 'insert into securitygroups(Description' + @sql + ') select ''Copy of ''+isnull(Description,'''') ' + @sql + ' from SecurityGroups where SecGroupID=" & DbQuote(Num, ID) & "'" & vbCrLf
    s = s & "--select @sql" & vbCrLf
    s = s & "exec sp_executesql @sql" & vbCrLf
    Call HFApp.SqlExec(s)
    ID = HFApp.SqlIdentity("SecurityGroups", dbHomeFront)
    
    Call LoadGroups
    r = gGroups.FindRow(ID, 0, gGroups.ColIndex("SecGroupID"), , True)
    If r = -1 Then Exit Sub
    gGroups.Row = r
    Call gGroups.ShowCell(r, gGroups.ColIndex("Description"))
    Call gGroups.EditCell
    
End Sub

Private Sub DeleteUser()
    Dim s As String
    Dim UserID As String
    
    
    With gUsers
        If .Row < 1 Then Exit Sub
        UserID = .TextMatrix(.Row, .ColIndex("UserID"))
    End With
    
    If MsgBox("Are you sure you want to delete this user?", vbOKCancel + vbQuestion, "Homefront") = vbCancel Then Exit Sub
    
    s = ""
    s = s & "delete user_manager where user_id=" & DbQuote(Str, UserID) & vbCrLf
    Call HFApp.SqlExec(s)
    
    Call Sales_DeactivateUser(UserID)
    
    Call LoadUsers(mSecGroupID)
    If gUsers.Rows > 1 Then
        gUsers.Row = 1
        gUsers.SetFocus
        Call gUsers.ShowCell(1, gUsers.ColIndex("UserID"))
    End If
    
End Sub

Private Sub CreateUser()
On Error GoTo eh
    Dim r As Long
    Dim s As String
    Dim gid As Long
    
    With gUsers
        
        '.AddItem "(login id)" & vbTab
        'Call .AutoSize(0, 1)
        
        'r = .FindRow("(login id)", 0, .ColIndex("UserID"), , True)
        'If r = -1 Then Exit Sub
        '.Row = r
        '.SetFocus
        Call gUsers_GotFocus
        Call .ShowCell(r, .ColIndex("userid"))
                
        txtUserID.tag = ""
        txtUserID.Text = "(login id)"
        txtUserFirstName.Text = "New User Name"
        txtUserLastName.Text = ""
        txtUserPswd.Text = ""
        txtUserDesignation.Text = ""
        txtUserOffice.Text = ""
        txtUserEmail.Text = ""
        
        gid = Val("" & gGroups.TextMatrix(gGroups.Row, gGroups.ColIndex("SecGroupID")))
        Call SetListIndex(cboSecurityGroups, gid)
        
        
        Call SetCtrlFocus(txtUserID)
        Call SelectAll(txtUserID)
    End With
    
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "CreateUser")
    End If
End Sub

Private Sub FindUser()
    Dim s As String
    Dim gid As Long
    Dim uid As String
    Dim r As Long
    
    s = ""
    s = s & "select" & vbCrLf
    s = s & " u.user_id LoginID" & vbCrLf
    s = s & ",u.user_name UserName" & vbCrLf
    s = s & ",u.Designation" & vbCrLf
    s = s & ",u.Office" & vbCrLf
    s = s & ",u.CRMAccess CRMUser" & vbCrLf
    s = s & ",u.SecGroupID" & vbCrLf
    s = s & ",g.Description SecurityGroup" & vbCrLf
    s = s & "from user_manager u " & vbCrLf
    s = s & "join securitygroups g on u.secgroupid=g.secgroupid" & vbCrLf
    If Not FPickList.Choose(HFApp.Databases(dbHomeFront), "User", s, , , , , "SecGroupID") Then Exit Sub
    
    gid = Val("" & FPickList.SelectedItem("SecGroupID"))
    uid = FPickList.SelectedItem("LoginID")

    r = gGroups.FindRow(gid, 0, gGroups.ColIndex("SecGroupID"), , True)
    If r = -1 Then
        MsgBox "Group not loaded.", vbInformation, "Homefront"
        Exit Sub
    End If
    gGroups.Row = r
    Call gGroups.ShowCell(r, gGroups.ColIndex("Description"))
    
    
    r = gUsers.FindRow(uid, 0, gUsers.ColIndex("UserID"), True, True)
    If r = -1 Then
        MsgBox "User not loaded.", vbInformation, "Homefront"
        Exit Sub
    End If
    gUsers.Row = r
    Call gUsers.ShowCell(r, gUsers.ColIndex("UserID"))
    gUsers.SetFocus

End Sub


Private Sub Form_Load()
On Error Resume Next
    
    Call IniGetForm(Me)
    UserPage.BorderStyle = 0
    
    'remove tabs largest to smallest so you can use the frame index
    Call TabStrip.Tabs.Remove(TDOCUMENTS - 1)
    If HFApp.LicensedSeats("JobSimplicity") > 0 Then
    Else
        If HFApp.LicensedSeats("CRM") < 1 Then Call TabStrip.Tabs.Remove(TCRM)
        If HFApp.LicensedSeats("Workticket") < 1 Then Call TabStrip.Tabs.Remove(TWORKTICKET)
        If HFApp.LicensedSeats("hfwarranty") < 1 Then Call TabStrip.Tabs.Remove(TWARRANTY)
        If HFApp.LicensedSeats("hfsched") < 1 Then Call TabStrip.Tabs.Remove(TSCHEDULING)
        If HFApp.LicensedSeats("hfpayables") < 1 Then Call TabStrip.Tabs.Remove(TPAYABLES)
        If HFApp.LicensedSeats("hfest") < 1 Then Call TabStrip.Tabs.Remove(TPURCHASING)
        If HFApp.LicensedSeats("hfsales") < 1 Then Call TabStrip.Tabs.Remove(TSALES)
    End If
    
    Call InitPermissionGrids
    Call LoadGroups
    Call LoadDivisions
    Call LoadSalesCommunities


    Call Form_Resize

    Dirty = e_none
    
End Sub

Private Sub LoadDivisions()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    s = "SELECT DivisionID,DivisionCode,DivisionName FROM Divisions ORDER BY 2"
    Set rs = HFApp.SqlExec(s)
    With gDivisions
        .Rows = 1
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, .ColIndex("DivisionID")) = "" & rs("DivisionID")
            .TextMatrix(.Rows - 1, .ColIndex("DivisionCode")) = "" & rs("DivisionCode")
            .TextMatrix(.Rows - 1, .ColIndex("DivisionName")) = "" & rs("DivisionName")
            rs.MoveNext
        Wend
        Call .AutoSize(0, .Cols - 1)
    End With
    
    Dirty = e_none
    
End Sub

Private Sub LoadSalesCommunities()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    s = "SELECT Area,Description from tbllocality where isnull(inactive,0)=0 ORDER BY 1"
    Set rs = HFApp.SqlExec(s)
    With gSalesCommunities
        .Rows = 1
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, .ColIndex("Area")) = "" & rs("Area")
            .TextMatrix(.Rows - 1, .ColIndex("Description")) = "" & rs("Description")
            rs.MoveNext
        Wend
        Call .AutoSize(0, .Cols - 1)
    End With
    
    Dirty = e_none
    
End Sub


Private Sub LoadGroups()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    mSecGroupID = 0
    s = "select SecGroupID,Description from securitygroups order by 2"
    Set rs = HFApp.SqlExec(s)
    With gGroups
    .Rows = 0
    While Not rs.EOF
        .AddItem "" & rs(0) & vbTab & rs(1)
        rs.MoveNext
    Wend
    If .Rows > 0 Then .Row = 0
    End With
    
    
    Call LoadComboBox(cboSecurityGroups, HFApp.Databases(dbHomeFront), "select Description,'',SecGroupID from securitygroups order by 1")
    
    Dirty = e_none
    
    
End Sub

Private Sub LoadUser(UserID As String)
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    Dim b As Boolean
    
    
    
    s = ""
    s = s & "select u.user_id,u.firstname,u.lastname,u.designation,u.office,u.email,u.CRMAccess,u.SecGroupID,u.PMUser,g.user_access" & vbCrLf
    s = s & "from user_manager u" & vbCrLf
    s = s & "left outer join securitygroups g on u.secgroupid=g.secgroupid" & vbCrLf
    s = s & "where u.user_id=" & DbQuote(Str, UserID) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        txtUserID.Text = ""
        txtUserFirstName.Text = ""
        txtUserLastName.Text = ""
        txtUserPswd.Text = ""
        txtUserDesignation.Text = ""
        txtUserOffice.Text = ""
        txtUserEmail.Text = ""
        chkCRMAccess.Value = vbUnchecked
        chkPMUser.Value = vbUnchecked
        cboSecurityGroups.ListIndex = -1
        
        gSalesCommunities.Visible = False
        lblSalesCommunities.Visible = False
    Else
    
        txtUserID.Text = "" & rs("User_id")
        txtUserID.tag = txtUserID.Text
        txtUserFirstName.Text = "" & rs("FirstName")
        txtUserLastName.Text = "" & rs("LastName")
        
        'initial load with random length string of chr(1)'s
        'if they change it we will save it. if not we wont save it :)
        txtUserPswd.Text = String(Rnd * (20 - 6) + 6, Chr(1))
        
        txtUserDesignation.Text = "" & rs("designation")
        txtUserOffice.Text = "" & rs("office")
        txtUserEmail.Text = "" & rs("email")
    
        Call SetListIndex(cboSecurityGroups, Val("" & rs("SecGroupID")))
    
        chkCRMAccess.Value = IIf("" & rs("crmaccess") = "True", vbChecked, vbUnchecked)
        chkPMUser.Value = IIf("" & rs("pmuser") = "True", vbChecked, vbUnchecked)
        
        'hide sales communities unless the user is in a salesperson group
        b = Val("" & rs("User_Access")) = 4 And HFApp.Options(AreaSpecSalesPerson) = "True"
        gSalesCommunities.Visible = b
        lblSalesCommunities.Visible = b
        
    End If
        
    
    
    'hide CRM if not licensed
    b = mCRMSeats <> 0
    chkCRMAccess.Visible = b
    lblCRMAccess.Visible = b
    
    'disable CRM if unchecked and all seats taken
    'always enable if checked so they can uncheck one user to check another
    s = "select count(*) from user_manager where CRMAccess=1"
    Set rs = HFApp.SqlExec(s)
    chkCRMAccess.Enabled = chkCRMAccess.Value = vbChecked Or mCRMSeats > Val("" & rs(0))
    lblCRMAccess.Caption = Val("" & rs(0)) & " of " & mCRMSeats & " seats have been assigned."
    
    
    
    'load user divisions
    With gDivisions
        'unset all
        For r = 1 To .Rows - 1
            .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexUnchecked
        Next
    
        'set selected
        s = "select divisionid from divisionusers where userid=" & DbQuote(Str, UserID)
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = .FindRow(Val("" & rs("DivisionID")), 0, .ColIndex("DivisionID"), , True)
            If r <> -1 Then .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexChecked
            rs.MoveNext
        Wend
    End With
    
    
    'load sales areas
    With gSalesCommunities
    
        'unset all
        For r = 1 To .Rows - 1
            .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexUnchecked
        Next
    
        'set selected
        s = "select area from tblsalespersonarea where sales_person_id=" & DbQuote(Str, UserID)
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = .FindRow("" & rs("area"), 0, .ColIndex("area"), False, True)
            If r <> -1 Then .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexChecked
            rs.MoveNext
        Wend
    End With
    
    
    Dirty = e_none

End Sub

Private Sub LoadGroup(SecGroupID As Long)
    mSecGroupID = SecGroupID
    Call LoadUsers(SecGroupID)
    Call LoadPermissions(SecGroupID)
    Dirty = e_none
End Sub

Private Sub LoadPermissions(SecGroupID As Long)
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    Dim i As Long
    
    s = "select launchSalesMgmt,* from securitygroups where secgroupid=" & DbQuote(Num, SecGroupID)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        Exit Sub
    Else
        'sales
        txtHomePage.Text = "" & rs("HomePage")
        i = Val("" & rs("User_Access"))
        If i < 1 Or i > 7 Then i = 6
        cboUserAccess.ListIndex = i - 1
        'purchasing
        txtMaxPOAmount.Text = Format(Val("" & rs("MaxPOAmount")), "0.00")
        'payables
        txtMaxInvoiceAmount.Text = Format(Val("" & rs("MaxInvoiceAmount")), "0.00")
        txtMaxInvoiceOverrideAmount.Text = Format(Val("" & rs("MaxInvoiceOverrideAmount")), "0.00")
        txtMaxInvoiceOverridePercent.Text = Val("" & rs("MaxInvoiceOverRidePercent"))
        txtMaxInvoicePOOverageAmount.Text = Format(Val("" & rs("MaxInvoicePOOverageAmount")), "0.00")
        txtMaxInvoicePOOveragePercent.Text = Val("" & rs("MaxInvoicePOOveragePercent"))
        
        'all the checkboxes
        For i = 1 To gPermissions.Count
        If i <> TDOCUMENTS Then
        With gPermissions(i)
            For r = 0 To .Rows - 1
                If Not .IsSubtotal(r) Then
                    .Cell(flexcpChecked, r, .ColIndex("Value")) = IIf("" & rs(.TextMatrix(r, .ColIndex("ColumnName"))) = "True", flexChecked, flexUnchecked)
                End If
            Next
        End With
        End If
        Next
    End If
    
    
    With gPermissions(TDOCUMENTS)
        
         'clear document classes
        For r = 0 To .Rows - 1
        If Not .IsSubtotal(r) And .TextMatrix(r, .ColIndex("section")) = "document classes" Then
            .Cell(flexcpChecked, r, .ColIndex("Value")) = flexUnchecked
        End If
        Next
    
        s = "select * from securitygroupdocumentclasses where secgroupid=" & DbQuote(Num, SecGroupID)
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            For r = 0 To .Rows - 1
            If Not .IsSubtotal(r) And .TextMatrix(r, .ColIndex("ColumnName")) = "" & rs("documentclass") Then
                .Cell(flexcpChecked, r, .ColIndex("Value")) = flexChecked
            End If
            Next
            rs.MoveNext
        Wend
    End With
    
    Dirty = e_none
    
End Sub
Private Sub LoadUsers(SecGroupID As Long)
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    s = "select User_id,user_name from user_manager where secgroupid=" & DbQuote(Num, SecGroupID) & " order by 2"
    Set rs = HFApp.SqlExec(s)
    With gUsers
    .Rows = 1
    While Not rs.EOF
        .AddItem "" & rs(0) & vbTab & rs(1)
        rs.MoveNext
    Wend
    Call .AutoSize(0, 1)
    End With
    
End Sub

Private Function GetPageIndex(Mode As String, Key As String) As Long
    'given a key, find the corresponding grid or tab
    
    Dim i As Long
    If Mode = "grid" Then
        For i = 1 To TabFrame.Count
            If TabFrame(i).Caption = Key Then
                GetPageIndex = i
                Exit Function
            End If
        Next
    Else
        For i = 1 To TabStrip.Tabs.Count
            If TabStrip.Tabs(i).Caption = Key Then
                GetPageIndex = i
                Exit Function
            End If
        Next
    End If
    GetPageIndex = -1
    
End Function

Private Sub InitPermissionGrids()
    
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim rs As Recordset
    
    Dim validpages As String
    Dim page As String
    Dim section As String
    Dim ColumnName As String
    Dim Description As String
    
    'get valid tabs for error message
    validpages = ""
    For i = 1 To TabStrip.Tabs.Count
        TabFrame(i).BorderStyle = 0
        validpages = validpages & ", " & TabStrip.Tabs(i).Key
    Next
    validpages = Mid(validpages, 3)
        
    'clear permission trees
    For i = 1 To gPermissions.Count
        gPermissions(i).Rows = 0
        gPermissions(i).ExtendLastCol = True
    Next
        
    'rebuild permission trees
    s = ""
    s = s & "select c.name,cast(p.value as varchar(500)) Description" & vbCrLf
    s = s & "from sys.tables t" & vbCrLf
    s = s & "join sys.columns c on t.object_id=c.object_id" & vbCrLf
    s = s & "join sys.extended_properties p on p.major_id=t.object_id and p.minor_id=c.column_id and p.name='MS_Description'" & vbCrLf
    s = s & "where t.name='SecurityGroups' " & vbCrLf
    s = s & "union all" & vbCrLf
    s = s & "select " & vbCrLf
    s = s & " documentclass name" & vbCrLf
    s = s & ",'Document Management:Document Classes:a:' + documentclass description" & vbCrLf
    s = s & "from dms_documentclasses" & vbCrLf
    s = s & "order by 2" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        'col 0 = column name
        'col 1 = page:section:sort:description
    
        s = "" & rs(1)
        page = Parse(s, 1, ":")
        section = Parse(s, 2, ":")
        Description = Parse(s, 4, ":")
        ColumnName = "" & rs(0)
                
        i = GetPageIndex("Grid", page)
        If i = -1 Then
            MsgBox "Error loading permissions. Invalid tab """ & page & """." & vbCrLf & vbCrLf & "Valid tabs: " & validpages, vbExclamation
        Else
            With gPermissions(i)
                Call .AddItem("")
                .TextMatrix(.Rows - 1, .ColIndex("Section")) = section
                .TextMatrix(.Rows - 1, .ColIndex("ColumnName")) = ColumnName
                .TextMatrix(.Rows - 1, .ColIndex("Description")) = Description
            End With
        End If
        
        rs.MoveNext
    Wend
    
    'apply grouping and formatting
    For i = 1 To gPermissions.Count
    With gPermissions(i)
    
        .GridLines = flexGridNone
        .Subtotal flexSTNone, 0, , , , vbHighlight, True, , , True
        .ColWidth(.ColIndex("Section")) = 360
        Call .AutoSize(.ColIndex("Description"))
        Call .AutoSize(.ColIndex("Value"))
        
        'grid has no rows so hide its tab
        If .Rows = 0 Then
            j = GetPageIndex("Tab", TabFrame(i).Caption)
            If j <> -1 Then Call TabStrip.Tabs.Remove(j)
        End If
        
    End With
    Next
    
    
    With cboUserAccess
        .Clear
        Call .AddItem("1 - Administrators")
        Call .AddItem("2 - Sales Managers")
        Call .AddItem("3 - Accounting")
        Call .AddItem("4 - Sales People")
        Call .AddItem("5 - Design Center Sales")
        Call .AddItem("6 - Read Only")
        Call .AddItem("7 - Estimating")
    End With
    
    
    mCRMSeats = HFApp.LicensedSeats("CRM")
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
Const margin = 60
    Dim i As Long
    
    HSlider.Min = 1500
    HSlider.Max = Me.ScaleHeight - 1500
    VSlider.Min = 1500
    VSlider.Max = Me.ScaleWidth - 1500
    
    cmdNav(1).Move Me.ScaleWidth - margin - cmdNav(1).Width, Me.ScaleHeight - margin - cmdNav(1).Height
    
    VSlider.Move Max(Min(VSlider.left, VSlider.Max), VSlider.Min), 0, VSlider.Width, Me.ScaleHeight
    HSlider.Move gGroups.left, Max(Min(HSlider.Top, HSlider.Max), HSlider.Min), VSlider.left - gGroups.left, HSlider.Height
    
    gGroups.Move gGroups.left, gGroups.Top, VSlider.left - gGroups.left, HSlider.Top - gGroups.Top
    lblUsers.Top = HSlider.Top + 3 * margin
    gUsers.Move gGroups.left, lblUsers.Top + lblUsers.Height + margin, gGroups.Width, Me.ScaleHeight - lblUsers.Top - lblUsers.Height - 3 * margin
    cmdTools(3).Top = gUsers.Top - cmdTools(3).Height - 30
    cmdTools(4).Top = cmdTools(3).Top
    cmdTools(5).Top = cmdTools(3).Top
    
    
    TabStrip.Move VSlider.left + VSlider.Width, margin, Me.ScaleWidth - VSlider.left - VSlider.Width - margin, Me.ScaleHeight - 3 * margin - cmdNav(1).Height
    UserPage.Move TabStrip.left, TabStrip.Top, TabStrip.Width, TabStrip.Height
    
    For i = 1 To TabFrame.Count
        TabFrame(i).Move TabStrip.ClientLeft, TabStrip.ClientTop, TabStrip.ClientWidth, TabStrip.ClientHeight
        TabFrame(i).Visible = TabStrip.Visible And TabStrip.SelectedItem.Caption = TabFrame(i).Caption
        
        Select Case i
        Case TSALES, TPURCHASING, TPAYABLES
            gPermissions(i).Move 0, gPermissions(i).Top, TabFrame(i).Width, TabFrame(i).Height - gPermissions(i).Top
        Case Else
            gPermissions(i).Move 0, 0, TabFrame(i).Width, TabFrame(i).Height
        End Select
        
    Next

End Sub



Private Sub Form_Unload(Cancel As Integer)
    Cancel = IsDirty()
    Call IniPutForm(Me)
End Sub

Private Sub gDivisions_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dirty = e_user
End Sub

Private Sub gDivisions_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gDivisions
        Select Case .ColKey(Col)
            Case "Selected"
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub gDivisions_KeyDown(KeyCode As Integer, Shift As Integer)
    With gDivisions
        If KeyCode = vbKeySpace And .Col <> .ColIndex("Selected") Then
            .Cell(flexcpChecked, .Row, .ColIndex("Selected")) = IIf(.Cell(flexcpChecked, .Row, .ColIndex("Selected")) = flexChecked, flexUnchecked, flexChecked)
            Dirty = e_user
        End If
    End With
End Sub

Private Sub gGroups_GotFocus()
    UserPage.Visible = False
    TabStrip.Visible = True
    Call Form_Resize
End Sub



Private Sub gGroups_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    
    With gGroups
        s = ""
        s = s & "update securitygroups" & vbCrLf
        s = s & "set description=" & DbQuote(Str, .EditText) & vbCrLf
        s = s & "where secgroupid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("SecGroupID"))) & vbCrLf
        Call HFApp.SqlExec(s)
    End With
    Call LoadComboBox(cboSecurityGroups, HFApp.Databases(dbHomeFront), "select Description,'',SecGroupID from securitygroups order by 1")
    
End Sub

Private Sub gPermissions_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dirty = e_perms
End Sub

Private Sub gPermissions_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gPermissions(Index)
        Select Case .ColKey(Col)
            Case "Value"
                .AutoSearch = flexSearchNone
            Case Else
                .AutoSearch = flexSearchFromTop
                Cancel = True
        End Select
        If .IsSubtotal(Row) Then Cancel = True
    End With
End Sub

Private Sub gPermissions_DrawCell(Index As Integer, ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    With gPermissions(Index)
        If Col = 0 And Row > .FixedRows And Not .IsSubtotal(Row) Then
            Done = True
        End If
    End With
End Sub

Private Sub gPermissions_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    With gPermissions(Index)
        If KeyCode = vbKeySpace And Not .IsSubtotal(.Row) And .Col <> .ColIndex("Value") Then
            .Cell(flexcpChecked, .Row, .ColIndex("Value")) = IIf(.Cell(flexcpChecked, .Row, .ColIndex("Value")) = flexChecked, flexUnchecked, flexChecked)
            Dirty = e_perms
        End If
    End With
End Sub

Private Sub gPermissions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    With gPermissions(Index)
        If InIde() Then
            .ToolTipText = .TextMatrix(.MouseRow, .ColIndex("ColumnName"))
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub gSalesCommunities_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dirty = e_user
End Sub

Private Sub gSalesCommunities_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gSalesCommunities
        Select Case .ColKey(Col)
            Case "Selected"
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub gSalesCommunities_KeyDown(KeyCode As Integer, Shift As Integer)
    With gSalesCommunities
        If KeyCode = vbKeySpace And .Col <> .ColIndex("Selected") Then
            .Cell(flexcpChecked, .Row, .ColIndex("Selected")) = IIf(.Cell(flexcpChecked, .Row, .ColIndex("Selected")) = flexChecked, flexUnchecked, flexChecked)
            Dirty = e_user
        End If
    End With
End Sub

Private Sub gUsers_GotFocus()
    UserPage.Visible = True
    TabStrip.Visible = False
    Call Form_Resize
End Sub




Private Sub HSlider_Move()
    Call Form_Resize
End Sub


Private Sub txtMaxInvoiceAmount_Change()
    Dirty = e_perms
End Sub

Private Sub txtMaxInvoiceOverrideAmount_Change()
    Dirty = e_perms
End Sub
Private Sub txtMaxInvoiceOverridePercent_Change()
    Dirty = e_perms
End Sub
Private Sub txtMaxInvoicePOOverageAmount_Change()
    Dirty = e_perms
End Sub
Private Sub txtMaxInvoicePOOveragePercent_Change()
    Dirty = e_perms
End Sub


Private Sub txtMaxInvoiceAmount_GotFocus()
    SelectAll txtMaxInvoiceAmount
End Sub

Private Sub txtMaxInvoiceOverrideAmount_GotFocus()
    SelectAll txtMaxInvoiceOverrideAmount
End Sub
Private Sub txtMaxInvoiceOverridePercent_GotFocus()
    SelectAll txtMaxInvoiceOverridePercent
End Sub
Private Sub txtMaxInvoicePOOverageAmount_GotFocus()
    SelectAll txtMaxInvoicePOOverageAmount
End Sub
Private Sub txtMaxInvoicePOOveragePercent_GotFocus()
    SelectAll txtMaxInvoicePOOveragePercent
End Sub



Private Sub txtMaxInvoiceAmount_Validate(Cancel As Boolean)
    txtMaxInvoiceAmount.Text = Format(Val(txtMaxInvoiceAmount.Text), "0.00")
End Sub
Private Sub txtMaxInvoiceOverrideAmount_Validate(Cancel As Boolean)
    txtMaxInvoiceOverrideAmount.Text = Format(Val(txtMaxInvoiceOverrideAmount.Text), "0.00")
End Sub
Private Sub txtMaxInvoiceOverridePercent_Validate(Cancel As Boolean)
    txtMaxInvoiceOverridePercent.Text = Abs(Val(txtMaxInvoiceOverridePercent))
End Sub
Private Sub txtMaxInvoicePOOverageAmount_Validate(Cancel As Boolean)
    txtMaxInvoicePOOverageAmount.Text = Format(Val(txtMaxInvoicePOOverageAmount.Text), "0.00")
End Sub
Private Sub txtMaxInvoicePOOveragePercent_Validate(Cancel As Boolean)
    txtMaxInvoicePOOveragePercent.Text = Abs(Val(txtMaxInvoicePOOveragePercent))
End Sub


Private Sub txtUserID_GotFocus()
    SelectAll txtUserID
End Sub
Private Sub txtUserFirstName_GotFocus()
    SelectAll txtUserFirstName
End Sub
Private Sub txtUserlastName_GotFocus()
    SelectAll txtUserLastName
End Sub
Private Sub txtUserPswd_GotFocus()
    SelectAll txtUserPswd
End Sub
Private Sub txtUserDesignation_GotFocus()
    SelectAll txtUserDesignation
End Sub
Private Sub txtUserOffice_GotFocus()
    SelectAll txtUserOffice
End Sub
Private Sub txtUserEmail_GotFocus()
    SelectAll txtUserEmail
End Sub
Private Sub txtHomePage_GotFocus()
    SelectAll txtHomePage
End Sub
Private Sub txtMaxPOAmount_GotFocus()
    SelectAll txtMaxPOAmount
End Sub

Private Sub txtUserPswd_KeyDown(KeyCode As Integer, Shift As Integer)
    If left(txtUserPswd.Text, 1) = Chr(1) Then
        txtUserPswd.Text = ""
    End If
End Sub

Private Sub VSlider_Move()
    Call Form_Resize
End Sub

Private Sub TabStrip_Click()
    Call Form_Resize
End Sub

Private Sub txtHomePage_Change()
    Dirty = e_perms
End Sub

Private Sub txtMaxPOAmount_Change()
    Dirty = e_perms
End Sub

Private Sub txtMaxPOAmount_Validate(Cancel As Boolean)
    txtMaxPOAmount.Text = Format(Val(txtMaxPOAmount.Text), "0.00")
End Sub

Private Sub txtUserDesignation_Change()
    Dirty = e_user
End Sub

Private Sub txtUserEmail_Change()
    Dirty = e_user
End Sub

Private Sub txtUserID_Change()
    Dirty = e_user
End Sub

Private Sub txtUserFirstName_Change()
    Dirty = e_user
End Sub
Private Sub txtUserLastName_Change()
    Dirty = e_user
End Sub

Private Sub txtUserOffice_Change()
    Dirty = e_user
End Sub

Private Sub txtUserPswd_Change()
    Dirty = e_user
End Sub



Private Sub gGroups_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal newrow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim ID As Long
    
    If IsDirty() Then Exit Sub

    With gGroups
        If newrow > -1 Then
            ID = Val("" & .TextMatrix(newrow, .ColIndex("SecGroupID")))
            If ID <> 0 Then Call LoadGroup(ID)
        End If
    End With

End Sub
Private Sub gUsers_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal newrow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim ID As String
    
    If IsDirty() Then Exit Sub
    
    With gUsers
        If newrow > -1 Then
            ID = "" & .TextMatrix(newrow, .ColIndex("UserID"))
            If ID <> "" Then
                mUserIndex = newrow
                Call LoadUser(ID)
            End If
        End If
    End With
    
End Sub
Private Function IsDirty() As Boolean
    If Dirty <> e_none Then
        Select Case MsgBox("Do you want to save your changes?", vbYesNoCancel + vbQuestion, "HomeFront")
        Case vbYes
            If Dirty = e_perms Then Call SaveGroup
            If Dirty = e_user Then Call SaveUser
        Case vbNo
            Dirty = e_none
        End Select
    End If
    IsDirty = Dirty <> e_none
    
End Function


Private Function SaveUser() As Boolean
On Error GoTo eh
    Dim hff   As Object
    Dim s As String
    Dim r As Long
    Dim oldid As String
    Dim newid As String
    Dim pswdChanged As Boolean
    
    
    oldid = txtUserID.tag
    newid = txtUserID.Text
    
    
    If Len(newid) > 10 Then
        Screen.MousePointer = vbDefault
        Call MsgBox("User id is too long. It must be 10 characters or less." & vbCrLf & vbCrLf & s, vbExclamation, "HomeFront")
        SaveUser = False
        Exit Function
    End If
    
    
    If oldid = "" Then
        oldid = newid
        s = "insert user_manager(user_id,SecGroupID) values(" & DbQuote(Str, newid) & "," & DbQuote(Num, GetComboBoxListID(cboSecurityGroups)) & ")"
        HFApp.SqlExec s
    End If
    
    
    pswdChanged = left(txtUserPswd.Text, 1) <> Chr(1)
    If pswdChanged Then
        s = "select dbo.ZYB_IsPswdComplex(" & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, oldid) & "," & DbQuote(Str, txtUserPswd.Text) & ")"
        s = "" & HFApp.SqlExec(s)(0)
        If s <> "" Then
            Screen.MousePointer = vbDefault
            Call MsgBox("Unable to change password. It does not meet complexity rules." & vbCrLf & vbCrLf & s, vbExclamation, "HomeFront")
            SaveUser = False
            Exit Function
        End If
        
        If txtUserPswd.Text = "" Then
            If MsgBox("Are you sure you want to leave the password blank?" & vbCrLf & vbCrLf & s, vbYesNo + vbQuestion, "HomeFront") = vbNo Then
                SaveUser = False
                Exit Function
            End If
        End If
        
        Set hff = CreateObject("ZYBFunctions.HomeFrontFunctions")
    End If
    
    
    s = ""
    s = s & "update user_manager set" & vbCrLf
    s = s & " user_id=" & DbQuote(Str, newid) & vbCrLf
    If pswdChanged Then
        s = s & ",user_Password=" & DbQuote(Str, hff.MyCrypt(txtUserPswd.Text, "Fazlul")) & vbCrLf
    End If
    s = s & ",firstname=" & DbQuote(Str, txtUserFirstName.Text) & vbCrLf
    s = s & ",lastname=" & DbQuote(Str, txtUserLastName.Text) & vbCrLf
    s = s & ",designation=" & DbQuote(Str, txtUserDesignation.Text) & vbCrLf
    s = s & ",office=" & DbQuote(Str, txtUserOffice.Text) & vbCrLf
    s = s & ",email=" & DbQuote(Str, txtUserEmail.Text) & vbCrLf
    s = s & ",crmaccess=" & DbQuote(Bit, chkCRMAccess.Value = vbChecked) & vbCrLf
    s = s & ",pmuser=" & DbQuote(Bit, chkPMUser.Value = vbChecked) & vbCrLf
    s = s & ",SecGroupID=" & DbQuote(Num, GetComboBoxListID(cboSecurityGroups)) & vbCrLf
    s = s & "where user_id=" & DbQuote(Str, oldid) & vbCrLf
    
    s = s & vbCrLf '---------------------------------------------------
    s = s & "update contacts set" & vbCrLf
    s = s & " userid=" & DbQuote(Str, newid) & vbCrLf
    s = s & ",email=" & DbQuote(Str, txtUserEmail.Text, , , 250) & vbCrLf
    s = s & ",firstname=" & DbQuote(Str, txtUserFirstName.Text) & vbCrLf
    s = s & ",lastname=" & DbQuote(Str, txtUserLastName.Text) & vbCrLf
    s = s & ",displayname=" & DbQuote(Str, txtUserFirstName.Text & " " & txtUserLastName.Text) & vbCrLf
    s = s & "where userid=" & DbQuote(Str, oldid) & vbCrLf
    s = s & "and contacttypeid in(1,4)" & vbCrLf
    
    s = s & vbCrLf '---------------------------------------------------
    s = s & "update tblsales_persons set" & vbCrLf
    s = s & " sales_person_id=" & DbQuote(Str, newid) & vbCrLf
    s = s & ",sales_person_name=" & DbQuote(Str, txtUserFirstName.Text & " " & txtUserLastName.Text, , , 20) & vbCrLf
    s = s & ",email=" & DbQuote(Str, txtUserEmail.Text, , , 250) & vbCrLf
    s = s & "where sales_person_id=" & DbQuote(Str, oldid) & vbCrLf
    
    s = s & vbCrLf '---------------------------------------------------
    s = s & "update tblprojectmanager set" & vbCrLf
    s = s & " pm=" & DbQuote(Str, newid) & vbCrLf
    s = s & ",pmname=" & DbQuote(Str, txtUserFirstName.Text & " " & txtUserLastName.Text, , , 20) & vbCrLf
    s = s & ",email=" & DbQuote(Str, txtUserEmail.Text, , , 250) & vbCrLf
    s = s & ",inactive=" & DbQuote(Bit, chkPMUser.Value = vbUnchecked) & vbCrLf
    If pswdChanged Then
        s = s & ",password=" & DbQuote(Str, txtUserPswd.Text) & vbCrLf
    End If
    s = s & "where pm=" & DbQuote(Str, oldid) & vbCrLf
    Call HFApp.SqlExec(s)
    txtUserEmail.tag = txtUserEmail.Text
    
    'save divisions
    With gDivisions
        s = "delete from divisionusers where userid=" & DbQuote(Str, oldid)
        Call HFApp.SqlExec(s)
        
        For r = 1 To .Rows - 1
            If .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexChecked Then
                s = "insert into divisionusers(divisionid,userid) values(" & DbQuote(Num, .TextMatrix(r, .ColIndex("DivisionID"))) & "," & DbQuote(Str, newid) & ")"
                Call HFApp.SqlExec(s)
            End If
        Next
    End With
    
    
    'save sales areas
    With gSalesCommunities
        s = "delete from tblsalespersonarea where sales_person_id=" & DbQuote(Str, oldid)
        Call HFApp.SqlExec(s)
        
        For r = 1 To .Rows - 1
            If .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexChecked Then
                s = "insert into tblsalespersonarea(area,sales_person_id) values(" & DbQuote(Str, .TextMatrix(r, .ColIndex("area"))) & "," & DbQuote(Str, newid) & ")"
                Call HFApp.SqlExec(s)
            End If
        Next
    End With
    
    
    Call InsertPMSP
    Call HFApp.SqlExec("ZYB_UpdateUserPerms " & DbQuote(Num, GetComboBoxListID(cboSecurityGroups)))
    Call SaveGroup
    
    Dirty = e_none
    UserPage.Visible = False
    TabStrip.Visible = True
    Call Form_Resize
    Call LoadGroup(mSecGroupID)
    Call LoadUsers(mSecGroupID)
    gGroups.SetFocus
    
Exit Function
eh: If Err.Description Like "*Duplicate*" Then
        MsgBox "Duplicate found. Make sure user ids are unique.", vbExclamation, App.ProductName
    Else
        Call errHandler(SRCFILE & "SaveUser", s)
    End If
End Function


Private Sub SaveGroup()
    
    Dim s As String
    Dim r As Long
    Dim i As Long
    
    s = ""
'do both of these :)
'    s = s & "update securitygroups set" & vbCrLf
'    s = s & "update user_manager set" & vbCrLf

    'sales
    s = s & " HomePage=" & DbQuote(Str, txtHomePage.Text) & vbCrLf
    s = s & ",User_Access=" & DbQuote(Num, cboUserAccess.ListIndex + 1) & vbCrLf
    'purchasing
    s = s & ",MaxPOAmount=" & DbQuote(Num, txtMaxPOAmount.Text) & vbCrLf
    'payables
    s = s & ",MaxInvoiceAmount=" & DbQuote(Num, txtMaxInvoiceAmount.Text) & vbCrLf
    s = s & ",MaxInvoiceOverrideAmount=" & DbQuote(Num, txtMaxInvoiceOverrideAmount.Text) & vbCrLf
    s = s & ",MaxInvoiceOverridePercent=" & DbQuote(Num, txtMaxInvoiceOverridePercent.Text) & vbCrLf
    s = s & ",MaxInvoicePOOverageAmount=" & DbQuote(Num, txtMaxInvoicePOOverageAmount.Text) & vbCrLf
    s = s & ",MaxInvoicePOOveragePercent=" & DbQuote(Num, txtMaxInvoicePOOveragePercent.Text) & vbCrLf
    
    'all the checkboxes
    For i = 1 To gPermissions.Count
    If i <> TDOCUMENTS Then
    With gPermissions(i)
        For r = 0 To .Rows - 1
            If Not .IsSubtotal(r) Then
                s = s & ",[" & .TextMatrix(r, .ColIndex("ColumnName")) & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("Value")) = flexChecked) & vbCrLf
            End If
        Next
    End With
    End If
    Next
    s = s & "where SecGroupID=" & DbQuote(Num, mSecGroupID) & vbCrLf
    
    Call HFApp.SqlExec("update securitygroups set" & vbCrLf & s)
    Call HFApp.SqlExec("update user_manager set" & vbCrLf & s)
    
    
    
    With gPermissions(TDOCUMENTS)
        s = ""
        For r = 0 To .Rows - 1
            If Not .IsSubtotal(r) Then
            If .Cell(flexcpChecked, r, .ColIndex("Value")) = flexChecked Then
                s = s & ",(" & DbQuote(Num, mSecGroupID) & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("ColumnName"))) & ")" & vbCrLf
            End If
            End If
        Next
        If s = "" Then
            s = "delete securitygroupdocumentclasses where SecGroupID=" & DbQuote(Num, mSecGroupID)
        Else
            s = "delete securitygroupdocumentclasses where SecGroupID=" & DbQuote(Num, mSecGroupID) & vbCrLf & _
                "insert securitygroupdocumentclasses(SecGroupID,documentclass) values" & vbCrLf & Mid(s, 2)
        End If
        Call HFApp.SqlExec(s)
    End With
    
    
    
    Call InsertPMSP
    
    Dirty = e_none
End Sub


Private Sub InsertPMSP()
On Error Resume Next
    Dim s As String
    Dim rs As Recordset
    
    
    'create warranty user contact--------------------------------------------------------------------
    s = ""
    s = s & "insert into contacts(userid,firstname,lastname,displayname,email,contacttypeid)" & vbCrLf
    s = s & "select u.user_id,u.firstname,u.lastname,u.user_name,u.email,1" & vbCrLf
    s = s & "from user_manager u" & vbCrLf
    s = s & "join securitygroups g on u.secgroupid=g.secgroupid and g.launchwarranty=1" & vbCrLf
    s = s & "left outer join contacts s on u.user_id=s.userid and s.contacttypeid=1" & vbCrLf
    s = s & "where s.userid is null" & vbCrLf
    Call HFApp.SqlExec(s)
    
    
    
    'create salespersons--------------------------------------------------------------------
    s = ""
    s = s & "insert into tblsales_persons(sales_person_id,sales_person_name,email)" & vbCrLf
    s = s & "select u.user_id,u.user_name,u.email" & vbCrLf
    s = s & "from user_manager u" & vbCrLf
    s = s & "join securitygroups g on u.secgroupid=g.secgroupid and g.user_access in(4,5)" & vbCrLf
    s = s & "left outer join tblsales_persons s on u.user_id=s.sales_person_id" & vbCrLf
    s = s & "where s.sales_person_id is null" & vbCrLf
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "insert into contacts(userid,firstname,lastname,displayname,email,contacttypeid)" & vbCrLf
    s = s & "select u.user_id,u.firstname,u.lastname,u.user_name,u.email,4" & vbCrLf
    s = s & "from user_manager u" & vbCrLf
    s = s & "join securitygroups g on u.secgroupid=g.secgroupid and g.user_access in(4,5)" & vbCrLf
    s = s & "left outer join contacts s on u.user_id=s.userid and s.contacttypeid=4" & vbCrLf
    s = s & "where s.userid is null" & vbCrLf
    Call HFApp.SqlExec(s)
    
    
    
    'create project managers--------------------------------------------------------------------
    s = ""
    s = s & "insert into tblprojectmanager(pm,pmname,email)" & vbCrLf
    s = s & "select u.user_id,u.user_name,u.email" & vbCrLf
    s = s & "from user_manager u" & vbCrLf
    s = s & "left outer join tblprojectmanager p on u.user_id=p.pm" & vbCrLf
    s = s & "where u.PMUser=1 and p.pm is null" & vbCrLf
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "insert into contacts(userid,firstname,lastname,displayname,email,contacttypeid)" & vbCrLf
    s = s & "select u.user_id,u.firstname,u.lastname,u.user_name,u.email,5" & vbCrLf
    s = s & "from user_manager u" & vbCrLf
    s = s & "left outer join contacts s on u.user_id=s.userid and s.contacttypeid=5" & vbCrLf
    s = s & "where u.pmuser=1 and s.userid is null" & vbCrLf
    Call HFApp.SqlExec(s)
    
    
    
    'create default grid layouts for any user that has none----------------------------------------
    s = ""
    s = s & "select u.user_id" & vbCrLf
    s = s & "from user_manager u" & vbCrLf
    s = s & "left outer join appgridlayout l on u.user_id=l.uid" & vbCrLf
    s = s & "where l.uid is null" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        Call HFApp.SqlExec("ZYB_ResetUserGridLayout " & DbQuote(Str, "" & rs(0)))
        rs.MoveNext
    Wend
    
    
End Sub





















Private Sub gUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call gUsers.OLEDrag
    End If
End Sub

Private Sub gUsers_OLECompleteDrag(Effect As Long)
    With gGroups
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, .Cols - 1) = vbWindowBackground
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbWindowText
    End With
End Sub


Private Sub gUsers_OLEStartDrag(Data As VSFlex8Ctl.VSDataObject, AllowedEffects As Long)
    Dim i As Long
    Dim s As String
    With gUsers
        For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            s = s & "," & .TextMatrix(i, .ColIndex("UserID"))
        Next
    End With
    s = Chr(2) & Mid(s, 2)
    Call Data.SetData(s, 1)
End Sub




Private Sub gGroups_OLEDragOver(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
    Dim r As Long
    With gGroups
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, .Cols - 1) = vbWindowBackground
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbWindowText
        
        r = .MouseRow
        If r < 0 Then
            Effect = vbDropEffectNone
            Exit Sub
        End If

        If left(Data.GetData(vbCFText), 1) <> Chr(2) Then
            Effect = vbDropEffectNone
            Exit Sub
        End If
                
        .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbHighlight
        .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbHighlightText
        
        Effect = vbDropEffectCopy
        
    End With
End Sub


Private Sub gGroups_OLEDragDrop(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim r As Long
    Dim UserID As String
    Dim s As String
    
    With gGroups
        r = .MouseRow
        If r < 0 Then Exit Sub

        If left(Data.GetData(vbCFText), 1) <> Chr(2) Then Exit Sub
        UserID = Mid(Data.GetData(vbCFText), 2)
        
        If MsgBox("Are you sure you want to move " & UserID & " to " & .TextMatrix(r, .ColIndex("Description")), vbQuestion + vbYesNo, "Homefront") = vbNo Then Exit Sub
        
        s = "update user_manager set secgroupid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("SecGroupID"))) & " where user_id=" & DbQuote(Str, UserID)
        Call HFApp.SqlExec(s)
        Call HFApp.SqlExec("ZYB_UpdateUserPerms " & DbQuote(Num, .ValueMatrix(r, .ColIndex("SecGroupID"))))
        
        gGroups.SetFocus
        Call LoadGroup(mSecGroupID)

               
    End With
End Sub

Private Sub CleanUNnamed()
    Dim s As String
    
End Sub
