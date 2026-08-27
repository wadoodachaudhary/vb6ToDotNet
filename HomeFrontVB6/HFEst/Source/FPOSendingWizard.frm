VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FSendingWizard 
   Caption         =   "PO Sending Wizard"
   ClientHeight    =   5835
   ClientLeft      =   5145
   ClientTop       =   4185
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FPOSendingWizard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6585
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3600
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   900
      Width           =   6315
      Begin VB.OptionButton chkSend 
         Caption         =   "Resend an RFP"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   52
         Top             =   1560
         Width           =   4455
      End
      Begin VB.OptionButton chkSend 
         Caption         =   "Send new RFP's"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   51
         Top             =   1290
         Width           =   4455
      End
      Begin VB.OptionButton chkSend 
         Caption         =   "Send new purchase orders"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   25
         Top             =   450
         Value           =   -1  'True
         Width           =   4455
      End
      Begin VB.OptionButton chkSend 
         Caption         =   "Resend a purchase order"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   24
         Top             =   720
         Width           =   4455
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   0
         Left            =   360
         Picture         =   "FPOSendingWizard.frx":000C
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Choose the task you would like to perform"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   26
         Top             =   150
         Width           =   3915
      End
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6585
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5250
      Width           =   6585
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   26540
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   1588
      Caption         =   "Send purchase orders to vendors"
      Description     =   "The PO Sending wizard will help you deliver your purchase orders."
      Icon            =   "FPOSendingWizard.frx":08D6
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3660
      Index           =   1
      Left            =   0
      TabIndex        =   19
      Top             =   870
      Width           =   6345
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   3300
         Index           =   1
         Left            =   1170
         TabIndex        =   20
         Top             =   270
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
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
         FormatString    =   $"FPOSendingWizard.frx":11B0
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
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All"
         Height          =   315
         Index           =   1
         Left            =   5310
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select Communities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1170
         TabIndex        =   22
         Top             =   30
         Width           =   1665
      End
      Begin VB.Image imgError 
         Height          =   480
         Left            =   330
         Picture         =   "FPOSendingWizard.frx":1202
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3600
      Index           =   2
      Left            =   8970
      TabIndex        =   32
      Top             =   5670
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   3300
         Index           =   2
         Left            =   1170
         TabIndex        =   33
         Top             =   270
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FPOSendingWizard.frx":1ACC
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
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All"
         Height          =   315
         Index           =   2
         Left            =   5310
         TabIndex        =   34
         Top             =   0
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   330
         Picture         =   "FPOSendingWizard.frx":1B33
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblJobs 
         AutoSize        =   -1  'True
         Caption         =   "Select Jobs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1170
         TabIndex        =   35
         Top             =   30
         Width           =   1005
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3600
      Index           =   3
      Left            =   2160
      TabIndex        =   15
      Top             =   4800
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   3300
         Index           =   3
         Left            =   1170
         TabIndex        =   16
         Top             =   270
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
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
         FormatString    =   $"FPOSendingWizard.frx":23FD
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
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All"
         Height          =   315
         Index           =   3
         Left            =   5310
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select Vendors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1170
         TabIndex        =   18
         Top             =   30
         Width           =   1305
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   330
         Picture         =   "FPOSendingWizard.frx":244A
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "3"
      Height          =   3600
      Index           =   4
      Left            =   1500
      TabIndex        =   11
      Top             =   4920
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   3300
         Index           =   4
         Left            =   1170
         TabIndex        =   12
         Top             =   270
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FPOSendingWizard.frx":2D14
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
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All"
         Height          =   315
         Index           =   4
         Left            =   5310
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select PO Groups"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1170
         TabIndex        =   14
         Top             =   30
         Width           =   1530
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   1
         Left            =   330
         Picture         =   "FPOSendingWizard.frx":2D58
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "3"
      Height          =   3600
      Index           =   5
      Left            =   2040
      TabIndex        =   27
      Top             =   6150
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   3300
         Index           =   5
         Left            =   1170
         TabIndex        =   28
         Top             =   270
         Width           =   5175
         _cx             =   9128
         _cy             =   5821
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FPOSendingWizard.frx":3622
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
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   1
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
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All"
         Height          =   315
         Index           =   5
         Left            =   5310
         TabIndex        =   29
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   2
         Left            =   330
         Picture         =   "FPOSendingWizard.frx":3671
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select Purchase Orders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1170
         TabIndex        =   30
         Top             =   30
         Width           =   2025
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4980
      Index           =   6
      Left            =   8460
      TabIndex        =   6
      Top             =   1890
      Visible         =   0   'False
      Width           =   6315
      Begin VB.PictureBox PrinterInfo 
         BorderStyle     =   0  'None
         Height          =   2235
         Left            =   1140
         ScaleHeight     =   2235
         ScaleWidth      =   5235
         TabIndex        =   36
         Top             =   990
         Visible         =   0   'False
         Width           =   5235
         Begin VB.Frame PrinterFrame 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   2235
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   5235
            Begin VB.ComboBox cboPrinters 
               Height          =   240
               Left            =   1020
               TabIndex        =   38
               Top             =   300
               Width           =   3570
               _ExtentX        =   6297
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
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "2 documents are set to print. They will be sent to:"
               Height          =   195
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Width           =   3465
            End
            Begin VB.Label lblPort 
               AutoSize        =   -1  'True
               Caption         =   "Port:            <hidden>"
               Height          =   195
               Left            =   240
               TabIndex        =   48
               Top             =   1740
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VB.Label lblComments 
               AutoSize        =   -1  'True
               Caption         =   "<Printer comments here>"
               Height          =   195
               Left            =   1020
               TabIndex        =   47
               Top             =   1380
               Width           =   1755
            End
            Begin VB.Label lblLocation 
               AutoSize        =   -1  'True
               Caption         =   "Red Deer"
               Height          =   195
               Left            =   1020
               TabIndex        =   46
               Top             =   1140
               Width           =   690
            End
            Begin VB.Label lblModel 
               AutoSize        =   -1  'True
               Caption         =   "HP 2300 Laser Color"
               Height          =   195
               Left            =   1020
               TabIndex        =   45
               Top             =   900
               Width           =   1470
            End
            Begin VB.Label lblStatus 
               AutoSize        =   -1  'True
               Caption         =   "ready"
               Height          =   195
               Left            =   1020
               TabIndex        =   44
               Top             =   660
               Width           =   390
            End
            Begin VB.Label Label1 
               Caption         =   "Comment:"
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   43
               Top             =   1380
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Where:"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   42
               Top             =   1140
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Model:"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   41
               Top             =   900
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Status:"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   40
               Top             =   660
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Printer: "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   39
               Top             =   300
               Width           =   765
            End
         End
      End
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   315
         Left            =   1170
         TabIndex        =   7
         Top             =   1230
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         Picture         =   "FPOSendingWizard.frx":3F3B
         ForeColor       =   0
         BarPicture      =   "FPOSendingWizard.frx":3F57
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
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   22000
         Left            =   1410
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   2100
         Width           =   4815
      End
      Begin VB.Label lblFinish 
         Caption         =   "Ready for processing. You have selected 412 purchase orders for sending. Click Finish to send these documents now."
         Height          =   675
         Left            =   1110
         TabIndex        =   31
         Top             =   210
         Width           =   4125
      End
      Begin VB.Label lblItemDesc 
         AutoSize        =   -1  'True
         Caption         =   "BL07001/053 -- Breckenridge Excavating"
         Height          =   195
         Left            =   1410
         TabIndex        =   10
         Top             =   1830
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order 4 of 14"
         Height          =   195
         Left            =   1410
         TabIndex        =   9
         Top             =   1590
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label lblSaving 
         Caption         =   "Processing:"
         Height          =   255
         Left            =   1170
         TabIndex        =   8
         Top             =   990
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   330
         Picture         =   "FPOSendingWizard.frx":3F73
         Top             =   270
         Width           =   480
      End
   End
End
Attribute VB_Name = "FSendingWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FSendingWizard::"

Private mJob    As String
Private bInHere As Boolean

Private WithEvents SMTPMail As HFMailer.SMTP
Attribute SMTPMail.VB_VarHelpID = -1


Private Sub chkSelectAll_Click(Index As Integer)
On Error Resume Next
    If Not bInHere Then
        gData(Index).Cell(flexcpChecked, 0, 0, gData(Index).Rows - 1, 0) = IIf(chkSelectAll(Index).Value = vbChecked, flexChecked, flexUnchecked)
    End If
End Sub



Public Sub ShowForm(POs As Boolean, RFQs As Boolean, Job As String)
    mJob = Job
    Me.Show vbModal
End Sub


Private Sub Form_Load()
Dim s As String




    Call IniGetForm(Me)
    Call LoadPrinters
    s = IniGet(AppIni, "UserPrinters", HFApp.LoginID, Printer.DeviceName)
    Call SetComboBoxListIndex(cboPrinters, s)
    
    
    Set SMTPMail = New HFMailer.SMTP
    s = HFApp.Options.ValueByName("SendPO_SMTPServer")
    SMTPMail.SMTPHost = Parse(s, 1, ":")
    If Parse(s, 2, ":") <> "" Then SMTPMail.SMTPPort = Parse(s, 2, ":")
    
    Call SetCurrentFrame(0)
    
End Sub

Private Sub LoadPrinters()
On Error GoTo eh
    Dim i As Long
    
    Dim printericon As Long
    Dim defaultPrinterIcon As Long
    
    cboPrinters.ImageList = FMain.SmallIcons
    
    printericon = ImageIndex(FMain.SmallIcons, "printer")
    defaultPrinterIcon = ImageIndex(FMain.SmallIcons, "defaultprinter")
    
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = DefaultPrinterName Then
            Call cboPrinters.AddItemAndData(Printers(i).DeviceName, defaultPrinterIcon, defaultPrinterIcon)
        Else
            Call cboPrinters.AddItemAndData(Printers(i).DeviceName, printericon, printericon)
        End If
    Next
    cboPrinters.ListIndex = 0
eh:
End Sub

Private Sub cboPrinters_Click()
    Dim DeviceName As String
    Dim Status As String
    Dim Model As String
    Dim Location As String
    Dim Comments As String
    Dim Port As String
    Dim IsDefault As Boolean
    Call GetPrinterInfo(cboPrinters.Text, Port, Status, Model, Location, Comments, IsDefault)
    lblPort.Caption = Port
    lblStatus.Caption = Status
    lblModel.Caption = Model
    lblLocation.Caption = Location
    lblComments.Caption = Comments
End Sub
Private Function GetCurrentFrame() As Long
    Dim i As Long
    For i = 0 To WizFrame.UBound
        If WizFrame(i).Visible = True Then
            GetCurrentFrame = i
            Exit Function
        End If
    Next
    GetCurrentFrame = WizFrame.LBound
End Function

Private Sub SetCurrentFrame(RHS As Long, Optional fwd As Boolean = True)
    Dim i As Long
    Dim s As String
    
    Dim rs As Recordset
    Dim AllPOs As Long
    Dim PrintedPOs As Long
    
    
    
    
    If SelectedTask < 2 Then
        'process po's
        lblJobs.Caption = "Select Jobs"
        If RHS = 1 And mJob <> "" Then RHS = 3
        If RHS = 2 And mJob <> "" Then RHS = 0
        For i = 0 To WizFrame.UBound
            WizFrame(i).Visible = i = RHS
        Next
        cmdNav(1).Enabled = RHS > 0
        cmdNav(2).Enabled = RHS < WizFrame.UBound
        cmdNav(3).Enabled = RHS = WizFrame.UBound
    Else
        'process rfps's - skip 1,4,5
        lblJobs.Caption = "Select Jobs/RFP's"
        If RHS = 1 And fwd Then RHS = 2
        If RHS = 1 Then RHS = 0
        If RHS = 4 Then RHS = 6
        If RHS = 5 Then RHS = 3
        For i = 0 To WizFrame.UBound
            WizFrame(i).Visible = i = RHS
        Next
        cmdNav(1).Enabled = RHS > 0
        cmdNav(2).Enabled = RHS < WizFrame.UBound
        cmdNav(3).Enabled = RHS = WizFrame.UBound
    End If
    
    If SelectedTask < 2 Then
        'process PO's
        Select Case RHS
            Case 0:
            Case 1: If fwd Then Call LoadData(RHS, "SELECT DISTINCT Community,CommunityDesc FROM PurchaseOrders p " & POWhereClause(RHS) & " ORDER BY 1")
            Case 2: If fwd Then Call LoadData(RHS, "SELECT DISTINCT Job,JobDesc FROM PurchaseOrders p " & POWhereClause(RHS) & " ORDER BY 1")
            Case 3: If fwd Then Call LoadData(RHS, "SELECT DISTINCT Vendor,VendorDesc FROM PurchaseOrders p " & POWhereClause(RHS) & " ORDER BY 1")
            Case 4: If fwd Then Call LoadData(RHS, "SELECT DISTINCT POGroup,POIndex,case when POIndex=POIndexDescription then '' else POIndexDescription end FROM PurchaseOrders p " & POWhereClause(RHS) & " ORDER BY 1,2")
            Case 5: If fwd Then Call LoadData(RHS, "SELECT DISTINCT PONumber,PODesc + char(9) + DeliveryAddress PODesc FROM PurchaseOrders p " & POWhereClause(RHS) & " ORDER BY 1")
            Case 6:
                On Error Resume Next
                Set rs = HFApp.SqlExec("SELECT COUNT(DISTINCT PONumber), SUM(CASE ISNULL(DeliveryMethod,0) WHEN 0 THEN 1 ELSE 0 END) FROM PurchaseOrders p" & POWhereClause(RHS))
                AllPOs = Val("" & rs(0))
                PrintedPOs = Val("" & rs(1))
                ProgressBar.Max = AllPOs
                
                
                lblFinish.Caption = "Ready for processing. You have selected " & AllPOs & " purchase" & vbCrLf & "orders for sending. Click Finish to send these documents" & vbCrLf & "now."
                lblInfo.Caption = PrintedPOs & " documents are set to print. They will be sent to:"
                
                'stupid thing sometimes doesn't draw itself
                PrinterInfo.Visible = False
                PrinterInfo.Visible = True
                PrinterFrame.Visible = PrintedPOs <> 0
                PrinterInfo.Refresh
                cboPrinters.SetFocus
                PrinterInfo.ZOrder 0
                
                WizFrame(6).Refresh
                WizFrame(6).Visible = False
                WizFrame(6).Visible = True
                
                Me.Refresh
                
                On Error GoTo 0
                            
        End Select
    Else
        'process RFP's
        Select Case RHS
            Case 0:
            Case 1:
            Case 2: If fwd Then Call LoadData(RHS, "SELECT DISTINCT Job,Title,RFP FROM rfps ORDER BY 1,2")
            Case 3: If fwd Then Call LoadData(RHS, "SELECT DISTINCT b.vendor,v.vendor_name FROM vendorbids b join tblvendors v on(b.vendor=v.vendor_id) " & RFPWhereClause(RHS) & " ORDER BY 1")
            Case 4:
            Case 5:
            Case 6:
                On Error Resume Next
                Set rs = HFApp.SqlExec("SELECT COUNT(*), SUM(CASE ISNULL(PurchDelMethod,0) WHEN 0 THEN 1 ELSE 0 END) FROM VendorBids b join tblvendors v on(b.vendor=v.vendor_id) " & RFPWhereClause(RHS))
                AllPOs = Val("" & rs(0))
                PrintedPOs = Val("" & rs(1))
                ProgressBar.Max = AllPOs
                
                
                lblFinish.Caption = "Ready for processing. You have selected " & AllPOs & " documents" & vbCrLf & " for sending. Click Finish to send these documents now."
                lblInfo.Caption = PrintedPOs & " documents are set to print. They will be sent to:"
                
                'stupid thing sometimes doesn't draw itself
                PrinterInfo.Visible = False
                PrinterInfo.Visible = True
                PrinterFrame.Visible = PrintedPOs <> 0
                PrinterInfo.Refresh
                cboPrinters.SetFocus
                PrinterInfo.ZOrder 0
                
                WizFrame(6).Refresh
                WizFrame(6).Visible = False
                WizFrame(6).Visible = True
                
                Me.Refresh
                
                On Error GoTo 0
                            
        End Select
    End If
End Sub


Private Function RFPWhereClause(FrameIndex As Long) As String
    Dim c
    Dim s As String
    Dim r As Long
    
    c = " WHERE b.SentDate" & IIf(SelectedTask = 3, " IS NOT NULL", " IS NULL")
    
    If FrameIndex > 2 Then
        If chkSelectAll(2).Value <> vbChecked Then
            s = ""
            For r = 0 To gData(2).Rows - 1
                If gData(2).Cell(flexcpChecked, r, 0) = flexChecked Then
                    s = s & "," & DbQuote(Num, gData(2).TextMatrix(r, 2))
                End If
            Next
            If s = "" Then
                c = c & " AND 1=2"
            Else
                c = c & " AND b.rfp IN(" & Mid(s, 2) & ")"
            End If
        End If
    End If
    
    If FrameIndex > 3 Then
        If chkSelectAll(3).Value <> vbChecked Then
            s = ""
            For r = 0 To gData(3).Rows - 1
                If gData(3).Cell(flexcpChecked, r, 0) = flexChecked Then
                    s = s & "," & DbQuote(Str, gData(3).TextMatrix(r, 0))
                End If
            Next
            If s = "" Then
                c = c & " AND 1=2"
            Else
                c = c & " AND b.vendor IN(" & Mid(s, 2) & ")"
            End If
        End If
    End If
    
    RFPWhereClause = c
    
End Function



Private Function POWhereClause(FrameIndex As Long) As String
    Dim c
    Dim s As String
    Dim r As Long
    
    c = " WHERE Delivered=" & DbQuote(Bit, SelectedTask = 1)
    
    If mJob <> "" Then
        c = c & " AND job=" & DbQuote(Str, mJob)
    Else
        If FrameIndex > 1 Then
            If chkSelectAll(1).Value <> vbChecked Then
                s = ""
                For r = 0 To gData(1).Rows - 1
                    If gData(1).Cell(flexcpChecked, r, 0) = flexChecked Then
                        s = s & "," & DbQuote(Str, gData(1).TextMatrix(r, 0))
                    End If
                Next
                If s = "" Then
                    c = c & " AND 1=2"
                Else
                    c = c & " AND p.Community IN(" & Mid(s, 2) & ")"
                End If
            End If
        End If
            
        
        If FrameIndex > 2 Then
            If chkSelectAll(2).Value <> vbChecked Then
                s = ""
                For r = 0 To gData(2).Rows - 1
                    If gData(2).Cell(flexcpChecked, r, 0) = flexChecked Then
                        s = s & "," & DbQuote(Str, gData(2).TextMatrix(r, 0))
                    End If
                Next
                If s = "" Then
                    c = c & " AND 1=2"
                Else
                    c = c & " AND p.Job IN(" & Mid(s, 2) & ")"
                End If
            End If
        End If
    End If
    
    If FrameIndex > 3 Then
        If chkSelectAll(3).Value <> vbChecked Then
            s = ""
            For r = 0 To gData(3).Rows - 1
                If gData(3).Cell(flexcpChecked, r, 0) = flexChecked Then
                    s = s & "," & DbQuote(Str, gData(3).TextMatrix(r, 0))
                End If
            Next
            If s = "" Then
                c = c & " AND 1=2"
            Else
                c = c & " AND p.Vendor IN(" & Mid(s, 2) & ")"
            End If
        End If
    End If
    
    If FrameIndex > 4 Then
        If chkSelectAll(4).Value <> vbChecked Then
            s = ""
            For r = 0 To gData(4).Rows - 1
                If gData(4).Cell(flexcpChecked, r, 0) = flexChecked Then
                    s = s & "," & DbQuote(Str, gData(4).TextMatrix(r, 1))
                End If
            Next
            If s = "" Then
                c = c & " AND 1=2"
            Else
                c = c & " AND p.POIndex IN(" & Mid(s, 2) & ")"
            End If
        End If
    End If
    
    If FrameIndex > 5 Then
        If chkSelectAll(5).Value <> vbChecked Then
            s = ""
            For r = 0 To gData(5).Rows - 1
                If gData(5).Cell(flexcpChecked, r, 0) = flexChecked Then
                    s = s & "," & DbQuote(Str, gData(5).TextMatrix(r, 0))
                End If
            Next
            If s = "" Then
                c = c & " AND 1=2"
            Else
                c = c & " AND p.PONumber IN(" & Mid(s, 2) & ")"
            End If
        End If
    End If
    
    POWhereClause = c
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPut(AppIni, "UserPrinters", HFApp.LoginID, cboPrinters.Text)
On Error Resume Next:
    Call Kill(PathAppend(TempPath, "*.*"))
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: Call SetCurrentFrame(GetCurrentFrame - 1, False)
        Case 2: Call SetCurrentFrame(GetCurrentFrame + 1, True)
        Case 3
            If SelectedTask < 2 Then
                If SendPOs Then Unload Me
            Else
                If SendRFPs Then Unload Me
            End If
    End Select
End Sub

Private Sub LoadData(FrameIndex As Long, SQL As String)
    Dim r As Long
    Dim rs As Recordset
    With gData(FrameIndex)
        .Rows = 0
        .Redraw = flexRDNone
        r = -1
        Set rs = HFApp.SqlExec(SQL)
        While Not rs.EOF
            r = r + 1
            If rs.fields.Count = 3 Then
                .AddItem "" & rs(0) & vbTab & rs(1) & vbTab & rs(2)
            Else
                .AddItem "" & rs(0) & vbTab & rs(1)
            End If
            rs.MoveNext
        Wend
        .Cell(flexcpChecked, 0, 0, .Rows - 1, 0) = flexChecked
        chkSelectAll(FrameIndex).Value = vbChecked
        .Redraw = flexRDBuffered
        Call .AutoSize(0, 1)
    End With
End Sub


Private Sub gData_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    bInHere = True
    With gData(Index)
    For i = 0 To .Rows - 1
        If .Cell(flexcpChecked, 0, 0) <> .Cell(flexcpChecked, i, 0) Then
            chkSelectAll(Index).Value = vbGrayed
            bInHere = False
            Exit Sub
        End If
    Next
    chkSelectAll(Index).Value = IIf(.Cell(flexcpChecked, 0, 0) = flexChecked, vbChecked, vbUnchecked)
    End With
    bInHere = False

End Sub

Private Function SendPOs() As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim i As Long
    Dim s As String
    Dim SendVia   As Long
    Dim Address   As String
    Dim PMEmail As String
    Dim RptFormat As String
    Dim PONumbers As String
    
    PrinterInfo.Visible = False
    lblSaving.Visible = True
    ProgressBar.Visible = True
    lblCount.Visible = True
    lblItemDesc.Visible = True
    
    
'    If HFApp.Options(FaxProvider) = "none" Then
'        Set rs = HFApp.SqlExec("SELECT COUNT(*) FROM PurchaseOrders" & POWhereClause(99) & " AND DeliveryMethod=1")
'        If Not rs.EOF Then
'            MsgBox "Some of these purchase orders are to be faxed but no fax provider has been configured. Please contact your system administrator for assistance.", vbExclamation, App.ProductName
'            SendPOs = False
'            Exit Function
'        End If
'    End If
    
    
    If HFApp.Options.ValueByName("SendPO_CCPrjMgr") Then
        s = ""
        s = s & "SELECT p.DeliveryMethod,p.DeliveryAddress,p.POFormat,p.PONumber,p.VendorDesc,pm.email PMEmail" & vbCrLf
        s = s & "FROM PurchaseOrders p" & vbCrLf
        s = s & "left outer join tbljobs j on (p.job=j.job_no)" & vbCrLf
        s = s & "left outer join tblprojectmanager pm on (j.pm=pm.pm)" & vbCrLf
        s = s & POWhereClause(99) & " ORDER BY 1,2,5"
    Else
        s = "SELECT DeliveryMethod,DeliveryAddress,POFormat,PONumber,VendorDesc,'' PMEmail FROM PurchaseOrders p " & POWhereClause(99) & " ORDER BY 1,2,5"
    End If
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        i = i + 1
        lblCount.Caption = "Purchase Order " & ProgressBar.Value + 1 & " of " & ProgressBar.Max
        lblItemDesc.Caption = "" & rs(3) & " -- " & rs(4)
        
        If SendVia <> Val("" & rs(0)) Or Address <> rs(1) Or RptFormat <> rs(2) Or PMEmail <> rs(5) Then
            ProgressBar.Value = i
            Call SendPOGroup(SendVia, Address, PMEmail, RptFormat, Mid(PONumbers, 2))
            SendVia = Val("" & rs(0))
            Address = "" & rs(1)
            RptFormat = "" & rs(2)
            PMEmail = "" & rs(5)
            PONumbers = ""
        End If
        
        PONumbers = PONumbers & "," & DbQuote(Str, "" & rs(3))
        rs.MoveNext
    Wend
    
    Call SendPOGroup(SendVia, Address, PMEmail, RptFormat, Mid(PONumbers, 2))
    SendPOs = True
Exit Function
eh: Call errHandler(SRCFILE & "SendPOs")
End Function

Private Sub SendPOGroup(SendVia As Long, Recipient As String, CcRecipient As String, RptFormat As String, PONumbers As String)
On Error GoTo eh
    
    Dim rs As Recordset
    Dim RptFile  As String
    Dim TmpFile  As String
    Dim Crystal  As New CRAXDRT.Application
    Dim Report   As CRAXDRT.Report
    Dim sections As CRAXDRT.sections
    Dim section  As CRAXDRT.section
    Dim rptObjs  As CRAXDRT.ReportObjects
    Dim rptObj   As Object
    Dim subRpt   As CRAXDRT.Report
    Dim tbl      As CRAXDRT.DatabaseTable
    Dim i        As Long
    Dim Dsn      As String
    Dim ddb      As String
    Dim uid      As String
    Dim pwd      As String
    Dim s           As String
    
    
    Dim Attachments As String
    Dim Subject     As String
    Dim body        As String
    Dim SenderName    As String
    Dim SenderCompany As String
    Dim SenderFax     As String
    Dim SenderPhone   As String
    Dim SenderEmail   As String
    Dim RecipientName    As String
    Dim RecipientCompany As String
    Dim RecipientFax     As String
    Dim RecipientPhone   As String
    Dim RecipientEmail   As String

    txtStatus.Text = ""
    If PONumbers = "" Or RptFormat = "" Then Exit Sub
    
    If SendVia = dtFax Then SendVia = dtPrint
    If Trim(Recipient) = "" Then SendVia = dtPrint
    
    txtStatus.Text = txtStatus.Text & "Loading PO format..." & vbCrLf
    
    RptFile = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", RptFormat & ".rpt")
    Dsn = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "DSN="), 1, ";")
    ddb = HFApp.Databases(dbHomefront).DefaultDatabase
    uid = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "UID="), 1, ";")
    pwd = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "PWD="), 1, ";")
       
    'open the .rpt file
    Set Report = Crystal.OpenReport(RptFile)
    
    txtStatus.Text = txtStatus.Text & "Reading data..." & vbCrLf
    'set connections
    For Each tbl In Report.Database.Tables
        tbl.SetLogOnInfo Dsn, ddb, uid, pwd
        tbl.Location = ddb & ".dbo." & tbl.Location
    Next
    For Each section In Report.sections
        For Each rptObj In section.ReportObjects
            If rptObj.Kind = crSubreportObject Then
                Set subRpt = rptObj.OpenSubreport
                For Each tbl In subRpt.Database.Tables
                    tbl.SetLogOnInfo Dsn, ddb, uid, pwd
                    tbl.Location = ddb & ".dbo." & tbl.Location
                Next
            End If
        Next
    Next

    
    On Error Resume Next
    Call Report.ParameterFields.GetItemByName("PONumber").AddCurrentValue(PONumbers)
    On Error GoTo 0
    
    'get fields data
    s = ""
    s = s & "SELECT pm.PMName SenderName" & vbCrLf
    s = s & "      ,s.CompanyName SenderCompany" & vbCrLf
    s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.fax   ELSE s.fax END SenderFax" & vbCrLf
    s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.phone ELSE s.fax END SenderPhone" & vbCrLf
    s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.email ELSE s.estimator_email END SenderEmail" & vbCrLf
    s = s & "      ,v.PurchContact RecipientName" & vbCrLf
    s = s & "      ,v.Vendor_Name RecipientCompany" & vbCrLf
    s = s & "      ,v.PurchFax RecipientFax" & vbCrLf
    s = s & "      ,v.PurchPhone RecipientPhone" & vbCrLf
    s = s & "      ,v.PurchEmail RecipientEmail" & vbCrLf
    s = s & "      ,p.Job Quote" & vbCrLf
    s = s & "      ,p.Job" & vbCrLf
    s = s & "      ,p.Vendor" & vbCrLf
    s = s & "      ,v.Vendor_Name VendorName" & vbCrLf
    s = s & "      ,j.Model" & vbCrLf
    s = s & "      ,j.Municipal_Address  MunicipalAddress " & vbCrLf
    s = s & "      ,j.Legal_Address  LegalAddress " & vbCrLf
    s = s & "      ,j.County" & vbCrLf
    s = s & "      ,j.Township" & vbCrLf
    s = s & "      ,ISNULL(NULLIF(l.Description,''),j.Community) Community" & vbCrLf
    s = s & "      ,ISNULL(NULLIF(a.Description,''),j.CommunityPhase) Phase" & vbCrLf
    s = s & "  FROM System_setup s" & vbCrLf
    s = s & "      ,POMaster p" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblJobs j on(p.Job=j.Job_no)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblProjectManager pm on(j.Purchaser=pm.PM)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblVendors v on(p.vendor=v.vendor_id)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblLocality l on(j.community=l.area)" & vbCrLf
    s = s & "       LEFT OUTER JOIN CommunityPhase a on(j.CommunityPhase=a.CommunityPhase)" & vbCrLf
    s = s & "WHERE p.PONumber IN(" & PONumbers & ")"
    Set rs = HFApp.SqlExec(s)
    SenderName = "" & rs("SenderName")
    SenderCompany = "" & rs("SenderCompany")
    SenderFax = "" & rs("SenderFax")
    SenderPhone = "" & rs("SenderPhone")
    SenderEmail = "" & rs("SenderEmail")
    RecipientName = "" & rs("RecipientName")
    RecipientCompany = "" & rs("RecipientCompany")
    RecipientFax = "" & rs("RecipientFax")
    RecipientPhone = "" & rs("RecipientPhone")
    RecipientEmail = "" & rs("RecipientEmail")
    Subject = Replace(HFApp.Options(Msg_SendPO_Subject), "<%PONumber%>", Replace(Replace(PONumbers, "','", ", "), "'", ""), , , vbTextCompare)
    body = Replace(HFApp.Options(Msg_SendPO_Body), "<%PONumber%>", PONumbers, , , vbTextCompare)
    For i = 0 To rs.fields.Count - 1
        Subject = Replace(Subject, "<%" & rs.fields(i).Name & "%>", "" & rs(i), , , vbTextCompare)
        body = Replace(body, "<%" & rs.fields(i).Name & "%>", "" & rs(i), , , vbTextCompare)
    Next
    
    
    
    Select Case SendVia
        Case dtPrint
            txtStatus.Text = txtStatus.Text & "Printing..." & vbCrLf
            Call Report.SelectPrinter(lblModel.Caption, cboPrinters.Text, lblLocation.Caption)
            Call Report.PrintOut(False)
            txtStatus.Text = ""
        
        Case dtEmail
            txtStatus.Text = txtStatus.Text & "Retrieving attachments..." & vbCrLf
            'extract attachments
            s = ""
            s = s & "SELECT DISTINCT p.Job,a.FileName,a.Embedded" & vbCrLf
            s = s & "  FROM POMaster p" & vbCrLf
            s = s & "       JOIN Attachments a ON(a.ObjectID='J~'+p.Job)" & vbCrLf
            s = s & "       JOIN POIndexDocuments i ON(p.POIndex=i.POIndex and a.DocumentClass=i.Class)" & vbCrLf
            s = s & " Where p.IncludeDocuments = 1" & vbCrLf
            s = s & "   and p.PONumber in(" & PONumbers & ")" & vbCrLf
            Set rs = HFApp.SqlExec(s)
            While Not rs.EOF
            
                If "" & rs("Embedded") = "True" Then
                    s = PathAppend(TempPath, FileTitle("" & rs("FileName")))
                    If FileExists(s) Then Call Kill(s)
                    Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Attachments WHERE ObjectID=" & DbQuote(Str, "J~" & rs("Job")) & " AND FileName=" & DbQuote(Str, "" & rs("FileName")), "FileImage")
                Else
                    s = "" & rs("FileName")
                End If
                
                Attachments = Attachments & ";" & s
                rs.MoveNext
            Wend
            TmpFile = PathAppend(TempPath, "Purchase Orders.pdf")
            Attachments = Mid(Attachments & ";" & TmpFile, 2)
            If FileExists(TmpFile) Then Call Kill(TmpFile)
            txtStatus.Text = txtStatus.Text & "Writing Purchase Orders.pdf" & vbCrLf
            Report.ExportOptions.DiskFileName = TmpFile
            Report.ExportOptions.DestinationType = crEDTDiskFile
            Report.ExportOptions.FormatType = crEFTPortableDocFormat
            Call Report.Export(False)
            
            
            
            'send it
            If HFApp.Options.ValueByName("SendPO_ViaMapi") <> "False" Then
                Call HFApp.SendMail(False, Recipient, CcRecipient, Subject, body, Attachments)
            Else
                txtStatus.Text = ""
                SMTPMail.ConnectTimeout = 120
                SMTPMail.From = HFApp.Options(Estimator_Email)
                SMTPMail.Recipient = Recipient
                SMTPMail.CcRecipient = CcRecipient
                SMTPMail.Subject = Subject
                SMTPMail.Message = body
                SMTPMail.Attachment = Attachments
                SMTPMail.send
            End If
        
        
        
        Case dtFax
            TmpFile = PathAppend(TempPath, "Purchase Orders.pdf")
            If FileExists(TmpFile) Then Call Kill(TmpFile)
            Report.ExportOptions.DiskFileName = TmpFile
            Report.ExportOptions.DestinationType = crEDTDiskFile
            Report.ExportOptions.FormatType = crEFTPortableDocFormat
            Call Report.Export(False)
            
            'send it -- this also deletes temp documents after processing
            Call HFApp.SendFax(SenderName, SenderCompany, SenderFax, SenderPhone, SenderEmail, _
                               RecipientName, RecipientCompany, RecipientFax, RecipientPhone, RecipientEmail, _
                               Subject, "", TmpFile)
        
    End Select




    'mark sent date
    Call HFApp.SqlExec("UPDATE POMaster SET DeliveryDate=GETDATE() WHERE PONumber IN(" & PONumbers & ")")
Exit Sub
eh: Call errHandler(SRCFILE & "SendPOGroup")
End Sub

Private Sub gData_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).Move 0, WizHead1.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead1.Height - Me.WizFoot.Height
        gData(i).Move gData(i).left, gData(i).Top, WizFrame(i).Width - gData(i).left - 2 * margin, WizFrame(i).Height - gData(i).Top - 2 * margin
        chkSelectAll(i).left = gData(i).left + gData(i).Width - chkSelectAll(i).Width
    Next
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
    
End Sub

Private Sub SMTPMail_SendFailed(Explanation As String)
    txtStatus.Text = txtStatus.Text & Explanation & vbCrLf
    MsgBox "Failed to send PO's" & vbCrLf & vbCrLf & Explanation, vbExclamation, App.ProductName
End Sub

Private Sub SMTPMail_SendSuccesful()
    txtStatus.Text = ""
End Sub

Private Sub SMTPMail_Status(Status As String)
    txtStatus.Text = txtStatus.Text & Replace(Status, Chr(0), "") & vbCrLf
End Sub


Private Property Get SelectedTask() As Long
    Dim i As Long
    For i = 0 To 3
        If chkSend(i).Value Then
            SelectedTask = i
            Exit Property
        End If
    Next
End Property




Private Function SendRFPs() As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim i As Long
    Dim s As String
    
    PrinterInfo.Visible = False
    lblSaving.Visible = True
    ProgressBar.Visible = True
    lblCount.Visible = True
    lblItemDesc.Visible = True
    
    
'    If HFApp.Options(FaxProvider) = "none" Then
'        Set rs = HFApp.SqlExec("SELECT COUNT(*) FROM PurchaseOrders" & POWhereClause(99) & " AND DeliveryMethod=1")
'        If Not rs.EOF Then
'            MsgBox "Some of these purchase orders are to be faxed but no fax provider has been configured. Please contact your system administrator for assistance.", vbExclamation, App.ProductName
'            SendPOs = False
'            Exit Function
'        End If
'    End If
    
    
    s = ""
    s = s & "SELECT PurchDelMethod" & vbCrLf
    s = s & "      ,case purchdelmethod when 1 then purchemail when 2 then purchfax else '' end address" & vbCrLf
    s = s & "      ,v.Vendor_id" & vbCrLf
    s = s & "      ,v.Vendor_Name" & vbCrLf
    s = s & "      ,b.rfp" & vbCrLf
    s = s & "FROM vendorbids b" & vbCrLf
    s = s & "JOIN tblvendors v ON(b.vendor=v.vendor_id)" & vbCrLf
    s = s & RFPWhereClause(99) & vbCrLf
    s = s & "ORDER BY 1,2,3,5" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        i = i + 1
        lblCount.Caption = "RFP " & ProgressBar.Value + 1 & " of " & ProgressBar.Max
        lblItemDesc.Caption = "" & rs(3)
        ProgressBar.Value = i
        Call SendRFPGroup(Val("" & rs(0)), "" & rs(1), Val("" & rs(4)), "" & rs(2))
        rs.MoveNext
    Wend
    
    SendRFPs = True
Exit Function
eh: Call errHandler(SRCFILE & "SendPOs")
End Function

Private Sub SendRFPGroup(SendVia As Long, Recipient As String, RFP As Long, Vendor As String)
On Error GoTo eh
    
    Dim rs As Recordset
    Dim RptFile  As String
    Dim TmpFile  As String
    Dim Crystal  As New CRAXDRT.Application
    Dim Report   As CRAXDRT.Report
    Dim sections As CRAXDRT.sections
    Dim section  As CRAXDRT.section
    Dim rptObjs  As CRAXDRT.ReportObjects
    Dim rptObj   As Object
    Dim subRpt   As CRAXDRT.Report
    Dim tbl      As CRAXDRT.DatabaseTable
    Dim i        As Long
    Dim Dsn      As String
    Dim ddb      As String
    Dim uid      As String
    Dim pwd      As String
    Dim s           As String
    
    
    Dim Attachments As String
    Dim Subject     As String
    Dim body        As String
    Dim SenderName    As String
    Dim SenderCompany As String
    Dim SenderFax     As String
    Dim SenderPhone   As String
    Dim SenderEmail   As String
    Dim RecipientName    As String
    Dim RecipientCompany As String
    Dim RecipientFax     As String
    Dim RecipientPhone   As String
    Dim RecipientEmail   As String

    txtStatus.Text = ""
    If RFP = 0 Or Vendor = "" Then Exit Sub
    
    If SendVia = dtFax Then SendVia = dtPrint
    If Trim(Recipient) = "" Then SendVia = dtPrint
    
    txtStatus.Text = txtStatus.Text & "Reading RFP..." & vbCrLf
    
    RptFile = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\RFQ.rpt")
    Dsn = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "DSN="), 1, ";")
    ddb = HFApp.Databases(dbHomefront).DefaultDatabase
    uid = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "UID="), 1, ";")
    pwd = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "PWD="), 1, ";")
       
    'open the .rpt file
    Set Report = Crystal.OpenReport(RptFile)
    
    txtStatus.Text = txtStatus.Text & "Reading data..." & vbCrLf
    'set connections
    For Each tbl In Report.Database.Tables
        tbl.SetLogOnInfo Dsn, ddb, uid, pwd
        tbl.Location = ddb & ".dbo." & tbl.Location
    Next
    For Each section In Report.sections
        For Each rptObj In section.ReportObjects
            If rptObj.Kind = crSubreportObject Then
                Set subRpt = rptObj.OpenSubreport
                For Each tbl In subRpt.Database.Tables
                    tbl.SetLogOnInfo Dsn, ddb, uid, pwd
                    tbl.Location = ddb & ".dbo." & tbl.Location
                Next
            End If
        Next
    Next

    
    On Error Resume Next
    Call Report.ParameterFields.GetItemByName("RFP").AddCurrentValue(RFP)
    Call Report.ParameterFields.GetItemByName("Vendor").AddCurrentValue(Vendor)
    On Error GoTo 0
    
    'get fields data
    s = ""
    s = s & "SELECT CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.pmname ELSE s.companyname END SenderName" & vbCrLf
    s = s & "      ,s.CompanyName SenderCompany" & vbCrLf
    s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.fax   ELSE s.fax END SenderFax" & vbCrLf
    s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.phone ELSE s.fax END SenderPhone" & vbCrLf
    s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.email ELSE s.estimator_email END SenderEmail" & vbCrLf
    s = s & "      ,v.PurchContact RecipientName" & vbCrLf
    s = s & "      ,v.Vendor_Name RecipientCompany" & vbCrLf
    s = s & "      ,v.PurchFax RecipientFax" & vbCrLf
    s = s & "      ,v.PurchPhone RecipientPhone" & vbCrLf
    s = s & "      ,v.PurchEmail RecipientEmail" & vbCrLf
    s = s & "      ,p.Job Quote" & vbCrLf
    s = s & "      ,p.Job" & vbCrLf
    s = s & "      ,p.RFP " & vbCrLf
    s = s & "      ,b.Vendor" & vbCrLf
    s = s & "      ,b.BidGUID" & vbCrLf
    s = s & "      ,v.Vendor_Name VendorName" & vbCrLf
    s = s & "      ,j.Model" & vbCrLf
    s = s & "      ,j.Municipal_Address  MunicipalAddress " & vbCrLf
    s = s & "      ,j.Legal_Address  LegalAddress " & vbCrLf
    s = s & "      ,j.County" & vbCrLf
    s = s & "      ,j.Township" & vbCrLf
    s = s & "      ,ISNULL(NULLIF(l.Description,''),j.Community) Community" & vbCrLf
    s = s & "      ,ISNULL(NULLIF(a.Description,''),j.CommunityPhase) Phase" & vbCrLf
    s = s & "  FROM System_setup s" & vbCrLf
    s = s & "      ,RFPs p" & vbCrLf
    s = s & "       JOIN VendorBids b on(p.rfp=b.rfp)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblJobs j on(p.Job=j.Job_no)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblProjectManager pm on(j.Purchaser=pm.PM)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblVendors v on(b.vendor=v.vendor_id)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblLocality l on(j.community=l.area)" & vbCrLf
    s = s & "       LEFT OUTER JOIN CommunityPhase a on(j.CommunityPhase=a.CommunityPhase)" & vbCrLf
    s = s & "WHERE p.RFP =" & DbQuote(Str, RFP)
    s = s & "  and b.vendor=" & DbQuote(Str, Vendor)
    Set rs = HFApp.SqlExec(s)
    SenderName = "" & rs("SenderName")
    SenderCompany = "" & rs("SenderCompany")
    SenderFax = "" & rs("SenderFax")
    SenderPhone = "" & rs("SenderPhone")
    SenderEmail = "" & rs("SenderEmail")
    RecipientName = "" & rs("RecipientName")
    RecipientCompany = "" & rs("RecipientCompany")
    RecipientFax = "" & rs("RecipientFax")
    RecipientPhone = "" & rs("RecipientPhone")
    RecipientEmail = "" & rs("RecipientEmail")
    
    Subject = HFApp.Options.ValueByName("Msg_RFP_Subject")
    body = HFApp.Options.ValueByName("Msg_RFP_Body")
    For i = 0 To rs.fields.Count - 1
        Subject = Replace(Subject, "<%" & rs.fields(i).Name & "%>", "" & rs(i), , , vbTextCompare)
        body = Replace(body, "<%" & rs.fields(i).Name & "%>", "" & rs(i), , , vbTextCompare)
    Next
    
    
    
    Select Case SendVia
        Case dtPrint
            txtStatus.Text = txtStatus.Text & "Printing..." & vbCrLf
            Call Report.SelectPrinter(lblModel.Caption, cboPrinters.Text, lblLocation.Caption)
            Call Report.PrintOut(False)
            txtStatus.Text = ""
        
        Case dtEmail
            txtStatus.Text = txtStatus.Text & "Retrieving attachments..." & vbCrLf
            'extract attachments
            s = "SELECT DISTINCT FileName,Embedded,FileTitle FROM Attachments Where ObjectID=" & DbQuote(Str, "RFP~" & RFP)
            Set rs = HFApp.SqlExec(s)
            While Not rs.EOF
            
                If "" & rs("Embedded") = "True" Then
                    s = PathAppend(TempPath, FileTitle("" & rs("FileName")))
                    If FileExists(s) Then Call Kill(s)
                    Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Attachments WHERE ObjectID=" & DbQuote(Str, "RFP~" & RFP) & " AND FileName=" & DbQuote(Str, "" & rs("FileName")), "FileImage")
                Else
                    s = "" & rs("FileName")
                End If
                
                Attachments = Attachments & ";" & s
                rs.MoveNext
            Wend
            
            
            'attach RFP
            If HFApp.Options.ValueByName("AttachRFP_PDF") = "True" Then
                TmpFile = PathAppend(TempPath, "RFP.pdf")
                Attachments = Mid(Attachments & ";" & TmpFile, 2)
                If FileExists(TmpFile) Then Call Kill(TmpFile)
                txtStatus.Text = txtStatus.Text & "Writing RFP.pdf" & vbCrLf
                Report.ExportOptions.DiskFileName = TmpFile
                Report.ExportOptions.DestinationType = crEDTDiskFile
                Report.ExportOptions.FormatType = crEFTPortableDocFormat
                Call Report.Export(False)
                
            ElseIf HFApp.Options.ValueByName("AttachRFP_XLS") = "True" Then
                TmpFile = PathAppend(TempPath, "RFP.xls")
                Attachments = Mid(Attachments & ";" & TmpFile, 2)
                If FileExists(TmpFile) Then Call Kill(TmpFile)
                txtStatus.Text = txtStatus.Text & "Writing RFP.pdf" & vbCrLf
                Report.ExportOptions.DiskFileName = TmpFile
                Report.ExportOptions.DestinationType = crEDTDiskFile
                Report.ExportOptions.FormatType = crEFTExcel97
                Report.ExportOptions.ExcelUseWorksheetFunctions = True
                Report.ExportOptions.ExcelUseConstantColumnWidth = False
                Call Report.Export(False)
            End If
            
            'send it
            If HFApp.Options.ValueByName("SendPO_ViaMapi") <> "False" Then
                Call HFApp.SendMail(False, Recipient, "", Subject, body, Attachments)
            Else
                txtStatus.Text = ""
                SMTPMail.ConnectTimeout = 120
                SMTPMail.From = HFApp.Options(Estimator_Email)
                SMTPMail.Recipient = Recipient
                SMTPMail.CcRecipient = ""
                SMTPMail.Subject = Subject
                SMTPMail.Message = body
                SMTPMail.Attachment = Attachments
                SMTPMail.send
            End If
        
        
        Case dtFax
            TmpFile = PathAppend(TempPath, "RFP.pdf")
            If FileExists(TmpFile) Then Call Kill(TmpFile)
            Report.ExportOptions.DiskFileName = TmpFile
            Report.ExportOptions.DestinationType = crEDTDiskFile
            Report.ExportOptions.FormatType = crEFTPortableDocFormat
            Call Report.Export(False)
            
            'send it -- this also deletes temp documents after processing
            Call HFApp.SendFax(SenderName, SenderCompany, SenderFax, SenderPhone, SenderEmail, _
                               RecipientName, RecipientCompany, RecipientFax, RecipientPhone, RecipientEmail, _
                               Subject, "", TmpFile)
        
    End Select




    'mark sent date
    Call HFApp.SqlExec("UPDATE vendorbids SET SentDate=GETDATE() WHERE RFP=" & DbQuote(Num, RFP) & " AND Vendor=" & DbQuote(Str, Vendor))
Exit Sub
eh: Call errHandler(SRCFILE & "SendPOGroup")
End Sub



