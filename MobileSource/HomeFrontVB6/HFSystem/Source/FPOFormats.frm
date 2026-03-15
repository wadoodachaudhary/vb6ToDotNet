VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FPOFormats 
   Caption         =   "Purchase Order Setup"
   ClientHeight    =   5940
   ClientLeft      =   1215
   ClientTop       =   1845
   ClientWidth     =   11130
   Icon            =   "FPOFormats.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   11130
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   9960
      TabIndex        =   2
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   8820
      TabIndex        =   1
      Top             =   5550
      Width           =   1035
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10845
      _cx             =   1999129273
      _cy             =   1999119484
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
      BackColorFixed  =   -2147483643
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   8
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   20
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FPOFormats.frx":058A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   6
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   2
      ExplorerBar     =   0
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
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   330
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   3
      Top             =   5520
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   582
      ButtonWidth     =   1746
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New PO"
            Key             =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit PO"
            Key             =   "Edit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Items"
            Key             =   "Items"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vendors"
            Key             =   "vendor"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Key             =   "preview"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallIcons 
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
         NumListImages   =   73
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":087B
            Key             =   "option"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":0E15
            Key             =   "RFP"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":13AF
            Key             =   "quote"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":1949
            Key             =   "sendreceive"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":1EE3
            Key             =   "communitystandards"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":247D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":2D57
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":3631
            Key             =   "pricelists"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":3F0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":47E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":50BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":5999
            Key             =   "MB"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":5F33
            Key             =   "QB"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":64CD
            Key             =   "custom"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":6A67
            Key             =   "assembly"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":7001
            Key             =   "links"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":759B
            Key             =   "ItemDB"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":7B35
            Key             =   "sendpos"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":80CF
            Key             =   "HelpSearch"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":8669
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":8C03
            Key             =   "New"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":919D
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":9737
            Key             =   "HelpContents"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":9CD1
            Key             =   "EditAssembly"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":A26B
            Key             =   "Forecast"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":A805
            Key             =   "shrink"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":AD9F
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":B339
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":B8D3
            Key             =   "close"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":BE6D
            Key             =   "expand"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":C407
            Key             =   "SaveAs"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":C9A1
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":CF3B
            Key             =   "RePrice"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":D4D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":DA6F
            Key             =   "estimating"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":E009
            Key             =   "jobcost"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":E5A3
            Key             =   "error"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":EB3D
            Key             =   "salesworksheet"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":F0D7
            Key             =   "ExcelExport"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":F671
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":FC0B
            Key             =   "newworksheet"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":101A5
            Key             =   "worksheet"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":1073F
            Key             =   "purchaseorder"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":10CD9
            Key             =   "ExcelImport"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":11273
            Key             =   "groupphase"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":1180D
            Key             =   "costcode"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":11DA7
            Key             =   "job"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":12341
            Key             =   "information"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":128DB
            Key             =   "item"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":12E75
            Key             =   "itemchecked"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":1340F
            Key             =   "phase"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":139A9
            Key             =   "vendor"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":13F43
            Key             =   "warning"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":144DD
            Key             =   "question"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":14A77
            Key             =   "category"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":15011
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":155AB
            Key             =   "AddItems"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":15B45
            Key             =   "model"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":160DF
            Key             =   "area"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":16679
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":16C13
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":16D6D
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":16EC7
            Key             =   "AlignCenter"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":17021
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":1717B
            Key             =   "AlignLeft"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":172D5
            Key             =   "Bullet"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":1742F
            Key             =   "BulletNumber"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":17589
            Key             =   "AlignRight"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":176E3
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":17C7D
            Key             =   "defaultprinter"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":18217
            Key             =   "costcodes"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":18AF1
            Key             =   "defaultvendors"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPOFormats.frx":193CB
            Key             =   "editpos"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FPOFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FPOFormats"
Private mDirty As Boolean

    

Private Sub Form_Load()
    Call SetToolbarIcons(Toolbar, SmallIcons)
    Call IniGetForm(Me)
    Toolbar.Visible = True
    Call LoadFormats
    Call LoadData
    Toolbar.Visible = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    gData.Move margin, margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - 3 * margin - cmdNav(0).Height
    Toolbar.Move margin, Me.ScaleHeight - Toolbar.Height - margin
    cmdNav(0).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
    cmdNav(1).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
End Sub

Private Sub LoadFormats()
    Dim s As String
    Dim list As String
    s = Dir(PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", "*.rpt"), , True)
    While s <> ""
        list = list & "|" & StripExtension(s)
        s = Dir
    Wend
    gData.ColComboList(gData.ColIndex("POFormat")) = list
End Sub

Private Sub LoadData()
    Dim rs As Recordset
    Dim s As String
    Dim r As Long
    With gData
        Screen.MousePointer = vbHourglass
        .Redraw = flexRDNone
        .Rows = 0
        
        .AddItem "": r = .Rows - 1
        .TextMatrix(r, .ColIndex("Heading")) = "Item"
        .TextMatrix(r, .ColIndex("POFormat")) = "Format"
        .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = True
        .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbHighlight
        
        .AddItem "": r = .Rows - 1
        .TextMatrix(r, .ColIndex("POIndex")) = "" & "Global Default"
        .TextMatrix(r, .ColIndex("POFormat")) = "" & HFApp.Options(POFormat)
        .RowData(r) = "GLOBAL"
        
        
        .AddItem "": r = .Rows - 1
        .TextMatrix(r, .ColIndex("Heading")) = "Purchase Orders"
        .TextMatrix(r, .ColIndex("POGroup")) = "Group"
        .TextMatrix(r, .ColIndex("HideQty")) = "#"
        .TextMatrix(r, .ColIndex("HidePrice")) = "$"
        .TextMatrix(r, .ColIndex("TotalOnly")) = "!"
        .TextMatrix(r, .ColIndex("POFormat")) = "Format"
        
        .TextMatrix(r, .ColIndex("ReleaseTaskDesc")) = "Release After"
        
        .TextMatrix(r, .ColIndex("PayPoint1Percent")) = "Pmt1"
        .TextMatrix(r, .ColIndex("PayPoint2Percent")) = "Pmt2"
        .TextMatrix(r, .ColIndex("PayPoint3Percent")) = "Pmt3"
        .TextMatrix(r, .ColIndex("PayPoint4Percent")) = "Pmt4"
        .TextMatrix(r, .ColIndex("PayPoint5Percent")) = "Pmt5"
        
        .TextMatrix(r, .ColIndex("PayPoint1Task")) = "Pmt1 Task"
        .TextMatrix(r, .ColIndex("PayPoint2Task")) = "Pmt2 Task"
        .TextMatrix(r, .ColIndex("PayPoint3Task")) = "Pmt3 Task"
        .TextMatrix(r, .ColIndex("PayPoint4Task")) = "Pmt4 Task"
        .TextMatrix(r, .ColIndex("PayPoint5Task")) = "Pmt5 Task"
        
        
        .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = True
        .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbHighlight
        
        .FrozenRows = 3
        
        
        Call HFApp.SqlExec("DELETE FROM tblPOIndex WHERE DivisionID = " & HFApp.DivisionID & " and ISNULL(POIndex,'')=''")
        
        
        s = ""
        s = s & "SELECT" & vbCrLf
        s = s & " p.*" & vbCrLf
        s = s & ",rt.Description ReleaseTaskDesc" & vbCrLf
        s = s & ",p1.Description PayPoint1Desc" & vbCrLf
        s = s & ",p2.Description PayPoint2Desc" & vbCrLf
        s = s & ",p3.Description PayPoint3Desc" & vbCrLf
        s = s & ",p4.Description PayPoint4Desc" & vbCrLf
        s = s & ",p5.Description PayPoint5Desc" & vbCrLf
        s = s & "FROM tblPOIndex p" & vbCrLf
        s = s & "LEFT OUTER JOIN LibraryTasks rt ON(p.releasetaskid=rt.taskid)" & vbCrLf
        s = s & "LEFT OUTER JOIN LibraryTasks p1 ON(p.PayPoint1SchedTask=p1.taskid)" & vbCrLf
        s = s & "LEFT OUTER JOIN LibraryTasks p2 ON(p.PayPoint2SchedTask=p2.taskid)" & vbCrLf
        s = s & "LEFT OUTER JOIN LibraryTasks p3 ON(p.PayPoint3SchedTask=p3.taskid)" & vbCrLf
        s = s & "LEFT OUTER JOIN LibraryTasks p4 ON(p.PayPoint4SchedTask=p4.taskid)" & vbCrLf
        s = s & "LEFT OUTER JOIN LibraryTasks p5 ON(p.PayPoint5SchedTask=p5.taskid)" & vbCrLf
        s = s & "where p.DivisionID = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "ORDER BY p.POIndex" & vbCrLf
        
        Set rs = HFApp.SqlExec(s)
        
        While Not rs.EOF
            .AddItem "": r = .Rows - 1
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("POGroup")) = "" & rs("POGroup")
            .Cell(flexcpChecked, r, .ColIndex("HideQty")) = IIf("" & rs("HideQty") = "True", flexChecked, flexUnchecked)
            .Cell(flexcpChecked, r, .ColIndex("HidePrice")) = IIf("" & rs("HidePrice") = "True", flexChecked, flexUnchecked)
            .Cell(flexcpChecked, r, .ColIndex("TotalOnly")) = IIf("" & rs("TotalOnly") = "True", flexChecked, flexUnchecked)
            .TextMatrix(r, .ColIndex("POFormat")) = "" & rs("POFormat")
            .TextMatrix(r, .ColIndex("ReleaseTaskID")) = "" & rs("ReleaseTaskID")
            .TextMatrix(r, .ColIndex("ReleaseTaskDesc")) = "" & rs("ReleaseTaskDesc")
            
            .TextMatrix(r, .ColIndex("PayPoint1Percent")) = IIf(Val("" & rs("PayPoint1Percent")) = 0, "", Val("" & rs("PayPoint1Percent")))
            .TextMatrix(r, .ColIndex("PayPoint2Percent")) = IIf(Val("" & rs("PayPoint2Percent")) = 0, "", Val("" & rs("PayPoint2Percent")))
            .TextMatrix(r, .ColIndex("PayPoint3Percent")) = IIf(Val("" & rs("PayPoint3Percent")) = 0, "", Val("" & rs("PayPoint3Percent")))
            .TextMatrix(r, .ColIndex("PayPoint4Percent")) = IIf(Val("" & rs("PayPoint4Percent")) = 0, "", Val("" & rs("PayPoint4Percent")))
            .TextMatrix(r, .ColIndex("PayPoint5Percent")) = IIf(Val("" & rs("PayPoint5Percent")) = 0, "", Val("" & rs("PayPoint5Percent")))
            
            
            .TextMatrix(r, .ColIndex("PayPoint1Task")) = "" & rs("PayPoint1Desc")
            .TextMatrix(r, .ColIndex("PayPoint2Task")) = "" & rs("PayPoint2Desc")
            .TextMatrix(r, .ColIndex("PayPoint3Task")) = "" & rs("PayPoint3Desc")
            .TextMatrix(r, .ColIndex("PayPoint4Task")) = "" & rs("PayPoint4Desc")
            .TextMatrix(r, .ColIndex("PayPoint5Task")) = "" & rs("PayPoint5Desc")
            .Cell(flexcpData, r, .ColIndex("PayPoint1Task")) = "K" & rs("PayPoint1SchedTask")
            .Cell(flexcpData, r, .ColIndex("PayPoint2Task")) = "K" & rs("PayPoint2SchedTask")
            .Cell(flexcpData, r, .ColIndex("PayPoint3Task")) = "K" & rs("PayPoint3SchedTask")
            .Cell(flexcpData, r, .ColIndex("PayPoint4Task")) = "K" & rs("PayPoint4SchedTask")
            .Cell(flexcpData, r, .ColIndex("PayPoint5Task")) = "K" & rs("PayPoint5SchedTask")
            
            
            .RowData(r) = "POINDEX"
            rs.MoveNext
        Wend
        
        .ColWidth(0) = 240
        Call .AutoSize(1, .Cols - 1)
        
        mDirty = False
        .Redraw = flexRDBuffered
        Screen.MousePointer = vbDefault
    End With
End Sub


Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    gData.ColWidth(0) = 240
    Call gData.AutoSize(1, gData.Cols - 1)
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    Dim r As Long
    
    With gData
        .AutoSearch = flexSearchNone
        Cancel = False
        Select Case True
            Case .RowData(Row) = ""
                Cancel = True
                
            Case .ColKey(Col) = "Heading"
                Cancel = True
                
            Case .ColKey(Col) = "POIndex"
                Cancel = True
                .AutoSearch = flexSearchFromCursor
            
            Case .RowData(Row) = "GLOBAL" And .ColKey(Col) <> "POFormat"
                Cancel = True
            
            Case .ColKey(Col) = "PayPoint5Percent":                Cancel = True
            Case .ColKey(Col) = "PayPoint4Percent":                Cancel = .TextMatrix(Row, .ColIndex("PayPoint3Percent")) = ""
            Case .ColKey(Col) = "PayPoint3Percent":                Cancel = .TextMatrix(Row, .ColIndex("PayPoint2Percent")) = ""
            Case .ColKey(Col) = "PayPoint2Percent":                Cancel = .TextMatrix(Row, .ColIndex("PayPoint1Percent")) = ""
            Case .ColKey(Col) = "PayPoint1Percent":
            
            Case .ColKey(Col) = "POGroup"
                s = ""
                For r = 0 To .Rows - 1
                    If Trim(.TextMatrix(r, .ColIndex("POGroup"))) <> "" And .RowData(r) <> "" And InStr(1, s, .TextMatrix(r, .ColIndex("POGroup"))) = 0 Then
                        s = s & "|" & .TextMatrix(r, .ColIndex("POGroup"))
                    End If
                Next
                .ColComboList(Col) = s
        End Select
        gData.EditMaxLength = IIf(gData.ColKey(Col) = "POGroup", 20, 0)
    End With
End Sub
Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        
        If .ColKey(Col) Like "PayPoint?Percent" Then
            .EditText = Abs(Val(.EditText))
        End If
        
        Select Case True
            Case .ColKey(Col) = "PayPoint4Percent"
                .TextMatrix(Row, .ColIndex("PayPoint5Percent")) = 100 - (.ValueMatrix(Row, .ColIndex("PayPoint1Percent")) + .ValueMatrix(Row, .ColIndex("PayPoint2Percent")) + .ValueMatrix(Row, .ColIndex("PayPoint3Percent")) + Val(.EditText))
                
            Case .ColKey(Col) = "PayPoint3Percent"
                .TextMatrix(Row, .ColIndex("PayPoint4Percent")) = 100 - (.ValueMatrix(Row, .ColIndex("PayPoint1Percent")) + .ValueMatrix(Row, .ColIndex("PayPoint2Percent")) + Val(.EditText))
                .TextMatrix(Row, .ColIndex("PayPoint5Percent")) = 0
            
            Case .ColKey(Col) = "PayPoint2Percent"
                .TextMatrix(Row, .ColIndex("PayPoint3Percent")) = 100 - (.ValueMatrix(Row, .ColIndex("PayPoint1Percent")) + Val(.EditText))
                .TextMatrix(Row, .ColIndex("PayPoint4Percent")) = 0
                .TextMatrix(Row, .ColIndex("PayPoint5Percent")) = 0
            
            Case .ColKey(Col) = "PayPoint1Percent"
                .EditText = Min(Val(.EditText), 100)
                .TextMatrix(Row, .ColIndex("PayPoint2Percent")) = 100 - Val(.EditText)
                .TextMatrix(Row, .ColIndex("PayPoint3Percent")) = 0
                .TextMatrix(Row, .ColIndex("PayPoint4Percent")) = 0
                .TextMatrix(Row, .ColIndex("PayPoint5Percent")) = 0
                
        End Select
        
        If .ValueMatrix(Row, .ColIndex("PayPoint2Percent")) = 0 Then .TextMatrix(Row, .ColIndex("PayPoint2Percent")) = ""
        If .ValueMatrix(Row, .ColIndex("PayPoint3Percent")) = 0 Then .TextMatrix(Row, .ColIndex("PayPoint3Percent")) = ""
        If .ValueMatrix(Row, .ColIndex("PayPoint4Percent")) = 0 Then .TextMatrix(Row, .ColIndex("PayPoint4Percent")) = ""
        If .ValueMatrix(Row, .ColIndex("PayPoint5Percent")) = 0 Then .TextMatrix(Row, .ColIndex("PayPoint5Percent")) = ""
        
        
        
        If Not Cancel Then
            mDirty = True
            .Cell(flexcpData, .Row, 0, .RowSel, 0) = "DIRTY"
        End If
        
    End With
End Sub


Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With gData
    Select Case True
    
        Case .ColKey(Col) Like "PayPoint?Task"
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Scheduling Task", "SELECT TaskID,Description FROM LibraryTasks", Mid(.Cell(flexcpData, Row, Col), 2), , , , "TaskID") Then
                .Cell(flexcpData, Row, Col) = "K" & FPickList.SelectedItem("TaskID")
                .TextMatrix(Row, Col) = FPickList.SelectedItem("Description")
                .Cell(flexcpData, .Row, 0, .RowSel, 0) = "DIRTY"
                mDirty = True
                .ColWidth(0) = 240
                Call .AutoSize(1, .Cols - 1)
            End If
    
        Case .ColKey(Col) = "ReleaseTaskDesc"
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Scheduling Task", "SELECT TaskID,Description FROM LibraryTasks", .TextMatrix(Row, .ColIndex("ReleaseTaskID")), , , , "TaskID") Then
                .TextMatrix(Row, .ColIndex("ReleaseTaskID")) = FPickList.SelectedItem("TaskID")
                .TextMatrix(Row, .ColIndex("ReleaseTaskDesc")) = FPickList.SelectedItem("Description")
                .Cell(flexcpData, .Row, 0, .RowSel, 0) = "DIRTY"
                mDirty = True
                .ColWidth(0) = 240
                Call .AutoSize(1, .Cols - 1)
            End If
    End Select
    End With
End Sub

Private Sub gData_DblClick()
    Call Toolbar_ButtonClick(Toolbar.Buttons("Edit"))
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    With gData
    Select Case True
        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.ShowForm(gData)
            
        Case KeyCode = vbKeyDelete And Shift = 0 And .ColKey(.Col) Like "PayPoint?Task" And .Row > 2 And .RowSel > 2
            .Cell(flexcpText, .Row, .Col, .RowSel, .Col) = ""
            .Cell(flexcpData, .Row, .Col, .RowSel, .Col) = ""
            .Cell(flexcpData, .Row, 0, .RowSel, 0) = "DIRTY"
            mDirty = True
            .ColWidth(0) = 240
            Call .AutoSize(1, .Cols - 1)
            
        Case KeyCode = vbKeyDelete And Shift = 0 And .ColKey(.Col) = "ReleaseTaskDesc" And .Row > 2 And .RowSel > 2
            .Cell(flexcpText, .Row, .Col, .RowSel, .Col) = ""
            .Cell(flexcpText, .Row, .ColIndex("ReleaseTaskID"), .RowSel, .ColIndex("ReleaseTaskID")) = ""
            .Cell(flexcpData, .Row, 0, .RowSel, 0) = "DIRTY"
            mDirty = True
            .ColWidth(0) = 240
            Call .AutoSize(1, .Cols - 1)
                    
       
        Case KeyCode = vbKeyDelete And Shift = 0
            If IsIn(.ColKey(.Col), "Description", "POGroup", "POFormat") And .Row > 2 And .RowSel > 2 Then
                .Cell(flexcpText, .Row, .Col, .RowSel, .Col) = ""
                .Cell(flexcpData, .Row, 0, .RowSel, 0) = "DIRTY"
                mDirty = True
            End If
            .ColWidth(0) = 240
            Call .AutoSize(1, .Cols - 1)
                    
    End Select
    End With
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
    
    
    Screen.MousePointer = vbHourglass
    Dim r As Long
    Dim s As String
    With gData
    For r = 0 To .Rows - 1
        If .Cell(flexcpData, r, 0) = "DIRTY" Then
            Select Case .RowData(r)
            
                Case "GLOBAL"
                    HFApp.Options.Value(POFormat) = .TextMatrix(r, .ColIndex("POFormat"))
                    HFApp.Options.SaveData
                    
                Case "POINDEX"
                    s = ""
                    s = s & "UPDATE tblPOIndex" & vbCrLf
                    s = s & "   SET POFormat=" & DbQuote(Str, .TextMatrix(r, .ColIndex("POFormat")), , True) & vbCrLf
                    s = s & "      ,Description=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                    s = s & "      ,POGroup=" & DbQuote(Str, .TextMatrix(r, .ColIndex("POGroup"))) & vbCrLf
                    s = s & "      ,HideQty=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("HideQty")) = flexChecked) & vbCrLf
                    s = s & "      ,HidePrice=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("HidePrice")) = flexChecked) & vbCrLf
                    s = s & "      ,TotalOnly=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("TotalOnly")) = flexChecked) & vbCrLf
                    s = s & "      ,ReleaseTaskID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ReleaseTaskID"))) & vbCrLf
                    s = s & "      ,PayPoint1Percent=" & DbQuote(Num, .TextMatrix(r, .ColIndex("PayPoint1Percent"))) & vbCrLf
                    s = s & "      ,PayPoint2Percent=" & DbQuote(Num, .TextMatrix(r, .ColIndex("PayPoint2Percent"))) & vbCrLf
                    s = s & "      ,PayPoint3Percent=" & DbQuote(Num, .TextMatrix(r, .ColIndex("PayPoint3Percent"))) & vbCrLf
                    s = s & "      ,PayPoint4Percent=" & DbQuote(Num, .TextMatrix(r, .ColIndex("PayPoint4Percent"))) & vbCrLf
                    s = s & "      ,PayPoint5Percent=" & DbQuote(Num, .TextMatrix(r, .ColIndex("PayPoint5Percent"))) & vbCrLf
                    
                    
                    s = s & "      ,PayPoint1SchedTask=" & DbQuote(Num, Mid(.Cell(flexcpData, r, .ColIndex("PayPoint1Task")), 2)) & vbCrLf
                    s = s & "      ,PayPoint2SchedTask=" & DbQuote(Num, Mid(.Cell(flexcpData, r, .ColIndex("PayPoint2Task")), 2)) & vbCrLf
                    s = s & "      ,PayPoint3SchedTask=" & DbQuote(Num, Mid(.Cell(flexcpData, r, .ColIndex("PayPoint3Task")), 2)) & vbCrLf
                    s = s & "      ,PayPoint4SchedTask=" & DbQuote(Num, Mid(.Cell(flexcpData, r, .ColIndex("PayPoint4Task")), 2)) & vbCrLf
                    s = s & "      ,PayPoint5SchedTask=" & DbQuote(Num, Mid(.Cell(flexcpData, r, .ColIndex("PayPoint5Task")), 2)) & vbCrLf
                    
                    s = s & "WHERE DivisionID = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                    s = s & " and POIndex=" & DbQuote(Str, Trim(.TextMatrix(r, .ColIndex("POIndex")))) & vbCrLf
                    Call HFApp.SqlExec(s)
                    
                    s = ""
                    s = s & "insert into DivisionPOIndexes(POIndex,DivisionID,POFormat)" & vbCrLf
                    s = s & "select p.POIndex,p.DivisionID,p.POFormat from tblPOindex p" & vbCrLf
                    's = s & "join Divisions d on 1 = 1" & vbCrLf
                    s = s & "Left Outer join DivisionPoIndexes dp on (dp.POIndex = p.POIndex and dp.DivisionID = p.DivisionID)" & vbCrLf
                    s = s & "where dp.DivisionID is null and p.DivisionID = " & HFApp.DivisionID & vbCrLf
                    Call HFApp.SqlExec(s, dbHomefront)
                    
            End Select
            .Cell(flexcpData, r, 0) = ""
        End If
    Next
    End With
    Screen.MousePointer = vbDefault
    mDirty = False
    SaveData = True
    
Exit Function
eh: Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function



Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0:  If SaveData(False) Then Unload Me
        Case 1:  Unload Me
    End Select
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim POIndex As String
    Dim PONumber As String
    Dim s As String
    'Dim v As New HFPrinter.ReportViewer
    With gData
        If .Row > 2 Then POIndex = .TextMatrix(.Row, .ColIndex("POIndex"))
        Select Case Button.Key
            
            Case "New"
                If Not SaveData(True) Then Exit Sub
                On Error Resume Next
                Call HFApp.SqlExec("INSERT INTO tblPOIndex(DivisionID,POIndex,POType) VALUES(" & HFApp.DivisionID & ",'(untitled)','Purchase Order')")
                On Error GoTo eh
                If FPOIndex.Edit("(untitled)") Then
                    Call LoadData
                Else
                    Call HFApp.SqlExec("DELETE FROM tblPOIndex WHERE DivisionID = " & HFApp.DivisionID & " and POIndex='(untitled)'")
                End If
            
            Case "Edit"
                If POIndex <> "" Then
                    If Not SaveData(True) Then Exit Sub
                    If FPOIndex.Edit(POIndex) Then
                        Call LoadData
                    End If
                End If
                
            Case "Delete"
                If POIndex <> "" Then
                    If vbNo = MsgBox("If you delete this purchase order it will be removed immediately" & vbCrLf & _
                                     "and permanently. Are you sure this is what you want to do?" & vbCrLf _
                                    , vbYesNo + vbCritical, "Confirm Delete") Then Exit Sub
                    
                    Call HFApp.SqlExec("DELETE FROM tblPOIndex WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, POIndex))
                    Call HFApp.SqlExec("UPDATE tblPhaseItem SET POIndex='' WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, POIndex))
                    
                    
                    Call .RemoveItem
                End If
            
            Case "vendor"
                Call HFApp.RunTask("EditDefaultVendors")
            
            Case "Items"
                FPOItems.Show vbModal
                
            Case "preview"
                If .Row > -1 Then
                If .RowData(.Row) & "" <> "" And .TextMatrix(.Row, .ColIndex("POFormat")) <> "" Then
                    On Error Resume Next
                    PONumber = ""
                    PONumber = HFApp.SqlExec("select top 1 m.ponumber from pomaster m join poitems i on m.ponumber=i.ponumber", dbHomefront)(0)
                    On Error GoTo eh

                    s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", .TextMatrix(.Row, .ColIndex("POFormat")) & ".rpt")
                    'Call v.ShowReport(HFApp.ConnectionString(dbHomefront), s, rvPreview, "", "", "PONumber", "'" & PONumber & "'")
                    Dim c As New ZybUtil.Crystal
                    Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
                    On Error Resume Next
                    Call c.ParameterValue("DivisionID", HFApp.DivisionID)
                    Call c.ParameterValue("PONumber", "'" & PONumber & "'")
                    On Error GoTo eh
                    Call c.PrintPreview("Print Preview")
                        
                End If
                End If
        End Select
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonClick")
End Sub

