VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FImportPricelists 
   Caption         =   "Import Pricelist Wizard"
   ClientHeight    =   6165
   ClientLeft      =   3360
   ClientTop       =   1515
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FImportPricelists.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7305
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7305
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5580
      Width           =   7305
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   -60
         X2              =   26420
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
      TabIndex        =   14
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1588
      Caption         =   "Import Vendor Pricelists"
      Description     =   "The vendor pricelist import wizard will help you update your pricing database"
      Icon            =   "FImportPricelists.frx":000C
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   930
      Width           =   7515
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   3060
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   556
         Picture         =   "FImportPricelists.frx":08E6
         ForeColor       =   0
         BarPicture      =   "FImportPricelists.frx":0902
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   2655
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   7155
         _cx             =   1996108109
         _cy             =   1996100171
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FImportPricelists.frx":091E
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
      Begin VB.Label lblFileDate 
         AutoSize        =   -1  'True
         Caption         =   "Last Modified April 18, 2007"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label lblFilename 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   45
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   5115
      Index           =   1
      Left            =   8490
      TabIndex        =   7
      Top             =   2805
      Visible         =   0   'False
      Width           =   7515
      Begin VB.Frame frmOptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1695
         Index           =   1
         Left            =   1020
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   5835
         Begin VSFlex8Ctl.VSFlexGrid gRejects 
            Height          =   855
            Left            =   4065
            TabIndex        =   28
            Top             =   480
            Visible         =   0   'False
            Width           =   1275
            _cx             =   1996097737
            _cy             =   1996096996
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
         Begin VB.CheckBox chkIgnoreZeroValues 
            Caption         =   "Ignore records with no price info (items priced at $0.00)"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Value           =   1  'Checked
            Width           =   4875
         End
         Begin VB.TextBox txtEffectiveDate 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   22
            Top             =   885
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.ComboBox cboImportColumn 
            Height          =   240
            Left            =   1200
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   630
            Width           =   1755
         End
         Begin VB.Label lblDestination 
            AutoSize        =   -1  'True
            Caption         =   "Destination:"
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
            Left            =   0
            TabIndex        =   27
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label lblEffectiveDate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Effective:"
            Height          =   195
            Left            =   420
            TabIndex        =   26
            Top             =   900
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblImportColumn 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Import into:"
            Height          =   195
            Left            =   315
            TabIndex        =   25
            Top             =   630
            Width           =   780
         End
         Begin VB.Image cmdEffectiveDate 
            Height          =   240
            Left            =   2730
            Picture         =   "FImportPricelists.frx":09F3
            Top             =   885
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Frame frmOptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   675
         Index           =   0
         Left            =   1020
         TabIndex        =   17
         Top             =   2340
         Visible         =   0   'False
         Width           =   5595
         Begin VB.TextBox txtRejectFile 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1200
            TabIndex        =   18
            Top             =   270
            Width           =   3915
         End
         Begin VB.CheckBox chkIgnoreErrors 
            Caption         =   "Import anyway (invalid records will be written to the reject file)"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   4875
         End
         Begin VB.Image cmdChooseFolder 
            Height          =   240
            Index           =   0
            Left            =   5220
            Picture         =   "FImportPricelists.frx":0B3D
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Reject File:"
            Height          =   195
            Left            =   300
            TabIndex        =   20
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Options:"
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
         Left            =   900
         TabIndex        =   15
         Top             =   2100
         Width           =   720
      End
      Begin VB.Image imgOK 
         Height          =   480
         Left            =   300
         Picture         =   "FImportPricelists.frx":10C7
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblShowReport 
         Caption         =   "Open the pricing comparison report."
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
         Left            =   1140
         TabIndex        =   13
         Top             =   1680
         Width           =   2715
      End
      Begin VB.Label lblShowReport 
         Caption         =   "Open the import summary."
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
         Height          =   255
         Index           =   0
         Left            =   1140
         TabIndex        =   12
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblErrors 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The pricelist import wizard is not able to import this data."
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
         Left            =   900
         TabIndex        =   10
         Top             =   180
         UseMnemonic     =   0   'False
         Width           =   4830
      End
      Begin VB.Image imgError 
         Height          =   480
         Left            =   300
         Picture         =   "FImportPricelists.frx":1991
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "FImportPricelists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FImportPricelists::"
Private mSessionID As String
Private mFilename  As String
Private mTempFile  As String 'is set and created in ConvertProfitMasterFile then deleted in LoadExcelSheet() in Let CurrentFrame()
Private Type ProfitMasterRec
    cost          As String * 8
    s1            As String * 1
    Retail        As String * 8
    s2            As String * 1
    Disc1         As String * 8
    s3            As String * 1
    Disc2         As String * 8
    s4            As String * 1
    Disc3         As String * 8
    s5            As String * 1
    Filler1       As String * 8
    s6            As String * 1
    Filler2       As String * 4
    s7            As String * 1
    Class         As String * 12
    s8            As String * 1
    ProductNumber As String * 16
    s9            As String * 1
    Filler3       As String * 12
    s10           As String * 1
    Desc1         As String * 30
    s11           As String * 1
    UOM           As String * 3
    s12           As String * 1
    Filler4       As String * 8
    s13           As String * 1
    QtyOnHand     As String * 7
    sCRLF         As String * 2
End Type

Private bCorporate As Boolean


Public Function ShowForm()
On Error GoTo eh
    Dim FilterList As String
    Dim ExtensionNumber As Long
    
    mFilename = IniGet(AppIni, "Options", "ImportPricelistFile")
    If Not FileExists(mFilename) Then mFilename = ""
    
    
    ExtensionNumber = Val("" & IniGet(AppIni, "Options", "ImportPricelistType"))

    FilterList = "Vendor Pricelists|*.xls*|ProfitMaster Export Files|*.txt"

    mTempFile = ""
    If VBGetOpenFileName(mFilename, , , , , True, FilterList, ExtensionNumber, , , , FMain.hwnd) Then
        Call IniPut(AppIni, "Options", "ImportPricelistFile", mFilename)
        Call IniPut(AppIni, "Options", "ImportPricelistType", ExtensionNumber)
        Select Case ExtensionNumber
            Case 1: mTempFile = "" 'normal HF pricelist.
            Case 2: Call ConvertProfitMasterFile
        End Select
        CurrentFrame = 0
        Me.Show vbModal
    Else
        Unload Me
    End If
        
Exit Function
eh: Call errHandler(SRCFILE & "ShowForm")
End Function


Private Function TextMatrix(r As Long, ColKey As String) As String
On Error Resume Next
    TextMatrix = gData.TextMatrix(r, gData.ColIndex(ColKey))
End Function

Private Function ValidateGrid() As Boolean
    Dim r As Long
    Dim c As Long
    Dim s As String
    With gData
    
        'find header row
        bCorporate = False
        For r = 0 To .Rows - 1
            If IsIn(Trim(.TextMatrix(r, 1)), "Community", FMain.CD_Community) Then Exit For
            If IsIn(Trim(.TextMatrix(r, 1)), "Group", FMain.CD_Community) Then
                bCorporate = True
                Exit For
            End If
        Next
        'set colindexes
        For c = 1 To .Cols - 1
            .ColKey(c) = .TextMatrix(r, c)
        Next
        If .ColIndex("Specification") <> -1 Then .ColKey(.ColIndex("Specification")) = "Assembly"
    
    
        'replace custom descriptions
        If .ColIndex(FMain.CD_Community) <> -1 Then .ColKey(.ColIndex(FMain.CD_Community)) = "Community"
        If .ColIndex(FMain.CD_Community & " Name") <> -1 Then .ColKey(.ColIndex(FMain.CD_Community & " Name")) = "Community Name"
        
        
        'verify that all columns are present
        If bCorporate Then
            If .ColIndex("Group") = -1 Then s = s & ",Group"
            If .ColIndex("Item") = -1 Then s = s & ",Item"
            If .ColIndex("Item Description") = -1 Then s = s & ",Item Description"
            If .ColIndex("Part Number") = -1 Then s = s & ",Part Number"
            If .ColIndex("UOM") = -1 Then s = s & ",UOM"
            If .ColIndex("Price") = -1 Then s = s & ",Price"
        Else
            If .ColIndex("Community") = -1 Then s = s & ",Community"
            If .ColIndex("Community Name") = -1 Then s = s & ",Community Name"
            If .ColIndex("Phase") = -1 Then s = s & ",Phase"
            If .ColIndex("Model") = -1 Then s = s & ",Model"
            If .ColIndex("Assembly") = -1 Then s = s & ",Specification"
            If .ColIndex("Assembly Description") = -1 Then s = s & ",Specification Description"
            If .ColIndex("Group") = -1 Then s = s & ",Group"
            If .ColIndex("Item") = -1 Then s = s & ",Item"
            If .ColIndex("Item Description") = -1 Then s = s & ",Item Description"
            If .ColIndex("Part Number") = -1 Then s = s & ",Part Number"
            If .ColIndex("UOM") = -1 Then s = s & ",UOM"
            If .ColIndex("Price") = -1 Then s = s & ",Price"
        End If
        
        If .ColIndex("Next Price 1") = -1 And .ColIndex("Next Price 2") = -1 And .ColIndex("Effective 1") = -1 And .ColIndex("Effective 2") = -1 Then
            lblDestination.Visible = True
            lblImportColumn.Visible = True
            cboImportColumn.Visible = True
        Else
            lblDestination.Visible = False
            lblImportColumn.Visible = False
            cboImportColumn.Visible = False
            cboImportColumn.ListIndex = 0 'current
        End If
        
        
    End With
    
    If s = "" Then
        ValidateGrid = True
    Else
        cmdNav(3).Enabled = False
        imgError.Visible = True
        imgOK.Visible = False
        lblDescription.Caption = "The pricelist import wizard is not able to import this file."
        lblErrors.Caption = "The file does not conform to Precision Builder specifications." & vbCrLf & "Unable to find columns (" & Mid(s, 2) & ")"
        ValidateGrid = False
    End If
End Function

Private Function ValidateTable() As Boolean
    Dim s As String
    Dim rs As Recordset
    
    'if tempfile has a path then this data came from a profitmaster and the partnumber needs to be looked up.
    If mTempFile <> "" Then
        s = ""
        s = s & "update importedcosts" & vbCrLf
        s = s & "   set phase=pi.phase" & vbCrLf
        s = s & "      ,item=pi.item" & vbCrLf
        s = s & "  from importedcosts ic" & vbCrLf
        s = s & "       join tblphaseitem pi on pi.DivisionID = " & HFApp.DivisionID & " and ic.partnumber=pi.partnumber" & vbCrLf
        s = s & " where isnull(ic.partnumber,'')<>'' and isnull(ic.phase,'')='' and isnull(ic.item,'')=''" & vbCrLf
        s = s & "   and SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
    End If
    
    
    
    s = ""
    s = s & "UPDATE ImportedCosts" & vbCrLf
    s = s & "   SET UnknownVendor= case when v.vendor_id is null then 1 else 0 end" & vbCrLf
    s = s & "      ,UnknownCommunity= case when c.community<>'' and l.area is null then 1 else 0 end" & vbCrLf
    s = s & "      ,UnknownAssembly= case when c.Assembly<>'' AND a.Assembly IS NULL then 1 else 0 end" & vbCrLf
    s = s & "      ,UnknownItem= case when i.Phase is null then 1 else 0 end" & vbCrLf
    s = s & "      ,UnknownUOM= case when upper(rtrim(ltrim(c.OrderUOM)))<>upper(rtrim(ltrim(i.OrderUOM))) then 1 else 0 end" & vbCrLf
    s = s & "      ,ExpectedUOM= case when upper(rtrim(ltrim(c.OrderUOM)))<>upper(rtrim(ltrim(i.OrderUOM))) then i.OrderUOM else '' end" & vbCrLf
    s = s & "      ,UnknownExpiryDate = case when ((Next_Cost1<>0 AND Next_Effective1 IS NULL) OR (Next_Cost2<>0 AND Next_Effective2 IS NULL)) then 1 else 0 end" & vbCrLf
    s = s & "  FROM ImportedCosts c" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblVendors v ON (c.Vendor=v.Vendor_ID and v.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblLocality l ON (c.Community=l.Area)" & vbCrLf
    s = s & "       LEFT OUTER JOIN DistinctAssemblies a ON (c.Assembly=a.Assembly)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPhaseItem i ON (i.DivisionID = " & HFApp.DivisionID & " and c.Phase=i.Phase AND c.Item=i.Item)" & vbCrLf
    s = s & " WHERE SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "   AND ((c.Phase<>'') or (c.phase='' and c.item=''))"
    Set rs = HFApp.SqlExec(s)
    
    
    s = ""
    s = s & "SELECT DISTINCT 'The file contains unrecognized vendor(s)'" & vbCrLf
    s = s & "  FROM ImportedCosts c" & vbCrLf
    s = s & " WHERE UnknownVendor=1" & vbCrLf
    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT DISTINCT 'The file contains unrecognized community(s)'" & vbCrLf
    s = s & "  FROM ImportedCosts c" & vbCrLf
    s = s & " WHERE UnknownCommunity=1" & vbCrLf
    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT DISTINCT 'The file contains unrecognized assembly(s)'" & vbCrLf
    s = s & "  FROM ImportedCosts c" & vbCrLf
    s = s & " WHERE UnknownAssembly=1" & vbCrLf
    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT DISTINCT 'The file contains unrecognized item(s)'" & vbCrLf
    s = s & "  FROM ImportedCosts c" & vbCrLf
    s = s & " WHERE UnknownItem=1" & vbCrLf
    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT DISTINCT 'The file contains conflicting units'" & vbCrLf
    s = s & "  FROM ImportedCosts c" & vbCrLf
    s = s & " WHERE UnknownUOM=1" & vbCrLf
    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT DISTINCT 'The file contains future pricing with no effective date(s)'" & vbCrLf
    s = s & "  FROM ImportedCosts c" & vbCrLf
    s = s & " WHERE UnknownExpiryDate=1" & vbCrLf
    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    s = ""
    While Not rs.EOF
        s = s & "" & rs(0) & vbCrLf
        rs.MoveNext
    Wend
    
    If s <> "" Then
        cmdNav(3).Enabled = False
        imgError.Visible = True
        imgOK.Visible = False
        lblDescription.Caption = "The pricelist import wizard is not able to import this data."
        lblErrors.Caption = s
        ValidateTable = False
        
        frmOptions(0).Visible = True
        frmOptions(1).Top = frmOptions(0).Top + frmOptions(0).Height
    Else
        imgError.Visible = False
        imgOK.Visible = True
        ValidateTable = True
        lblDescription.Caption = "The pricelist import wizard is ready to process the file."
        lblErrors.Caption = "View the reports listed below to see how this will change your vendor pricing" & vbCrLf & _
                            "database. If you are satisfied with the changes and would like to commit them" & vbCrLf & _
                            "to the database click ""Finish"". If you are not satisfied click ""Cancel"" to" & vbCrLf & _
                            "exit the wizard without updating your database."
    
        frmOptions(1).Visible = True
        frmOptions(1).Top = frmOptions(0).Top
    
    End If
    
    
    
    
    

    
    
End Function

Private Property Get CurrentFrame() As Long
    Dim i As Long
    For i = 0 To WizFrame.UBound
        If WizFrame(i).Visible = True Then
            CurrentFrame = i
            Exit Property
        End If
    Next
    CurrentFrame = WizFrame.LBound
End Property

Private Property Let CurrentFrame(RHS As Long)
On Error GoTo eh
    Dim i As Long
    Dim fileinfo As New ClsFileInfo
    Dim s As String
    Dim r As Long
    Dim Vendor As String
    Dim VendorDesc As String
    
    
    Screen.MousePointer = vbHourglass
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
        
        
    With gData
    Select Case RHS
    
        Case 0 ' show file
            If WizFrame(RHS).Tag = "" Then
                WizFrame(RHS).Tag = "LOADED"
                lblFilename.Caption = "File name: " & mFilename
                fileinfo.FullPathName = mFilename
                lblFileDate.Caption = "Last Modified: " & format(fileinfo.ModifyTime, "long date")
                
                If mTempFile = "" Then
                    Call LoadExcelSheet(GetObject(mFilename), "", gData)
                Else
                    Call LoadExcelSheet(, , gData, mTempFile)
                    For i = 1 To 8
                        .ColHidden(i) = True
                    Next
                    .ColHidden(13) = True
                    .ColHidden(14) = True
                    .RowHidden(1) = True
                    .RowHidden(2) = True
                    
                    
                End If
                If gData.Cols > 19 Then gData.Cols = 19
                
            End If
            
        Case 1 ' write data then show options
            If WizFrame(RHS).Tag = "" Then
                WizFrame(RHS).Tag = "LOADED"
                If ValidateGrid() Then
                    r = 1
                    WizFrame(0).Visible = True
                    ProgressBar.Visible = True
                    WizFrame(1).Visible = False
                    While r < .Rows
                        ProgressBar.value = (r / (.Rows - 1)) * 100
                        If .TextMatrix(r, 1) = "Vendor" Then
                            Vendor = .TextMatrix(r, 2)
                            If r < .Rows Then VendorDesc = .TextMatrix(r + 1, 2)
                            r = r + 1
                        End If
                        
                        If r < .Rows Then
                            If Trim(.TextMatrix(r, .ColIndex("Price"))) <> "" And Trim(.TextMatrix(r, .ColIndex("Price"))) <> "Price" Then
                            
                                If bCorporate Then
                                    s = ""
                                    s = s & "INSERT INTO ImportedCosts(SessionID,FileName,Community,CommunityDesc,CommunityPhase,Vendor,VendorDesc,Model,Assembly,AssemblyDesc,Phase,Item,ItemDesc,PartNumber,OrderUOM,Current_Cost,Next_Cost1,Next_Effective1,Next_Cost2,Next_Effective2,PriceLink,DivisionID)" & vbCrLf
                                    s = s & "VALUES(" & DbQuote(Str, mSessionID, , , 50) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, FileTitle(mFilename), , True, 100) & vbCrLf
                                    s = s & "      ,''" & vbCrLf
                                    s = s & "      ,''" & vbCrLf
                                    s = s & "      ,''" & vbCrLf
                                    s = s & "      ," & DbQuote(Str, Vendor, , True, 40) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, VendorDesc, , True, 100) & vbCrLf
                                    s = s & "      ,''" & vbCrLf
                                    s = s & "      ,''" & vbCrLf
                                    s = s & "      ,''" & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Group")), , True, 20) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item")), , True, 16) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item Description")), , True, 100) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Part Number")), , True, 50) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("UOM")), , True, 10) & vbCrLf
                                    s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Price"))) & vbCrLf
                                    s = s & "      ," & DbQuote(Num, MYTextMatrix(r, .ColIndex("Next Price 1"))) & vbCrLf
                                    s = s & "      ," & DbQuote(Date, MYTextMatrix(r, .ColIndex("Effective 1"))) & vbCrLf
                                    s = s & "      ," & DbQuote(Num, MYTextMatrix(r, .ColIndex("Next Price 2"))) & vbCrLf
                                    s = s & "      ," & DbQuote(Date, MYTextMatrix(r, .ColIndex("Effective 2"))) & vbCrLf
                                    If ("" & .TextMatrix(r, .ColIndex("Group"))) = "" And ("" & .TextMatrix(r, .ColIndex("Item"))) <> "" Then
                                        s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item")), , True, 10) & vbCrLf
                                    Else
                                        s = s & "      ,0" & vbCrLf
                                    End If
                                    s = s & "      ," & HFApp.DivisionID & vbCrLf
                                    s = s & ")"
                                Else
                                    s = ""
                                    s = s & "INSERT INTO ImportedCosts(SessionID,FileName,Community,CommunityDesc,CommunityPhase,Vendor,VendorDesc,Model,Assembly,AssemblyDesc,Phase,Item,ItemDesc,PartNumber,OrderUOM,Current_Cost,Next_Cost1,Next_Effective1,Next_Cost2,Next_Effective2,PriceLink,DivisionID)" & vbCrLf
                                    s = s & "VALUES(" & DbQuote(Str, mSessionID, , , 50) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, FileTitle(mFilename), , True, 100) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Community")), , True, 10) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Community Name")), , True, 100) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase")), , True, 20) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, Vendor, , True, 40) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, VendorDesc, , True, 100) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Model")), , True, 20) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly")), , True, 20) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly Description")), , True, 100) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Group")), , True, 20) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item")), , True, 16) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item Description")), , True, 100) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Part Number")), , True, 50) & vbCrLf
                                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("UOM")), , True, 10) & vbCrLf
                                    s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Price"))) & vbCrLf
                                    s = s & "      ," & DbQuote(Num, MYTextMatrix(r, .ColIndex("Next Price 1"))) & vbCrLf
                                    s = s & "      ," & DbQuote(Date, MYTextMatrix(r, .ColIndex("Effective 1"))) & vbCrLf
                                    s = s & "      ," & DbQuote(Num, MYTextMatrix(r, .ColIndex("Next Price 2"))) & vbCrLf
                                    s = s & "      ," & DbQuote(Date, MYTextMatrix(r, .ColIndex("Effective 2"))) & vbCrLf
                                    If ("" & .TextMatrix(r, .ColIndex("Group"))) = "" And ("" & .TextMatrix(r, .ColIndex("Item"))) <> "" Then
                                        s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item")), , True, 10) & vbCrLf
                                    Else
                                        s = s & "      ,0" & vbCrLf
                                    End If
                                    s = s & "      ," & HFApp.DivisionID & vbCrLf
                                    s = s & ")"
                                End If
                                HFApp.SqlExec s
                            End If
                        End If
                        r = r + 1
                        
                    Wend
                    WizFrame(0).Visible = False
                    ProgressBar.Visible = False
                    WizFrame(1).Visible = True
                    Call ValidateTable
                    Screen.MousePointer = vbDefault
                End If
            End If
                
    End Select
    End With
    
    
    
    
    Screen.MousePointer = vbDefault
    Exit Property
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "CurrentFrame", s)
    End If
End Property

Private Sub chkIgnoreErrors_Click()
    frmOptions(1).Visible = chkIgnoreErrors.value = vbChecked
    cmdNav(3).Enabled = chkIgnoreErrors.value = vbChecked And Trim(txtRejectFile.Text) <> ""
End Sub

Private Sub cmdChooseFolder_Click(Index As Integer)
    Dim s As String
    txtRejectFile.SetFocus
    s = txtRejectFile.Text
    If VBGetOpenFileName(s, , False, , , True, "Excel Files (*.xls)|*.xls", , , , "xls", Me.hwnd) Then
        txtRejectFile.Text = s
        cmdNav(3).Enabled = chkIgnoreErrors.value = vbChecked And Trim(txtRejectFile.Text) <> ""
    End If
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3: If SaveData Then Unload Me
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).BorderStyle = 0
        WizFrame(i).Move 0, WizHead1.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead1.Height - WizFoot.Height
    Next
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
    gData.Move margin, gData.Top, WizFrame(0).Width - 2 * margin, WizFrame(0).Height - gData.Top - margin
    ProgressBar.Move gData.Left, gData.Top + gData.Height - ProgressBar.Height, gData.Width
End Sub

Private Sub Form_Load()
On Error GoTo eh
    mSessionID = MachineName & App.hInstance
    Call IniGetForm(Me)
    Call LoadCostTypes(cboImportColumn, , True)
    Call cboImportColumn.AddItem("Current", 0)
    Call cboImportColumn.AddItem("Next 1", 1)
    Call cboImportColumn.AddItem("Next 2", 2)
    cboImportColumn.ListIndex = HFApp.Options(PricelistImportDestination)
    chkIgnoreZeroValues.value = IIf(HFApp.Options(PricelistImportSkipZeros), vbChecked, vbUnchecked)
Exit Sub
eh: Call errHandler(SRCFILE & "SaveData")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HFApp.Databases(dbHomefront).CommandTimeout = 0
    Call HFApp.SqlExec("DELETE FROM ImportedCosts WHERE SessionID=" & DbQuote(Str, mSessionID))
    HFApp.Options.value(PricelistImportSkipZeros) = chkIgnoreZeroValues.value = vbChecked
    HFApp.Options.value(PricelistImportDestination) = cboImportColumn.ListIndex

    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub


Private Sub lblShowReport_Click(Index As Integer)
On Error Resume Next
    Dim s As String
    Select Case Index
        'Case 0: Call FRptViewer.ShowReport(HFApp.SystemFolder & "System\Reports\Estimating\PriceImportSummary.rpt", True, True, "SessionID", mSessionID, "ImportDestination", cboImportColumn.ListIndex)
        'Case 1: Call FRptViewer.ShowReport(HFApp.SystemFolder & "System\Reports\Estimating\PriceImportComparison.rpt", True, True, "SessionID", mSessionID, "ImportDestination", cboImportColumn.ListIndex)
        Case 0: s = HFApp.SystemFolder & "System\Reports\Estimating\PriceImportSummary.rpt"
        Case 1: s = HFApp.SystemFolder & "System\Reports\Estimating\PriceImportComparison.rpt"
    End Select
    If s = "" Then Exit Sub
    
    Dim c As New ZybUtil.Crystal
    Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
    On Error Resume Next
    Call c.ParameterValue("DivisionID", HFApp.DivisionID)
    Call c.ParameterValue("SessionID", mSessionID)
    Call c.ParameterValue("ImportDestination", cboImportColumn.ListIndex)
    'On Error GoTo eh
    Call c.PrintPreview("Print Preview")
    
    
    
End Sub

Private Function SaveData() As Boolean
On Error GoTo eh
    Dim s As String
    
    
    s = ""
    s = s & "UPDATE tblVendorCost " & vbCrLf
    Select Case cboImportColumn.ListIndex
        Case 0
            If chkIgnoreZeroValues.value = vbChecked Then
                s = s & "   SET Current_Cost=   Case when ic.Current_Cost <> 0                                  then ic.Current_Cost    else vc.Current_Cost    end" & vbCrLf
                s = s & "      ,Next_Cost1=     Case when ic.Next_Cost1 <> 0 And Not ic.Next_Effective1 Is Null then ic.Next_Cost1      else vc.Next_Cost1      end" & vbCrLf
                s = s & "      ,Next_Effective1=Case when ic.Next_Cost1 <> 0 And Not ic.Next_Effective1 Is Null then ic.Next_Effective1 else vc.Next_Effective1 end" & vbCrLf
                s = s & "      ,Next_Cost2=     Case when ic.Next_Cost2 <> 0 And Not ic.Next_Effective2 Is Null then ic.Next_Cost2      else vc.Next_Cost2      end" & vbCrLf
                s = s & "      ,Next_Effective2=Case when ic.Next_Cost2 <> 0 And Not ic.Next_Effective2 Is Null then ic.Next_Effective2 else vc.Next_Effective2 end" & vbCrLf
            Else
                s = s & "   SET Current_Cost=ic.Current_Cost" & vbCrLf
                s = s & "      ,Next_Cost1=     Case when Not ic.Next_Effective1 Is Null then ic.Next_Cost1      else vc.Next_Cost1      end" & vbCrLf
                s = s & "      ,Next_Effective1=Case when Not ic.Next_Effective1 Is Null then ic.Next_Effective1 else vc.Next_Effective1 end" & vbCrLf
                s = s & "      ,Next_Cost2=     Case when Not ic.Next_Effective2 Is Null then ic.Next_Cost2      else vc.Next_Cost2      end" & vbCrLf
                s = s & "      ,Next_Effective2=Case when Not ic.Next_Effective2 Is Null then ic.Next_Effective2 else vc.Next_Effective2 end" & vbCrLf
            End If
            
        Case 1
            If Not IsDate(txtEffectiveDate) Then
                MsgBox "You must specify an effective date", vbExclamation, App.ProductName
                Exit Function
            End If
            If chkIgnoreZeroValues.value = vbChecked Then
                s = s & "   SET Next_Cost1=Case when ic.Current_Cost <> 0 then ic.Current_Cost else vc.Next_Cost1 end" & vbCrLf
            Else
                s = s & "   SET Next_Cost1=ic.Current_Cost" & vbCrLf
            End If
            s = s & "      ,Next_Effective1=" & DbQuote(Date, txtEffectiveDate) & vbCrLf
            
        Case 2
            If Not IsDate(txtEffectiveDate) Then
                MsgBox "You must specify an effective date", vbExclamation, App.ProductName
                Exit Function
            End If
            If chkIgnoreZeroValues.value = vbChecked Then
                s = s & "   SET Next_Cost2=Case when ic.Current_Cost <> 0 then ic.Current_Cost else vc.Next_Cost2 end" & vbCrLf
            Else
                s = s & "   SET Next_Cost2=ic.Current_Cost" & vbCrLf
            End If
            s = s & "      ,Next_Effective2=" & DbQuote(Date, txtEffectiveDate) & vbCrLf
            
            Case Else
            If chkIgnoreZeroValues.value = vbChecked Then
                s = s & "   SET Forecast" & cboImportColumn.ListIndex - 2 & "=case when ic.Current_Cost<>0 then ic.Current_Cost else vc.Current_Cost end" & vbCrLf
            Else
                s = s & "   SET Forecast" & cboImportColumn.ListIndex - 2 & "=ic.Current_Cost" & vbCrLf
            End If
            
    End Select
    s = s & "      ,Last_Cost3=vc.Last_Cost2" & vbCrLf
    s = s & "      ,Last_Cost2=vc.Last_Cost1" & vbCrLf
    s = s & "      ,Last_Cost1=vc.Current_Cost" & vbCrLf
    s = s & "      ,Last3_Expiry=vc.Last2_Expiry" & vbCrLf
    s = s & "      ,Last2_Expiry=vc.Last1_Expiry" & vbCrLf
    s = s & "      ,Last1_Expiry=getdate()-1" & vbCrLf
    s = s & "      ,PartNumber=ic.PartNumber" & vbCrLf
    s = s & "  FROM ImportedCosts ic " & vbCrLf
    s = s & "  JOIN tblPhaseItem pi on(pi.DivisionID = " & HFApp.DivisionID & " and ((ic.phase=pi.phase and ic.item=pi.item) or (isnull(ic.phase,'')='' and ic.item=cast(pi.PriceLink as varchar))))" & vbCrLf
    s = s & "  JOIN tblVendorCost vc " & vbCrLf
    s = s & "    ON(vc.DivisionID = " & IIf(bCorporate, 0, HFApp.DivisionID) & " and" & vbCrLf
    s = s & "       ic.Community=vc.Community AND" & vbCrLf
    s = s & "       ic.CommunityPhase=vc.CommunityPhase AND" & vbCrLf
    s = s & "       ic.Assembly=vc.Assembly AND" & vbCrLf
    s = s & "       ic.Model=vc.Model AND" & vbCrLf
    s = s & "       pi.Phase=vc.Phase AND" & vbCrLf
    s = s & "       pi.Item=vc.Item AND" & vbCrLf
    s = s & "       ic.Vendor=vc.Vendor AND" & vbCrLf
    s = s & "       ic.UnknownVendor=0 AND" & vbCrLf
    s = s & "       ic.UnknownCommunity=0 AND" & vbCrLf
    s = s & "       ic.UnknownAssembly=0 AND" & vbCrLf
    s = s & "       ic.UnknownItem=0 AND" & vbCrLf
    s = s & "       ic.UnknownUOM=0 AND" & vbCrLf
    s = s & "       ic.UnknownExpiryDate=0 AND" & vbCrLf
    s = s & "       ic.SessionID=" & DbQuote(Str, mSessionID) & ")" & vbCrLf
    Call HFApp.SqlExec(s)
    
    
    
    s = ""
    s = s & "INSERT INTO tblVendorCost(Community,DivisionID,CommunityPhase,Assembly,Model,Phase,Item,PriceLink,Vendor" & vbCrLf
    s = s & "      ,Current_Cost" & vbCrLf
    s = s & "      ,Next_Cost1,Next_Effective1" & vbCrLf
    s = s & "      ,Next_Cost2,Next_Effective2" & vbCrLf
    s = s & "      ,PartNumber,Forecast1,Forecast2,Forecast3,Forecast4,Forecast5,Forecast6,Forecast7,Forecast8,Forecast9,Forecast10,Forecast11,Forecast12)" & vbCrLf
    s = s & "SELECT ic.Community," & IIf(bCorporate, 0, HFApp.DivisionID) & ",ic.CommunityPhase,ic.Assembly,ic.Model,pi.Phase,pi.Item,case when ISNULL(ic.Assembly,'')<>'' then 0 else pi.PriceLink end,ic.Vendor" & vbCrLf
    s = s & "      ,MAX(ic.Current_Cost)" & vbCrLf
    s = s & "      ,MAX(ic.Next_Cost1)" & vbCrLf
    s = s & "      ,case when MAX(ic.Next_Cost1)<>0 then MIN(ic.Next_Effective1) else NULL end" & vbCrLf
    s = s & "      ,MAX(ic.Next_Cost2)" & vbCrLf
    s = s & "      ,case when MAX(ic.Next_Cost2)<>0 then MIN(ic.Next_Effective2) else NULL end" & vbCrLf
    s = s & "      ,max(ic.Partnumber)" & vbCrLf
    s = s & "      ,MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost)" & vbCrLf
    s = s & "  FROM ImportedCosts ic " & vbCrLf
    s = s & "  JOIN tblPhaseItem pi on(pi.DivisionID = " & HFApp.DivisionID & " and ((ic.phase=pi.phase and ic.item=pi.item) or (isnull(ic.phase,'')='' and ic.item=cast(pi.PriceLink as varchar))))" & vbCrLf
    s = s & "  LEFT OUTER JOIN tblVendorCost vc " & vbCrLf
    s = s & "    ON(vc.DivisionID = " & IIf(bCorporate, 0, HFApp.DivisionID) & " and" & vbCrLf
    s = s & "       ic.Community=vc.Community AND" & vbCrLf
    s = s & "       ic.CommunityPhase=vc.CommunityPhase AND" & vbCrLf
    s = s & "       ic.Assembly=vc.Assembly AND" & vbCrLf
    s = s & "       ic.model=vc.model AND" & vbCrLf
    s = s & "       pi.Phase=vc.Phase AND" & vbCrLf
    s = s & "       pi.Item=vc.Item AND" & vbCrLf
    s = s & "       ic.Vendor=vc.Vendor AND" & vbCrLf
    s = s & "       ic.SessionID=" & DbQuote(Str, mSessionID) & ")" & vbCrLf
    s = s & " WHERE vc.Vendor IS NULL" & vbCrLf
    s = s & "   AND ic.SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "   AND ic.UnknownVendor=0" & vbCrLf
    s = s & "   AND ic.UnknownCommunity=0" & vbCrLf
    s = s & "   AND ic.UnknownAssembly=0" & vbCrLf
    s = s & "   AND ic.UnknownItem=0" & vbCrLf
    s = s & "   AND ic.UnknownUOM=0" & vbCrLf
    s = s & "   AND ic.UnknownExpiryDate=0" & vbCrLf
    s = s & "GROUP BY ic.Community,ic.CommunityPhase,ic.Model,ic.Assembly,pi.PriceLink,pi.Phase,pi.Item,ic.Vendor" & vbCrLf
    Call HFApp.SqlExec(s)


    SaveData = True
    If Me.chkIgnoreErrors.value = vbChecked Then Call WriteRejectFile
    Exit Function
    
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function

Private Sub cmdEffectiveDate_Click()
On Error Resume Next
    Call txtEffectiveDate_KeyDown(vbKeyF4, 0)
End Sub
Private Sub txtEffectiveDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo eh
    If KeyCode = vbKeyF4 And Shift = 0 Then Call DCalendar.Popup(txtEffectiveDate)
Exit Sub
eh: Call errHandler(SRCFILE & "txtEffectiveDate_KeyDown")
End Sub
Private Sub txtEffectiveDate_Validate(Cancel As Boolean)
On Error GoTo eh
    If IsDate(txtEffectiveDate) Or txtEffectiveDate = "" Then
        txtEffectiveDate.Text = format(txtEffectiveDate, HFApp.Options(DateFormat))
    Else
        MsgBox "Not a valid date", vbExclamation, App.ProductName
        Call SelectAll(txtEffectiveDate)
        Cancel = True
    End If
    Exit Sub
eh: Call errHandler(SRCFILE & "txtEffectiveDate_Validate")
End Sub
Private Sub cboImportColumn_Click()
    cmdEffectiveDate.Visible = cboImportColumn.ListIndex = 1 Or cboImportColumn.ListIndex = 2
    lblEffectiveDate.Visible = cmdEffectiveDate.Visible
    txtEffectiveDate.Visible = cmdEffectiveDate.Visible
End Sub

Private Function MYTextMatrix(Row As Long, Col As Long) As String
On Error Resume Next
    MYTextMatrix = gData.TextMatrix(Row, Col)
End Function


Private Sub WriteRejectFile()
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    Dim i As Long
    Dim Vendor As String
    
    
    s = ""
    s = s & "SELECT Vendor" & vbCrLf
    s = s & "      ,VendorDesc Companyname" & vbCrLf
    s = s & "      ,Community" & vbCrLf
    s = s & "      ,communityPhase Phase" & vbCrLf
    s = s & "      ,CommunityDesc CommunityName" & vbCrLf
    s = s & "      ,Assembly" & vbCrLf
    s = s & "      ,AssemblyDesc AssemblyDescription" & vbCrLf
    s = s & "      ,Phase gGroup" & vbCrLf
    s = s & "      ,Item" & vbCrLf
    s = s & "      ,ItemDesc ItemDescription" & vbCrLf
    s = s & "      ,PartNumber" & vbCrLf
    s = s & "      ,OrderUOM UOM" & vbCrLf
    s = s & "      ,Current_Cost Price" & vbCrLf
    s = s & "      ,Next_Cost1 NextPrice1" & vbCrLf
    s = s & "      ,Next_Effective1 Effective1" & vbCrLf
    s = s & "      ,Next_Cost2 NextPrice2" & vbCrLf
    s = s & "      ,Next_Effective2 Effective2" & vbCrLf
    s = s & "      ,substring(case when UnknownVendor=1 then '; unknown vendor' else '' end" & vbCrLf
    s = s & "        + case when UnknownCommunity=1 then '; unknown community' else '' end" & vbCrLf
    s = s & "        + case when UnknownAssembly=1 then '; unknown assembly' else '' end" & vbCrLf
    s = s & "        + case when UnknownItem=1 then '; unknown phase/item' else '' end" & vbCrLf
    s = s & "        + case when UnknownUOM=1 then '; HF expects UOM to be ""' + ISNULL(ExpectedUOM,'') + '"" ' else '' end" & vbCrLf
    s = s & "        + case when UnknownExpiryDate=1 then '; effective date is required' else '' end,3,9999) Comments" & vbCrLf
    s = s & "  FROM ImportedCosts " & vbCrLf
    s = s & " WHERE SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "   AND (UnknownVendor=1 OR" & vbCrLf
    s = s & "        UnknownCommunity=1 OR" & vbCrLf
    s = s & "        UnknownAssembly=1 OR" & vbCrLf
    s = s & "        UnknownItem=1 OR" & vbCrLf
    s = s & "        UnknownUOM=1 OR" & vbCrLf
    s = s & "        UnknownExpiryDate=1)" & vbCrLf
    s = s & "ORDER BY 1,3,4,8,9" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    With gRejects
        .Rows = 0
        .Cols = 16
        .FixedCols = 0
        .FixedRows = 0
        While Not rs.EOF
            If Vendor <> "" & rs("vendor") Then
                Vendor = "" & rs("vendor")
                .AddItem ""
                .AddItem ""
                .AddItem "Vendor" & vbTab & Vendor
                .AddItem "Company Name" & vbTab & rs("CompanyName")
                .Cell(flexcpFontBold, .Rows - 2, 0, .Rows - 1, 0) = True
                .AddItem "Community" & vbTab & "Phase" & vbTab & "Community Name" & vbTab & "Assembly" & vbTab & "Assembly Description" & vbTab & "Group" & vbTab & "Item" & vbTab & "Item Description" & vbTab & "Part Number" & vbTab & "UOM" & vbTab & "Price" & vbTab & "Next Price 1" & vbTab & "Effective 1" & vbTab & "Next Price 2" & vbTab & "Effective 2"
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 15) = True
                
                .Refresh
                .Row = .Rows - 1
            End If
            .AddItem "" & rs("Community") & vbTab & rs("Phase") & vbTab & rs("CommunityName") & vbTab & rs("Assembly") & vbTab & rs("AssemblyDescription") & vbTab & rs("gGroup") & vbTab & rs("Item") & vbTab & rs("ItemDescription") & vbTab & rs("PartNumber") & vbTab & rs("UOM") & vbTab & rs("Price") & vbTab & rs("NextPrice1") & vbTab & rs("Effective1") & vbTab & rs("NextPrice2") & vbTab & rs("Effective2") & vbTab & rs("Comments")
            rs.MoveNext
        Wend
        .Cell(flexcpForeColor, 0, 15, .Rows - 1, 15) = vbRed
        
        
        Call .SaveGrid(txtRejectFile.Text, flexFileExcel)
        
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "WriteRejectFile")
End Sub



Private Sub txtRejectFile_Change()
On Error Resume Next
    cmdNav(3).Enabled = chkIgnoreErrors.value = vbChecked And Trim(txtRejectFile.Text) <> ""
End Sub



Private Sub ConvertProfitMasterFile()
    
    Dim i As Long
    Dim iRec As ProfitMasterRec
    Dim o As Long
    Dim s As String
    
    Dim Vendor As String
    Dim VendorName As String
    
    'open input file
    i = FreeFile
    Open mFilename For Random As #i Len = Len(iRec)
    
    'open output file
    o = FreeFile
    mTempFile = TempFile
    Open mTempFile For Output As #o

    'get vendor
    Vendor = IniGet(AppIni, "Options", "PMImportVendor")
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", "Select vendor_id Vendor,vendor_name Company from tblvendors where DivisionID = " & HFApp.DivisionID, Vendor) Then
        Vendor = FPickList.SelectedItem("Vendor")
        VendorName = FPickList.SelectedItem("Company")
        Call IniPut(AppIni, "Options", "PMImportVendor", Vendor)
    End If


    'write output file
    Print #o, "Vendor," & Vendor
    Print #o, "Company Name," & VendorName
    Print #o, "Community,Phase,Community Name,Model,Specification,Assembly Description,Group,Item,Item Description,Part Number,UOM,Price"
    While Not EOF(i)
        Get #i, , iRec
        If Trim(Replace(iRec.ProductNumber, Chr(0), "")) <> "" Then
            Print #o, ",,,,,,,," & Trim(iRec.Desc1) & "," & Trim(iRec.ProductNumber) & "," & Trim(iRec.UOM) & "," & Val(iRec.cost) / 100
        End If
    Wend
    
    Close #i
    Close #o

End Sub

