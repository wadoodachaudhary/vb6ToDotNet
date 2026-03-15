VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FJournalEntries 
   Caption         =   "Journal Entries"
   ClientHeight    =   5160
   ClientLeft      =   4125
   ClientTop       =   -11535
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   10440
   Begin VB.PictureBox pHeading 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10440
      TabIndex        =   2
      Top             =   570
      Width           =   10440
      Begin VB.Frame frmTotals 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   465
         Left            =   8430
         TabIndex        =   5
         Top             =   30
         Width           =   2025
         Begin VB.Label Label112 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total Credits"
            ForeColor       =   &H80000011&
            Height          =   195
            Index           =   3
            Left            =   30
            TabIndex        =   9
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label112 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total Debits"
            ForeColor       =   &H80000011&
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   8
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblDebits 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "99,999.99"
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   990
            TabIndex        =   7
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblCredits 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "99,999.99"
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   990
            TabIndex        =   6
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   945
         MaxLength       =   50
         TabIndex        =   4
         Top             =   90
         Width           =   3675
      End
      Begin VB.TextBox txtAccountingDate 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   230
         Left            =   6015
         MaxLength       =   50
         TabIndex        =   3
         Top             =   90
         Width           =   1605
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label112 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Accounting Date"
         Height          =   195
         Index           =   1
         Left            =   4755
         TabIndex        =   10
         Top             =   120
         Width           =   1200
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   1050
      Width           =   10425
      _cx             =   18389
      _cy             =   7223
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
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FJournalEntry.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   1005
      ButtonWidth     =   873
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   4350
         Top             =   -30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   55
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":016C
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":0A46
               Key             =   "RFP"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1320
               Key             =   "CreateJob"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1BFA
               Key             =   "takeoffsettings"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":24D4
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":2DAE
               Key             =   "DecreaseDecimals"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":3688
               Key             =   "IncreaseDecimals"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":3F62
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":483C
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":5116
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":59F0
               Key             =   "UpdatePrices"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":62CA
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":6BA4
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":747E
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":7D58
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":8632
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":8F0C
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":97E6
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":A0C0
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":A99A
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":B274
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":BB4E
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":C428
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":CD02
               Key             =   "New"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":D5DC
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":DEB6
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":E790
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":F06A
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":F944
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1021E
               Key             =   "AddPricelist"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":10AF8
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":113D2
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":11CAC
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":12586
               Key             =   "PricelistExport"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":12E60
               Key             =   "PricelistImport"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1373A
               Key             =   "NewPricelist"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":14014
               Key             =   "View"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":148EE
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":151C8
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":15AA2
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1637C
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":16C56
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":17530
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":17E0A
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":186E4
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":18FBE
               Key             =   "Models"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":19898
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1A172
               Key             =   "Generate"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1AA4C
               Key             =   "Timberline"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1B326
               Key             =   "MBImport"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1BC00
               Key             =   "EEEstimating"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1C4DA
               Key             =   "EEExport"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1CDB4
               Key             =   "EEImport"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1D68E
               Key             =   "MBExport"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJournalEntry.frx":1DF68
               Key             =   "MasterBuilder"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FJournalEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FJournalEntries"
Private Dirty As Boolean

Private mEntryID As Long


Private Sub Form_Load()
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    mEntryID = 0
    Call LoadData
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    gData.Move 0, pHeading.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - pHeading.Height - Toolbar.Height
    frmTotals.Move Max(7620, Me.ScaleWidth - frmTotals.Width)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyS And Shift = vbCtrlMask:       Call SaveData(False)
        Case KeyCode = vbKeyEscape:                         Unload Me
    End Select
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
    .ComboList = ""
    Select Case .ColKey(Col)
        Case "TranDate", "Job", "Extra", "CostCode", "Category", "Account"
            .ComboList = "|..."
        Case "Description", "DebitAmount", "CreditAmount"
        Case Else: Cancel = True
    End Select
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With gData
    Select Case .ColKey(Col)
        Case "TranDate"
        Case "Job"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Job", "select job_no Job,Description from tblJobs where isnull(inactive,0)=0") Then
                .TextMatrix(Row, Col) = FPickList.SelectedItem("Job")
            End If
        Case "Extra"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Extra", "select Extra from jobextras where job=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("job")))) Then
                .TextMatrix(Row, Col) = FPickList.SelectedItem("Extra")
            End If
        Case "CostCode"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Cost Code", "select CostCode,Description from standardcostcodes") Then
                .TextMatrix(Row, Col) = FPickList.SelectedItem("CostCode")
            End If
        Case "Category"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Category", "select Category,Description from standardcategories") Then
                .TextMatrix(Row, Col) = FPickList.SelectedItem("Category")
            End If
        Case "Account"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Account", "select Account,Description from GLAccounts") Then
                .TextMatrix(Row, Col) = FPickList.SelectedItem("Account")
            End If
    End Select
    End With
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    With gData
        s = .EditText
        Select Case .ColKey(Col)
            Case "TranDate":   Cancel = s
            Case "Job":        Cancel = Not ValidateItem(s, "Job not found", "select job from tbljobs where job=" & DbQuote(Str, .EditText))
            Case "Extra":      Cancel = Not ValidateItem(s, "Extra not found", "select extra from jobextras where job=" & DbQuote(Str, .TextMatrix(r, .ColIndex("job"))) & " And extra=" & DbQuote(Str, .EditText))
            Case "CostCode":   Cancel = Not ValidateItem(s, "Cost Code not found", "select costcode from standardcostcodes where costcode=" & DbQuote(Str, .EditText))
            Case "Category":   Cancel = Not ValidateItem(s, "Category not found", "select category from standardcategories where category=" & DbQuote(Str, .EditText))
            Case "Account":    Cancel = Not ValidateItem(s, "Account not found", "SELECT account FROM glaccounts WHERE account=" & DbQuote(Str, .EditText))
        End Select
        If Not Cancel Then
            .EditText = s
            If Row = .Rows - 1 Then
                .RowData(r) = "NEW"
                .AddItem ""
            End If
            If .RowData(r) = "" <> "NEW" Then .RowData(r) = "DIRTY"
        End If
    End With
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim r As Long
    Dim s As String
    Select Case UCase(Trim(Button.Key))
        
        Case "OPEN"
            If SaveData(True) Then
                s = "select EntryID,AccountingDate,Description from journalentries"
                If Not FPickList.Choose(HFApp.Databases(dbHomeFront), "Journal Entry", s, , , , , "EntryID") Then Exit Sub
                mEntryID = FPickList.SelectedItem("EntryID")
                Call LoadData
            End If
        
        Case "NEW"
            If SaveData(True) Then
                mEntryID = 0
                Call LoadData
            End If
            
        Case "SAVE"
            Call SaveData(False)
    End Select
End Sub





Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh:
    Dim rc As Long
    Dim s As String
    Dim r As Long
    
    If Not Dirty Then
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
    
    If mEntryID = 0 Then
        s = ""
        s = s & "insert into journalentries(description,accountingdate)" & vbCrLf
        s = s & "values(" & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "      ," & DbQuote(Date, txtAccountingDate.Text) & ")"
        Call HFApp.SqlExec(s, dbHomeFront)
        mEntryID = HFApp.SqlIdentity("journalentries", dbHomeFront)
    Else
        s = ""
        s = s & "update journalentries" & vbCrLf
        s = s & "set description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "   ,accountingdate=" & DbQuote(Date, txtAccountingDate.Text) & vbCrLf
        s = s & "where entryid=" & DbQuote(Num, mEntryID) & vbCrLf
        Call HFApp.SqlExec(s, dbHomeFront)
    End If
    
    
        
        
    With gData
        For r = .Rows - 2 To 1 Step -1
            If .RowHidden(r) And .ValueMatrix(r, .ColIndex("TranID")) <> 0 Then
                s = "delete from journalentrylines where tranid=" & DbQuote(Str, .ValueMatrix(r, .ColIndex("TranID")))
                Call HFApp.SqlExec(s, dbHomeFront)
            End If
        Next
    
        For r = 1 To .Rows - 1
            If .RowData(r) = "DIRTY" Then
                s = ""
                s = s & "UPDATE journalentrylines" & vbCrLf
                s = s & "   SET TranDate=" & DbQuote(Date, .TextMatrix(r, .ColIndex("TranDate"))) & vbCrLf
                s = s & "      ,Description=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                s = s & "      ,Job=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Job"))) & vbCrLf
                s = s & "      ,Extra=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Extra"))) & vbCrLf
                s = s & "      ,CostCode=" & DbQuote(Str, .TextMatrix(r, .ColIndex("CostCode"))) & vbCrLf
                s = s & "      ,Category=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Category"))) & vbCrLf
                s = s & "      ,Account=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Account"))) & vbCrLf
                s = s & "      ,DebitAmount=" & DbQuote(Num, .TextMatrix(r, .ColIndex("DebitAmount"))) & vbCrLf
                s = s & "      ,CreditAmount=" & DbQuote(Num, .TextMatrix(r, .ColIndex("CreditAmount"))) & vbCrLf
                s = s & "      ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
                s = s & "      ,TStmp=getdate()" & vbCrLf
                s = s & "WHERE TranID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("TranID")))
                Call HFApp.SqlExec(s, dbHomeFront)
                .RowData(r) = ""
            End If
            
            If .RowData(r) = "NEW" Then
                s = ""
                s = s & "INSERT journalentrylines(TranDate,Description,Job,Extra,CostCode,Category,Account,DebitAmount,CreditAmount,UStmp,TStmp" & vbCrLf
                s = s & "VALUES(" & DbQuote(Date, .TextMatrix(r, .ColIndex("TranDate"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Job"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Extra"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("CostCode"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Category"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Account"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("DebitAmount"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("CreditAmount"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
                s = s & "      ,getdate()" & ")"
                Call HFApp.SqlExec(s, dbHomeFront)
                .TextMatrix(r, .ColIndex("TranID")) = HFApp.SqlIdentity("journalentrylines", dbHomeFront)
                .RowData(r) = ""
            End If
        Next
    End With
    
    Dirty = False
    SaveData = True
    Screen.MousePointer = vbDefault

Exit Function
eh: Call errHandler(SRCFILE & "SaveData")
End Function


Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    
    Set rs = HFApp.SqlExec("SELECT * FROM JournalEntries WHERE EntryID=" & DbQuote(Num, mEntryID))
    If rs.EOF Then
        txtDescription.Text = ""
        txtAccountingDate.Text = ""
        lblDebits = "0.00"
        lblCredits = "0.00"
    Else
        txtDescription.Text = "" & rs("Description")
        txtAccountingDate.Text = Format("" & rs("Description"), "short date")
    End If
        
        
    With gData
        .Redraw = flexRDNone
        .Rows = 1
        Set rs = HFApp.SqlExec("SELECT * FROM JournalEntryLines WHERE EntryID=" & DbQuote(Num, mEntryID))
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            .TextMatrix(r, .ColIndex("TranID")) = "" & rs("TranID")
            .TextMatrix(r, .ColIndex("TranDate")) = Format("" & rs("TranDate"), "Medium Date")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("Job")) = "" & rs("Job")
            .TextMatrix(r, .ColIndex("Extra")) = "" & rs("Extra")
            .TextMatrix(r, .ColIndex("CostCode")) = "" & rs("CostCode")
            .TextMatrix(r, .ColIndex("Category")) = "" & rs("Category")
            .TextMatrix(r, .ColIndex("Account")) = "" & rs("Account")
            .TextMatrix(r, .ColIndex("DebitAmount")) = "" & rs("DebitAmount")
            .TextMatrix(r, .ColIndex("CreditAmount")) = "" & rs("CreditAmount")
            rs.MoveNext
        Wend
        .AddItem ""
        .Redraw = flexRDBuffered
    End With
    Dirty = False

End Sub

