VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FCommunities 
   Caption         =   "Communities"
   ClientHeight    =   5520
   ClientLeft      =   1290
   ClientTop       =   2625
   ClientWidth     =   12750
   Icon            =   "FCommunities.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   12750
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   11865
      _cx             =   20929
      _cy             =   3201
      Appearance      =   2
      BorderStyle     =   0
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   47
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FCommunities.frx":000C
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
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   12750
      _ExtentX        =   22490
      _ExtentY        =   1058
      ButtonWidth     =   1455
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    New    "
            Key             =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   4260
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   39
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":0721
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":0FFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":18D5
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":21AF
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":2A89
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":3363
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":3C3D
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":4517
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":4DF1
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":56CB
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":5FA5
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":687F
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":7159
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":7A33
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":830D
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":8BE7
               Key             =   "New"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":94C1
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":9D9B
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":A675
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":AF4F
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":B829
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":C103
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":C9DD
               Key             =   "PricebookSearch"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":D2B7
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":DB91
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":E46B
               Key             =   "PricebookExport"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":ED45
               Key             =   "PricebookImport"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":F61F
               Key             =   "PricebookNew"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":FEF9
               Key             =   "View"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":107D3
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":110AD
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":11987
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":12261
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":1257B
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":12E55
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":1372F
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":14009
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":148E3
               Key             =   "Models and Options"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCommunities.frx":151BD
               Key             =   "Generate"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FCommunities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FCommunities::"
Private mDirty As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyS And Shift = vbCtrlMask:       Call SaveData(False)
        Case KeyCode = vbKeyEscape:                         Unload Me
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo eh
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    
    
    Call LoadCustomDescriptions
    Call LoadData
    
    If HFApp.Options(Show_CDN_GST) = False Then
        gData.ColHidden(gData.ColIndex("GSTRate")) = True
    End If
    If HFApp.Options(UsePst) = False Then
        gData.ColHidden(gData.ColIndex("PSTRate")) = True
    End If
    
    If HFApp.Options(AccountingSystem) <> asD365_ABN Then
        gData.ColHidden(gData.ColIndex("ABN_D02Function")) = True
    End If
    
    
    
Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height
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


Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh:
    Dim i As Long
    Dim r As Long
    Dim s As String
    Dim rs As Recordset
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
    
    With gData
    
        For r = .Rows - 1 To 2 Step -1
            If .RowData(r) = "DELETE" Then
                s = "exec sp_DeleteCommunity " & DbQuote(Str, .TextMatrix(r, .ColIndex("Area")))
                Call HFApp.SqlExec(s, dbHomefront)
                Call .RemoveItem(r)
            End If
        Next
        
        For r = 2 To .Rows - 1

            .Row = r
            Call .ShowCell(r, 0)
            .Refresh

            If .RowData(r) = "DIRTY" Then

                If Trim(.TextMatrix(r, .ColIndex("Area"))) = "" Then
                    Screen.MousePointer = vbDefault
                    Call MsgBox("The community field cannot be blank", vbInformation, mProductName)
                    Exit Function
                End If
                
                If .Cell(flexcpData, r, .ColIndex("Area")) = "" Then
                    s = "INSERT INTO tblLocality(Area) VALUES(" & DbQuote(Str, .TextMatrix(r, .ColIndex("Area"))) & ")"
                    Call HFApp.SqlExec(s, dbHomefront)
                    s = "insert into DivisionCommunities(DivisionID,Community) select " & HFApp.DivisionID & ",area from tbllocality a left outer join DivisionCommunities dc on dc.Community =a.Area and dc.DivisionID = " & HFApp.DivisionID & "  where a.area = " & DbQuote(Str, .TextMatrix(r, .ColIndex("Area"))) & " and dc.community is null"
                    Call HFApp.SqlExec(s, dbHomefront)
                    .Cell(flexcpData, r, .ColIndex("Area")) = .TextMatrix(r, .ColIndex("Area"))
                End If
                
                s = ""
                s = s & "UPDATE tblLocality" & vbCrLf
                s = s & "SET Description=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                s = s & "   ,Comments=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Comments"))) & vbCrLf
                s = s & "   ,CompanyName=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Company"))) & vbCrLf
                s = s & "   ,Address1=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Address1"))) & vbCrLf
                s = s & "   ,Address2=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Address2"))) & vbCrLf
                s = s & "   ,Email=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Email"))) & vbCrLf
                s = s & "   ,City=" & DbQuote(Str, .TextMatrix(r, .ColIndex("City"))) & vbCrLf
                s = s & "   ,Province=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Province"))) & vbCrLf
                s = s & "   ,Zip=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Postal"))) & vbCrLf
                s = s & "   ,Country=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Country"))) & vbCrLf
                s = s & "   ,County=" & DbQuote(Str, .TextMatrix(r, .ColIndex("County"))) & vbCrLf
                s = s & "   ,Phone=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phone"))) & vbCrLf
                s = s & "   ,Fax=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Fax"))) & vbCrLf
                s = s & "   ,GSTNumber=" & DbQuote(Str, .TextMatrix(r, .ColIndex("TaxNumber"))) & vbCrLf
                s = s & "   ,Prefix1=" & DbQuote(Str, .TextMatrix(r, .ColIndex("BalanceSheet"))) & vbCrLf
                s = s & "   ,Prefix2=" & DbQuote(Str, .TextMatrix(r, .ColIndex("IncomePrefix"))) & vbCrLf
                
                s = s & "   ,ABN_D02Function=" & DbQuote(Str, .TextMatrix(r, .ColIndex("ABN_D02Function"))) & vbCrLf
                
                s = s & "   ,mortgage_credit=" & DbQuote(Str, .TextMatrix(r, .ColIndex("mortgage_credit"))) & vbCrLf
                s = s & "   ,BankAccount=" & DbQuote(Str, .TextMatrix(r, .ColIndex("BankAccount"))) & vbCrLf
                s = s & "   ,TarionBuilderNumber=" & DbQuote(Str, .TextMatrix(r, .ColIndex("TarionBuilderNumber"))) & vbCrLf
                s = s & "   ,PreconScheduleTemplate=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PreconScheduleTemplate"))) & vbCrLf
                
                s = s & "   ,lot_inv_debit=" & DbQuote(Str, .TextMatrix(r, .ColIndex("LotInventoryDebit"))) & vbCrLf
                s = s & "   ,lot_inv_credit=" & DbQuote(Str, .TextMatrix(r, .ColIndex("LotInventoryCredit"))) & vbCrLf
                s = s & "   ,project_manager=" & DbQuote(Str, .TextMatrix(r, .ColIndex("project_manager"))) & vbCrLf
                s = s & "   ,custserviceid=" & DbQuote(Str, .TextMatrix(r, .ColIndex("custserviceid"))) & vbCrLf
                s = s & "   ,dc_sales_person=" & DbQuote(Str, .TextMatrix(r, .ColIndex("dc_sales_person"))) & vbCrLf
                s = s & "   ,Contract_Document=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Contract_Document"))) & vbCrLf
                s = s & "   ,WarrantyJob=" & DbQuote(Str, .TextMatrix(r, .ColIndex("WarrantyJob"))) & vbCrLf
                s = s & "   ,Purchaser=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Purchaser"))) & vbCrLf
                s = s & "   ,WrapInsuranceExempt=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("WrapInsuranceExempt"))) & vbCrLf
                s = s & "   ,Inactive=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("Inactive"))) & vbCrLf
                s = s & "   ,LabourTaxGroup=" & DbQuote(Str, .TextMatrix(r, .ColIndex("LabourTaxGroup"))) & vbCrLf
                s = s & "   ,MaterialTaxGroup=" & DbQuote(Str, .TextMatrix(r, .ColIndex("MaterialTaxGroup"))) & vbCrLf
                s = s & "   ,SubContractTaxGroup=" & DbQuote(Str, .TextMatrix(r, .ColIndex("SubContractTaxGroup"))) & vbCrLf
                s = s & "   ,EquipmentTaxGroup=" & DbQuote(Str, .TextMatrix(r, .ColIndex("EquipmentTaxGroup"))) & vbCrLf
                s = s & "   ,OverheadTaxGroup=" & DbQuote(Str, .TextMatrix(r, .ColIndex("OverheadTaxGroup"))) & vbCrLf
                s = s & "   ,OtherTaxGroup=" & DbQuote(Str, .TextMatrix(r, .ColIndex("OtherTaxGroup"))) & vbCrLf
                s = s & "   ,GSTRate=" & DbQuote(Num, .TextMatrix(r, .ColIndex("GSTRate"))) & vbCrLf
                s = s & "   ,PSTRate=" & DbQuote(Num, .TextMatrix(r, .ColIndex("PSTRate")))
                s = s & "   ,Abr=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Abr")))
                s = s & "   ,SalesManagerEmail = " & DbQuote(Str, .TextMatrix(r, .ColIndex("SalesManagerEmail")))
                s = s & "   ,IntacctEntity = " & DbQuote(Str, .TextMatrix(r, .ColIndex("IntacctEntity")))
                s = s & "   ,IntacctParentJob= " & DbQuote(Str, .TextMatrix(r, .ColIndex("IntacctParentJob")))
                s = s & "   ,IntacctDepartment= " & DbQuote(Str, .TextMatrix(r, .ColIndex("IntacctDepartment")))
                s = s & "WHERE area=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Area"))) & vbCrLf
                Call HFApp.SqlExec(s)
                

                'remove other divisions
                s = "delete divisioncommunities where community=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Area"))) & vbCrLf
                Call HFApp.SqlExec(s)

                'update community divisions
                s = ""
                s = s & "insert into divisioncommunities(divisionid,community,buildproenabled)" & vbCrLf
                s = s & "select d.divisionid," & DbQuote(Str, .TextMatrix(r, .ColIndex("Area"))) & "," & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("BuildPro")) = flexChecked) & vbCrLf
                s = s & "from divisions d where divisioncode=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Divisions")))
                
                Call HFApp.SqlExec(s, dbHomefront)
            

                'clear dirty flag
                .RowData(r) = ""
            End If
        Next
    End With
    
    mDirty = False
    SaveData = True
    Screen.MousePointer = vbDefault

Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Screen.MousePointer = vbDefault
        Call MsgBox("This ID is taken. Please choose something else.", vbExclamation, App.ProductName)
    Else
        Call errHandler(SRCFILE & "SaveData", s)
        Screen.MousePointer = vbDefault
    End If
End Function


Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim i As Long
    Dim Area As String
    
    With gData
        .Redraw = flexRDNone
        .Rows = 1
        
        s = ""
        s = s & "SELECT l.BankAccount,l.Area, l.Description, l.Abr, l.Last_Customer, l.Contract_Document, l.Lot_Inv_Debit, l.Lot_Inv_Credit, l.Project_Manager, l.CustServiceID, l.DC_Sales_Person, l.last_job_no, l.developer, l.Prefix1"
        s = s & " ,l.Prefix2, l.CompanyName, l.CompanyLogo, l.Address1, l.Address2, l.City, l.Province, l.Zip, l.GSTNumber, l.Phone, l.FAX,l.Email, l.Est_DB_Path, l.WMS_DB_path, l.Sales_SystemID, l.EST_Phase_code, l.seq,l.PreconScheduleTemplate"
        s = s & " ,l.Comments, l.Inactive, l.ProjectCost, l.DepositGL, l.UsesPhases, l.Country, l.County, l.LabourTaxGroup, l.MaterialTaxGroup, l.SubContractTaxGroup, l.EquipmentTaxGroup, l.OverheadTaxGroup, l.TarionBuilderNumber" & vbCrLf
        s = s & " ,l.OtherTaxGroup, l.WarrantyJob, l.GSTRate, l.PSTRate, l.SalesManagerEmail, l.LastPOSeq,l.purchaser ,dc.BuildProEnabled BuildPro,l.WrapInsuranceExempt, l.IntacctEntity, l.IntacctParentJob,l.IntacctDepartment,l.Mortgage_Credit" & vbCrLf
        s = s & " ,ABN_D02Function" & vbCrLf
        s = s & "FROM tbllocality l " & vbCrLf
        s = s & "left outer join DivisionCommunities dc on dc.Community = l.Area and dc.DivisionID = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            
            'save original id
            .Cell(flexcpData, r, .ColIndex("Area")) = "" & rs("Area")
            
            .TextMatrix(r, .ColIndex("Inactive")) = "" & rs("Inactive")
            .TextMatrix(r, .ColIndex("BuildPro")) = "" & rs("BuildPro")
            .TextMatrix(r, .ColIndex("WrapInsuranceExempt")) = "" & rs("WrapInsuranceExempt")
            
            
            .TextMatrix(r, .ColIndex("ABN_D02Function")) = "" & rs("ABN_D02Function")
            
            .TextMatrix(r, .ColIndex("Area")) = "" & rs("Area")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("Comments")) = "" & rs("Comments")
            .TextMatrix(r, .ColIndex("Company")) = "" & rs("CompanyName")
            .TextMatrix(r, .ColIndex("Address1")) = "" & rs("Address1")
            .TextMatrix(r, .ColIndex("Address2")) = "" & rs("Address2")
            .TextMatrix(r, .ColIndex("City")) = "" & rs("City")
            .TextMatrix(r, .ColIndex("Abr")) = "" & rs("Abr")
            .TextMatrix(r, .ColIndex("Province")) = "" & rs("Province")
            .TextMatrix(r, .ColIndex("Postal")) = "" & rs("Zip")
            .TextMatrix(r, .ColIndex("Country")) = "" & rs("Country")
            .TextMatrix(r, .ColIndex("County")) = "" & rs("County")
            .TextMatrix(r, .ColIndex("Phone")) = "" & rs("Phone")
            .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
            .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
            .TextMatrix(r, .ColIndex("TaxNumber")) = "" & rs("GSTNumber")
            .TextMatrix(r, .ColIndex("WarrantyJob")) = "" & rs("WarrantyJob")
            .TextMatrix(r, .ColIndex("IncomePrefix")) = "" & rs("Prefix2")
            .TextMatrix(r, .ColIndex("BalanceSheet")) = "" & rs("Prefix1")
            .TextMatrix(r, .ColIndex("Mortgage_Credit")) = "" & rs("Mortgage_Credit")
            .TextMatrix(r, .ColIndex("BankAccount")) = "" & rs("BankAccount")
            .TextMatrix(r, .ColIndex("TarionBuilderNumber")) = "" & rs("TarionBuilderNumber")
            .TextMatrix(r, .ColIndex("LotInventoryDebit")) = "" & rs("Lot_Inv_Debit")
            .TextMatrix(r, .ColIndex("LotInventoryCredit")) = "" & rs("Lot_Inv_Credit")
            .TextMatrix(r, .ColIndex("Project_Manager")) = "" & rs("Project_Manager")
            .TextMatrix(r, .ColIndex("CustServiceID")) = "" & rs("CustServiceID")
            .TextMatrix(r, .ColIndex("DC_Sales_Person")) = "" & rs("DC_Sales_Person")
            .TextMatrix(r, .ColIndex("Contract_Document")) = "" & rs("Contract_Document")
            .TextMatrix(r, .ColIndex("PreconScheduleTemplate")) = "" & rs("PreconScheduleTemplate")
            
            
            .TextMatrix(r, .ColIndex("LabourTaxGroup")) = "" & rs("LabourTaxGroup")
            .TextMatrix(r, .ColIndex("MaterialTaxGroup")) = "" & rs("MaterialTaxGroup")
            .TextMatrix(r, .ColIndex("SubContractTaxGroup")) = "" & rs("SubContractTaxGroup")
            .TextMatrix(r, .ColIndex("EquipmentTaxGroup")) = "" & rs("EquipmentTaxGroup")
            .TextMatrix(r, .ColIndex("OverheadTaxGroup")) = "" & rs("OverheadTaxGroup")
            .TextMatrix(r, .ColIndex("OtherTaxGroup")) = "" & rs("OtherTaxGroup")
            .TextMatrix(r, .ColIndex("GSTRate")) = rs("GSTRate")
            .TextMatrix(r, .ColIndex("PSTRate")) = rs("PSTRate")
            
            .TextMatrix(r, .ColIndex("WarrantyJob")) = "" & rs("WarrantyJob")
            .TextMatrix(r, .ColIndex("SalesManagerEmail")) = "" & rs("SalesManagerEmail")
            .TextMatrix(r, .ColIndex("Purchaser")) = "" & rs("Purchaser")
            
            .TextMatrix(r, .ColIndex("IntacctEntity")) = "" & rs("IntacctEntity")
            .TextMatrix(r, .ColIndex("IntacctParentJob")) = "" & rs("IntacctParentJob")
            .TextMatrix(r, .ColIndex("IntacctDepartment")) = "" & rs("IntacctDepartment")
            
            rs.MoveNext
        Wend
        
        
        
        
        ' now load divisions
        s = ""
        s = s & "select l.area,d.divisioncode" & vbCrLf
        s = s & "  from tbllocality l " & vbCrLf
        s = s & "       left outer join divisioncommunities dc on(l.area=dc.community)" & vbCrLf
        s = s & "       left outer join divisions d on(d.divisionid=dc.divisionid)" & vbCrLf
        s = s & "order by 1,2" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        s = ""
        Area = ""
        While Not rs.EOF
            If Area <> "" & rs("area") And Area <> "" Then
                r = .FindRow(Area, , .ColIndex("Area"), False)
                If r > 0 Then .TextMatrix(r, .ColIndex("Divisions")) = Mid(s, 3)
                s = ""
                Area = ""
            End If
            Area = "" & rs("area")
            s = s & ", " & rs("divisioncode")
            rs.MoveNext
        Wend
        r = .FindRow(Area, , .ColIndex("Area"), False)
        If r > 0 Then .TextMatrix(r, .ColIndex("Divisions")) = Mid(s, 3)
        
        'add filter bar
        .AddItem "", 1
        .Cell(flexcpBackColor, 1, 0, 1, .Cols - 1) = FILTERBARBACKCOLOR
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = FILTERBARFORECOLOR
        .Cell(flexcpFontBold, 1, 0, 1, .Cols - 1) = FILTERBARFONTBOLD
        
        On Error Resume Next
        .FrozenRows = 2
        
        .Redraw = flexRDBuffered
    End With
    mDirty = False
    
    

End Sub


Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'apply filterbar
    If Row = gData.FrozenRows - 1 Then
        Call ApplyFilter(gData)
    End If

End Sub

Private Sub gData_AfterSort(ByVal Col As Long, Order As Integer)
    gData.FixedRows = 1
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        .ComboList = ""
        If .Row = 1 Then Exit Sub
        Select Case .ColKey(Col)
            Case "Area":                   .EditMaxLength = 10:  Cancel = .Cell(flexcpData, Row, .ColIndex("Area")) <> ""
            Case "Description":            .EditMaxLength = 50
            Case "Comments":               .ComboList = "..."
            Case "Company":                .ComboList = "..."
            Case "LotInventoryDebit":      .ComboList = "..."
            Case "Divisions":              .ComboList = "..."
            Case "Project_Manager":        .ComboList = "..."
            Case "Purchaser":              .ComboList = "..."
            Case "CustServiceID":          .ComboList = "..."
            Case "DC_Sales_Person":        .ComboList = "..."
            Case "Contract_Document":      .ComboList = "..."
            
            Case "ABN_D02Function":        .ComboList = "..."
            
            Case "PreconScheduleTemplate": .ComboList = HFApp.Options.ValueByName("ScheduleTemplates")
            
            Case "WarrantyJob":            .ComboList = "|...": .EditMaxLength = 12
        End Select
    End With
End Sub

Private Sub gData_BeforeSort(ByVal Col As Long, Order As Integer)
    gData.FixedRows = 2
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim i As Long
    Dim hidecols As String
    
    With gData
        Select Case .ColKey(Col)
        
            Case "WarrantyJob":
                hidecols = "id"
                s = "select job_no Job,Description,job_no ID from tblJobs where DivisionID = " & HFApp.DivisionID & " and isnull(inactive,0)=0"
            
            Case "Comments":
                s = .Text
                If FComments.Edit(s, gData, , , 8000) Then .Text = s
                s = ""
                
            Case "Contract_Document":
                s = .Text
                If VBGetOpenFileName(s, , , , , , "Contract Documents (*.dot;*.rpt)|*.dot;*.rpt", , , "Select Contract Document", , Me.hwnd) Then .Text = s
                s = ""
            
            Case "Project_Manager":     s = "select '' ID,'' Name union select pm id,pmname Name from tblprojectmanager where isnull(inactive,0)=0 and isnull(PrjMgr,0)=1"
            Case "Purchaser":           s = "select '' ID,'' Name union select pm id,pmname Name from tblprojectmanager where isnull(inactive,0)=0 and isnull(Purchaser,0)=1"
            Case "CustServiceID":       s = "select '' ID,'' Name union select custserviceid id ,custservicename Name from tblcustserviceperson"
            Case "DC_Sales_Person":     s = "select '' ID,'' Name union select sales_person_id id ,sales_person_name Name from tblsales_persons where isnull(inactive,0)=0 and isnull(dcsales,0)=1"
        
            Case "Company":             Call FCompany.ShowForm(gData)
            Case "LotInventoryDebit":   Call FCommunityAccounting.ShowForm(gData)
            
            Case "ABN_D02Function":
                s = "select Value,Description,value id from D365FinancialDimensionValues where Dimension='D02_Function' order by 1"
                hidecols = "id"
            
            Case "Divisions":
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Division", "select divisioncode Code,divisionname Name from divisions", , , , , , False) Then
                    s = ""
                    For i = 1 To FPickList.SelectedItems
                        s = s & ", " & FPickList.SelectedItem("Code", i)
                    Next
                    .TextMatrix(Row, Col) = Mid(s, 3)
                    s = ""
                End If
                
        End Select
        
        If s <> "" Then
            If FPickList.Choose(HFApp.Databases(dbHomefront), .TextMatrix(0, Col), s, .TextMatrix(Row, Col), , False, , hidecols) Then
                .TextMatrix(Row, Col) = FPickList.SelectedItem("id")
            End If
        End If
        
        mDirty = True
        .RowData(Row) = "DIRTY"
    End With
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim rs As Recordset
    If KeyCode = vbKeyDelete And Shift = vbCtrlMask And gData.Row > 1 Then
        Set rs = HFApp.SqlExec("select dbo.sp_CanDeleteCommunity(" & DbQuote(Str, gData.TextMatrix(gData.Row, gData.ColIndex("Area"))) & ")", dbHomefront)
        If "" & rs(0) = "False" Then
            MsgBox "You cannot delete a community that has customers, jobs or assemblies associated with it. ", vbInformation, "HomeFront"
            Exit Sub
        Else
            mDirty = True
            gData.RowData(gData.Row) = "DELETE"
            gData.RowHidden(gData.Row) = True
        End If
    End If
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As Recordset
    mDirty = True
    With gData
        Select Case .ColKey(Col)
            Case "WarrantyJob"
                If .EditText <> "" Then
                Set rs = HFApp.SqlExec("select * from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, .EditText))
                If rs.EOF Then
                    MsgBox "invalid job", vbExclamation, App.ProductName
                    Cancel = True
                    Exit Sub
                Else
                    .EditText = "" & rs("Job_no")
                End If
                End If
                
            Case "GSTRate"
                .EditText = Val(.EditText)
        End Select
        .RowData(Row) = "DIRTY"
    End With
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim r As Long
    Select Case UCase(Trim(Button.Key))
        Case "DELETE"
            Call gData_KeyDown(vbKeyDelete, vbCtrlMask)
            
        Case "NEW"
            With gData
                r = .Rows
                .AddItem ""
                .Row = .Rows - 1
                .RowData(r) = "DIRTY"
                On Error Resume Next
                .TextMatrix(r, .ColIndex("Divisions")) = "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & HFApp.DivisionID)(0)
                mDirty = True
            End With
        Case "SAVE"
            Call SaveData(False)
    End Select
End Sub
Public Function GetCustomDesc(ItemName As String, Optional Plural As Boolean = False) As String
On Error Resume Next
    Dim s As String
    s = ItemName
    s = HFApp.SqlExec("select isnull(nullif(custom_description,''),item) from customdescriptions where item=" & DbQuote(Str, ItemName), dbHomefront)(0)
    
    If Plural Then
        Select Case Right(s, 1)
            Case "s"
            Case "y":   s = left(s, Len(s) - 1) & "ies"
            Case Else:  s = s & "s"
        End Select
    End If
    
    GetCustomDesc = s
End Function


Private Sub LoadCustomDescriptions()
    Dim s As String
    
    s = GetCustomDesc("Community")
    Me.Caption = s
    gData.TextMatrix(0, gData.ColIndex("area")) = s
    
End Sub


