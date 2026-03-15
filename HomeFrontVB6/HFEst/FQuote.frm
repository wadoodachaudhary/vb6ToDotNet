VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FQuote 
   Caption         =   "Quote"
   ClientHeight    =   7950
   ClientLeft      =   330
   ClientTop       =   1710
   ClientWidth     =   12885
   Icon            =   "FQuote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   12885
   Begin VB.TextBox txtJob 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   0
      Text            =   " "
      Top             =   900
      Width           =   2865
   End
   Begin VB.TextBox txtNotes 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   4320
      MaxLength       =   8000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "FQuote.frx":000C
      Top             =   1170
      Width           =   8745
   End
   Begin VB.TextBox txtDescription 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1140
      Width           =   2865
   End
   Begin VB.ComboBox cboGLPrefix 
      Height          =   240
      Left            =   1320
      TabIndex        =   2
      Top             =   1380
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   423
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
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2655
      Left            =   4350
      TabIndex        =   4
      Top             =   2670
      Width           =   7695
      _cx             =   13573
      _cy             =   4683
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   62
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FQuote.frx":000E
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
      OleDragMode     =   1
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
      Begin VB.PictureBox picWarningMessages 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2040
         ScaleHeight     =   615
         ScaleWidth      =   3795
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   3795
         Begin VB.TextBox lblWarningMessages 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "FQuote.frx":098B
            Top             =   60
            Width           =   1095
         End
         Begin VB.Image imgTipIcon 
            Height          =   240
            Left            =   60
            Top             =   60
            Width           =   240
         End
         Begin VB.Shape shpWarningMessages 
            BorderColor     =   &H80000017&
            Height          =   555
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   3675
         End
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   0
         Picture         =   "FQuote.frx":0991
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgInfo 
         Height          =   240
         Left            =   0
         Picture         =   "FQuote.frx":0F1B
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   1005
      ButtonWidth     =   1799
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save As"
            Key             =   "SaveAs"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Attachments"
            Key             =   "Attachments"
            Style           =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "One Time"
            Key             =   "TakeoffOneTime"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Item"
            Key             =   "TakeoffItem"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Assembly"
            Key             =   "TakeoffAssembly"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Custom"
            Key             =   "TakeoffCustom"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Preview"
            Key             =   "Preview"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Rfrsh Costs"
            Key             =   "RePrice"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Create Job"
            Key             =   "GenerateJob"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Quote #"
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
      Left            =   495
      TabIndex        =   8
      Top             =   930
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Items"
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
      Left            =   4320
      TabIndex        =   7
      Top             =   2430
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Notes"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   930
      Width           =   510
   End
   Begin VB.Label Label112 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   405
      TabIndex        =   5
      Top             =   1170
      Width           =   795
   End
End
Attribute VB_Name = "FQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public EventTraps As Collection
Const srcFile = "FQuote::"

Private mFunnyFlag As Boolean ' used to ignore the first mouse move event see gitems_mousemove and gitems_rowcolchanged

Private mDirty         As Boolean
Private mJob           As String
Private mEstAssemblyID As Long

Const mShowPhases = False
Const mShowBudgets = True


Private Const mcITEM_COPYITEMS = 0
Private Const mcITEM_CHANGEITEM = 1
Private Const mcITEM_REMOVEITEMS = 2
Private Const mcITEM_VIEWFILES = 4
Private Const mcITEM_INSERTFILE = 5
Private Const mcITEM_LINKTOFILE = 6
Private Const mcITEM_CANCELBUDGETS = 8
Private Const mcITEM_SAVEONETIMETODB = 10
Private Const mcITEM_UPDATEPRICELIST = 11
Private Const mcITEM_COMPAREPRICES = 12

Private Property Get Dirty() As Boolean
    Dirty = mDirty
End Property
Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
    Toolbar.Buttons("Save").Enabled = mDirty
End Property

Public Property Get Job() As String
    Job = mJob
End Property

Private Sub cboCommunity_Click()
    Dirty = True
    Call LoadComboBox(cboPhase, HFApp.Databases(dbHomefront), "select distinct communityphase ,'',0 from tbljobs where community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & " and isnull(communityphase,'')<>'' ORDER BY 1")
End Sub

Private Sub cboGLPrefix_Click()
    Dirty = True
End Sub

Private Sub cboModel_Change()
    Dirty = True
End Sub

Private Sub cboModel_Click()
    Dirty = True

End Sub

Private Sub cboModel_Validate(Cancel As Boolean)
    If cboModel.Text <> left(cboModel.Text, 20) Then
        cboModel.Text = left(cboModel.Text, 20)
    End If
End Sub

Private Sub cboPhase_Click()
    Dirty = True
End Sub

Private Sub cboStatus_Click()
    Dirty = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbCtrlMask And KeyCode = vbKeyO:  Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case Shift = vbCtrlMask And KeyCode = vbKeyS:  Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub

Private Sub Form_Load()

    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gItems)
    cboStatus.Clear
    cboStatus.AddItem "In progress"
    cboStatus.AddItem "Closed"
    Call LoadCommunities
    Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), "select distinct isnull(model,'') ,'',0 from tbldbassemblymaster where assemblytype=0")
    Call LoadGLPrefixes
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    txtNotes.Width = Me.ScaleWidth - margin - txtNotes.left
    gProperties.Height = Me.ScaleHeight - margin - gProperties.Top
    gItems.Move gItems.left, gItems.Top, Me.ScaleWidth - gItems.left - margin, Me.ScaleHeight - gItems.Top - margin
End Sub


Private Sub gProperties_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
On Error GoTo eh

    With gProperties
        .ComboList = ""

        If col <> .ColIndex("Value") Then
            Cancel = True
            Exit Sub
        End If

        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt
            Case Num
            Case Cur
            Case Bit:      .ComboList = "Yes|No"
            Case DateTime: .ComboList = "|..."
            Case Str:      .EditMaxLength = .ValueMatrix(Row, .ColIndex("Length"))
                           .ComboList = .TextMatrix(Row, .ColIndex("PickList"))
            Case Else:     Cancel = True
        End Select
        
    End With

Exit Sub
eh: Call errHandler(srcFile & "gProperties_BeforeEdit")
End Sub

Private Sub gProperties_CellButtonClick(ByVal Row As Long, ByVal col As Long)
On Error GoTo eh
    With gProperties
        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt:
            Case Num:
            Case Cur:
            Case Bit:
            Case DateTime: Call DCalendar.Popup(gProperties, .RowPos(.Row) + .RowHeight(.Row), .colPos(.col))
            Case Str:
            Case Else:
        End Select
    End With
Exit Sub
eh: Call errHandler(srcFile & "gProperties_CellButtonClick")
End Sub

Private Sub gProperties_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gProperties
        If .Row < 0 Then Exit Sub
        Select Case KeyCode

            Case vbKeySpace
                .IsCollapsed(.Row) = IIf(.IsCollapsed(.Row) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
                
            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
                
            Case vbKeyDelete
                If .col = .ColIndex("value") Then
                    .Text = ""
                    Call gProperties_ValidateEdit(.Row, .col, False)
                    Dirty = True
                    
                    Select Case .ValueMatrix(.Row, .ColIndex("DataType"))
                        Case Cur, Num, NumInt:  .Text = "0"
                    End Select
                End If
            
            
            
            Case vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If
                
        End Select
    End With
End Sub

Private Sub gProperties_RowColChange()
On Error GoTo eh

    With gProperties
        If .Row >= 0 Then
            If .TextMatrix(.Row, .ColIndex("DataType")) = "" Then
                .col = .ColIndex("Name")
            Else
                .col = .ColIndex("Value")
            End If
        End If
    End With

Exit Sub
eh: Call errHandler(srcFile & "gProperties_ValidateEdit")
End Sub

Private Sub gProperties_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
On Error GoTo eh
    
    With gProperties
        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt:    .EditText = Int(Val(.EditText))
            Case Num:       .EditText = Val(.EditText)
            Case Cur:       .EditText = format(Round(Val("" & Replace(Replace(Replace(.EditText, "%", ""), ",", ""), "$", "")), 2), "Currency")
            Case Bit:
            Case DateTime:
                If IsDate(.EditText) Or .EditText = "" Then
                    .EditText = format(.EditText, HFApp.Options(DateFormat))
                Else
                    Cancel = True
                End If
            Case Str:
            Case Else:       Cancel = True
        End Select
    End With
    If Not Cancel Then Dirty = True
    
Exit Sub
eh: Call errHandler(srcFile & "gProperties_ValidateEdit")
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
        Case "View Files"
            Call Toolbar_ButtonClick(Toolbar.Buttons("Attachments"))
            
        Case "Insert File", "Link to File"
            Call FAttachments.AddFile("J~" & mJob, ButtonMenu.Text = "Insert File")
            
    End Select
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim bRePrice As Boolean
    Dim i As Long
    Dim rs As Recordset
    Dim s As String
    Dim quote As String
    Dim f As FEstimateItems
    Dim fj As FJob
    
    Dim CostBasis As CostBasisTypes
    Dim EffectiveDate As Date

    Select Case Button.Key

         Case "GenerateJob"
         
            'save quote
            If Not SaveData(False) Then Exit Sub

            'get job number
            s = InputBox(vbCrLf & vbCrLf & vbCrLf & "Enter Job Number", "Generate Job")
            If s = "" Then Exit Sub
            
            
            If Not ValidateJobNumber(s) Then
                Screen.MousePointer = vbDefault
                MsgBox "Incorrect job format." & vbCrLf & "Correct format is " & Replace(HFApp.Options(Job_Mask), "&", "x"), vbExclamation, App.ProductName
                Call SetCtrlFocus(txtJob)
                Exit Sub
            End If
            
            

            'save quote as job
            quote = txtJob.Text
            txtJob.Text = s
            mJob = Chr(1)
            For i = 1 To gItems.Rows - 1
                gItems.RowData(i) = "NEW"
            Next
            If SaveData(False, True) Then
                'open new job
                Unload Me
                Set fj = New FJob
                Call fj.EditJob(s)
            Else
                'restore quote number
                txtJob.Text = quote
                mJob = quote
            End If
            
         
         Case "SaveAs"
            s = InputBox(vbCrLf & vbCrLf & vbCrLf & "Enter a new quote number", "Save As")
            If s = "" Then Exit Sub
         
            txtJob.Enabled = True
            txtJob.Text = s
            mJob = Chr(1)
            For i = 1 To gItems.Rows - 1
                gItems.RowData(i) = "NEW"
            Next
            Call SaveData(False)
            

        Case "Attachments"
            Call FAttachments.Showform("File Attachments - " & txtJob.Text & " - " & txtDescription.Text, "J~" & mJob, "Job Documents")
        
        
        Case "Open"
            If Not SaveData(True) Then Exit Sub
            
            s = "SELECT j.Job_No Quote,j.Description,c.Description Community,j.CommunityPhase Phase,j.Municipal_Address Address,j.PM,j.Purchaser,j.Estimator FROM tblJobs j LEFT OUTER JOIN tblLocality c on(j.community=c.area) where j.isQuote=1"
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Quote", s, mJob, , True) Then
            
                mJob = FPickList.SelectedItem("Quote")
                If mJob = "" Then mJob = Chr(1) 'new job
                Set f = FindForm("FJob", mJob)
                
                If f Is Nothing Then
                    Call LoadData
                Else
                    If f.WindowState = vbMinimized Then f.WindowState = vbNormal
                    f.SetFocus
                    If Not Toolbar.Buttons(3).Enabled Then
                        'close empty screen
                        Unload Me
                    End If
                End If
            Else
                If mJob = "" Then mJob = Chr(1) 'new job
            End If
            
            
         Case "Save"
            Call SaveData(False)
            


        Case "RePrice"
            If Not SaveData(True) Then Exit Sub
            If mShowBudgets Then
                If 0 = HFApp.SqlExec("SELECT COUNT(*) FROM EstimatedItems WHERE BudgetGenerated=0 AND " & WhereClause)(0) Then
                    MsgBox "You have nothing selected that can be refreshed." & vbCrLf & _
                           "Use the check boxes to select some assemblies or" & vbCrLf & _
                           "items that have not already been generated.", vbExclamation, App.ProductName
                Else
                    bRePrice = FEstimateItemsRefreshCosts.Showform(CostBasis, EffectiveDate)
                    If bRePrice Then
                        
                        'convert costbasis to lookuptype
                        If CostBasis <> 0 Then CostBasis = CostBasis + 1
                    
                        'get item rates
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET BudgetOverridden=0" & vbCrLf
                        s = s & "      ,BudgetRate= case when ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,0,v.BudgetVendor," & DbQuote(Date, EffectiveDate) & "),0) <>0 then ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,0,v.BudgetVendor," & DbQuote(Date, EffectiveDate) & "),0)" & vbCrLf
                        s = s & "                        when 'True'=" & DbQuote(Str, "" & HFApp.Options(ZeroRateOnRefreshCosts)) & " and i.BudgetOverridden=0 then 0" & vbCrLf
                        s = s & "                        else i.BudgetRate end" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.BudgetDeleted=0" & vbCrLf
                        s = s & "   AND v.BudgetGenerated=0" & vbCrLf
                        s = s & "   AND " & WhereClause("v")
                        Call HFApp.SqlExec(s, dbHomefront, i)
                        
                        
                        're-extend prices and taxes
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET BudgetPretax = i.BudgetRate * i.BudgetQty" & vbCrLf
                        s = s & "      ,BudgetJCTax = i.BudgetRate * i.BudgetQty * i.BudgetJCTaxRate/100" & vbCrLf
                        s = s & "      ,BudgetNJCTax = i.BudgetRate * i.BudgetQty * i.BudgetNJCTaxRate/100" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.BudgetDeleted=0" & vbCrLf
                        s = s & "   AND v.BudgetGenerated=0" & vbCrLf
                        If Not HFApp.Options(ZeroRateOnRefreshCosts) Then
                            s = s & "   AND ISNULL(dbo.Purch_GetItemRate(" & CostBasis & ",0,v.Community,v.CommunityPhase,v.Assembly,'','',v.EstPhase,v.EstItem,0,v.BudgetVendor," & DbQuote(Date, EffectiveDate) & "),0)<>0" & vbCrLf
                        End If
                        s = s & "   AND " & WhereClause("v")
                        Call HFApp.SqlExec(s, dbHomefront)
                                                
                        'keep po values in synch
                        s = ""
                        s = s & "UPDATE EstimateItems" & vbCrLf
                        s = s & "   SET PORate=v.BudgetRate" & vbCrLf
                        s = s & "      ,POPretax=v.BudgetPretax" & vbCrLf
                        s = s & "      ,POJCTax=v.BudgetJCTax" & vbCrLf
                        s = s & "      ,PONJCTax=v.BudgetNJCTax" & vbCrLf
                        s = s & "      ,POOverridden=0" & vbCrLf
                        s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                        s = s & " WHERE v.BudgetRate IS NOT NULL" & vbCrLf
                        s = s & "   AND v.BudgetDeleted=0" & vbCrLf
                        s = s & "   AND v.BudgetGenerated=0" & vbCrLf
                        s = s & "   AND v.POGenBatch=0" & vbCrLf
                        s = s & "   AND " & WhereClause("v")
                        Call HFApp.SqlExec(s, dbHomefront)
                        
                        Dirty = False
                        Call LoadItems
                        MsgBox "Costs have been successfully refreshed." & vbCrLf & i & " items were updated.", vbInformation, App.ProductName
                    End If
                
                End If
            End If
            
        
        Case "TakeoffOneTime", "TakeoffItem", "TakeoffAssembly", "TakeoffCustom"
            Call FTakeoff.Takeoff(Me, Mid(Button.Key, 8), "", GetComboBoxListKey(cboCommunity), cboPhase.Text, "", "", mJob)
            Call ColorizeItems(-1)
            
        
        Case "Preview"
            
            
    End Select
    Call LogError(srcFile & "Toolbar_ButtonClick() closing", i, s)

Exit Sub
eh: Call errHandler(srcFile & "Toolbar_ButtonClick", s)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gItems)
End Sub

Public Function SaveData(prompt As Boolean, Optional CreateJob As Boolean = False) As Boolean
On Error GoTo eh
    
    Dim GLPrefix As String
    Dim s As String
    Dim Community As String
    Dim CommunityPhase As String
    Dim r As Long
    
    If Not Dirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox(Me.Caption & " has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    txtJob.Text = Trim(txtJob.Text)
    If Trim(txtJob.Text) = "" Then
        MsgBox "A job number is required.", vbExclamation, App.ProductName
        Call SetCtrlFocus(txtJob)
        Exit Function
    End If
    
    
    
    
    
    Screen.MousePointer = vbHourglass
    
    

    Community = GetComboBoxListKey(cboCommunity)
    CommunityPhase = cboPhase.Text
    GLPrefix = IIf(cboGLPrefix.Visible, GetComboBoxListKey(cboGLPrefix), txtGLPrefix.Text)
    
    
    'save job
    With gProperties
        If mJob = Chr(1) Then
            
            s = ""
            s = s & "INSERT INTO tblJobs(isquote,Job_No,Description,Notes,GL_Prefix,Community,CommunityPhase,municipal_address,city,province,zip,sitephone,sitefax,Inactive" & vbCrLf
            s = s & "                   ,Lot,Block,LotPlan,Model" & vbCrLf
            s = s & "                   ,LabourTaxGroup,MaterialTaxGroup,SubContractTaxGroup,EquipmentTaxGroup,OverheadTaxGroup,OtherTaxGroup)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Bit, Not CreateJob) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtNotes.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, GLPrefix) & vbCrLf
            s = s & "      ," & DbQuote(Str, Community) & vbCrLf
            s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtAddress.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtCity.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtProvince.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtPostal.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtPhone.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtFax.Text) & vbCrLf
            s = s & "      ," & DbQuote(Bit, cboStatus.ListIndex = 1) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(1, .ColIndex("value"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(2, .ColIndex("value"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(3, .ColIndex("value"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(4, .ColIndex("value"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(5, .ColIndex("value"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(6, .ColIndex("value"))) & ")"
            HFApp.SqlExec s
            mJob = txtJob.Text
            Me.tag = mJob
            txtJob.Enabled = False
            Toolbar.Buttons("TakeoffOneTime").Enabled = True
            Toolbar.Buttons("TakeoffItem").Enabled = True
            Toolbar.Buttons("TakeoffAssembly").Enabled = True
            Toolbar.Buttons("TakeoffCustom").Enabled = True
            Toolbar.Buttons("SaveAs").Enabled = True
            Toolbar.Buttons("Attachments").Enabled = True
        
            s = ""
            s = s & "INSERT INTO tblcustomers (Customer_No,Job_no,Description,bal_sheet_Prefix,community,phase,address1,city,province,zip,phone,fax,purchased,approved,contract_assigned,cancelled,inactive,lot,block,lotplan,model,sale_posted)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, GLPrefix) & vbCrLf
            s = s & "      ," & DbQuote(Str, Community) & vbCrLf
            s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtAddress.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtCity.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtProvince.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtPostal.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtPhone.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtFax.Text) & ",1,1,1,0,0"
            s = s & "      ," & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "      ,1)"
            HFApp.SqlExec s
            
            s = ""
            s = s & "INSERT INTO EstimateAssemblies(job,Customer_no,EstimateIndex,HFDescription,AssemblyType,OptionType,salesqty)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtJob.Text) & vbCrLf
            s = s & "      ,0,'Quoted Items',-1,-1,1)" & vbCrLf
            HFApp.SqlExec s
            mEstAssemblyID = HFApp.SqlIdentity("EstimateAssemblies", dbHomefront)
        
        Else
            s = ""
            s = s & "UPDATE tblJobs" & vbCrLf
            s = s & "SET Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "   ,Notes=" & DbQuote(Str, txtNotes.Text) & vbCrLf
            s = s & "   ,GL_Prefix=" & DbQuote(Str, GLPrefix) & vbCrLf
            s = s & "   ,Community=" & DbQuote(Str, Community) & vbCrLf
            s = s & "   ,CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf
            s = s & "   ,Municipal_Address=" & DbQuote(Str, txtAddress.Text) & vbCrLf
            s = s & "   ,City=" & DbQuote(Str, txtCity.Text) & vbCrLf
            s = s & "   ,Province=" & DbQuote(Str, txtProvince.Text) & vbCrLf
            s = s & "   ,Zip=" & DbQuote(Str, txtPostal.Text) & vbCrLf
            s = s & "   ,SitePhone=" & DbQuote(Str, txtPhone.Text) & vbCrLf
            s = s & "   ,SiteFax=" & DbQuote(Str, txtFax.Text) & vbCrLf
            s = s & "   ,Inactive=" & DbQuote(Bit, cboStatus.ListIndex = 1) & vbCrLf
            s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "   ,Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
            s = s & "   ,LabourTaxGroup=" & DbQuote(Str, .TextMatrix(1, .ColIndex("value"))) & vbCrLf
            s = s & "   ,MaterialTaxGroup=" & DbQuote(Str, .TextMatrix(2, .ColIndex("value"))) & vbCrLf
            s = s & "   ,SubContractTaxGroup=" & DbQuote(Str, .TextMatrix(3, .ColIndex("value"))) & vbCrLf
            s = s & "   ,EquipmentTaxGroup=" & DbQuote(Str, .TextMatrix(4, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OverheadTaxGroup=" & DbQuote(Str, .TextMatrix(5, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OtherTaxGroup=" & DbQuote(Str, .TextMatrix(6, .ColIndex("value"))) & vbCrLf
            s = s & "WHERE Job_No=" & DbQuote(Str, txtJob.Text) & vbCrLf
            HFApp.SqlExec s
            
            
            If HFApp.Options(SalesSystem) = SalesSystems.asNone Then
                s = ""
                s = s & "UPDATE tblcustomers" & vbCrLf
                s = s & "SET Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
                s = s & "   ,bal_sheet_Prefix=" & DbQuote(Str, GLPrefix) & vbCrLf
                s = s & "   ,Community=" & DbQuote(Str, Community) & vbCrLf
                s = s & "   ,Phase=" & DbQuote(Str, CommunityPhase) & vbCrLf
                s = s & "   ,Address1=" & DbQuote(Str, txtAddress.Text) & vbCrLf
                s = s & "   ,City=" & DbQuote(Str, txtCity.Text) & vbCrLf
                s = s & "   ,Province=" & DbQuote(Str, txtProvince.Text) & vbCrLf
                s = s & "   ,Zip=" & DbQuote(Str, txtPostal.Text) & vbCrLf
                s = s & "   ,Phone=" & DbQuote(Str, txtPhone.Text) & vbCrLf
                s = s & "   ,Fax=" & DbQuote(Str, txtFax.Text) & vbCrLf
                s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
                s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
                s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
                s = s & "   ,Model=" & DbQuote(Str, cboModel.Text) & vbCrLf
                s = s & "WHERE Customer_No=" & DbQuote(Str, txtJob.Text)
                HFApp.SqlExec s
            End If
            
        End If
        
    
        'save properties
        s = ""
        For r = 7 To .Rows - 1
            If .TextMatrix(r, .ColIndex("DataType")) <> "" Then
                s = s & "," & vbQuote & .TextMatrix(r, .ColIndex("Name")) & vbQuote & "=" & DbQuote(.ValueMatrix(r, .ColIndex("DataType")), .TextMatrix(r, .ColIndex("Value")))
            End If
        Next
        If s <> "" Then
            On Error Resume Next
            Call HFApp.SqlExec("INSERT INTO JobCustomFields(Job_No) VALUES(" & DbQuote(Str, mJob) & ")", dbHomefront)
            On Error GoTo eh
            Call HFApp.SqlExec("UPDATE JobCustomFields SET " & Mid(s, 2) & " WHERE Job_No=" & DbQuote(Str, mJob), dbHomefront)
        End If
    End With
    
    
    Call SaveItems
    
    SaveData = True
    Dirty = False
    Screen.MousePointer = vbDefault
    
Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Screen.MousePointer = vbDefault
        If CreateJob Then
            MsgBox "Unable to create job " & vbQuote & txtJob.Text & vbQuote & ". It has already exists." & vbCrLf & vbCrLf & "Please use a different job number.", vbExclamation, App.ProductName
        Else
            MsgBox "Unable to save this quote. " & vbQuote & txtJob.Text & vbQuote & " has already been used." & vbCrLf & vbCrLf & "Please enter a different number.", vbExclamation, App.ProductName
        End If
    Else
        Call errHandler(srcFile & "SaveData", s)
    End If
End Function


Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim b As Boolean
    
    Me.tag = mJob
    
    
    If mJob = Chr(1) Then 'newjob
        txtJob.Enabled = True
        b = True
    Else
        txtJob.Enabled = False
        s = ""
        s = s & "select j.job_no,j.Notes,j.description,j.legal_address,j.gl_prefix" & vbCrLf
        's = s & "      ,j.arcustomer,isnull(nullif(c.description,''),j.arcustomer) arcustomerdesc" & vbCrLf
        s = s & "      ,j.community,j.communityphase,j.model,j.lot,j.block,j.lotplan,j.inactive,j.municipal_address,j.city,j.province,j.zip,j.sitephone" & vbCrLf
        s = s & "      ,j.sitefax" & vbCrLf
        s = s & "  from tblJobs j" & vbCrLf ' left outer join arcustomers c on(j.arcustomer=c.arcustomer)" & vbCrLf
        s = s & " where job_no=" & DbQuote(Str, mJob)
        Set rs = HFApp.SqlExec(s)
        b = Not rs.EOF
    End If
    
    txtDescription.Enabled = b
    txtNotes.Enabled = b
    cboGLPrefix.Enabled = b
    txtGLPrefix.Enabled = b
    txtAddress.Enabled = b
    txtCity.Enabled = b
    txtProvince.Enabled = b
    txtPostal.Enabled = b
    txtPhone.Enabled = b
    txtFax.Enabled = b
    cboStatus.Enabled = b
    cboModel.Enabled = b
    txtLot.Enabled = b
    txtBlock.Enabled = b
    txtLotPlan.Enabled = b
    gProperties.Enabled = b
    
    Toolbar.Buttons("TakeoffOneTime").Enabled = mJob <> Chr(1)
    Toolbar.Buttons("TakeoffItem").Enabled = mJob <> Chr(1)
    Toolbar.Buttons("TakeoffAssembly").Enabled = mJob <> Chr(1)
    Toolbar.Buttons("TakeoffCustom").Enabled = mJob <> Chr(1)
    Toolbar.Buttons("SaveAs").Enabled = mJob <> Chr(1)
    Toolbar.Buttons("Attachments").Enabled = mJob <> Chr(1)
    
    Toolbar.Buttons("Preview").Enabled = mJob <> Chr(1)
    Toolbar.Buttons("RePrice").Enabled = mJob <> Chr(1)
    Toolbar.Buttons("GenerateJob").Enabled = mJob <> Chr(1)
    
    If mJob = Chr(1) Or Not b Then
        SetCtrlFocus txtJob
        Me.Caption = "Job Setup"
        txtJob.Text = ""
        txtDescription.Text = ""
        txtNotes.Text = ""
        txtGLPrefix.Text = ""
        cboGLPrefix.ListIndex = -1
        txtAddress.Text = ""
        txtCity.Text = ""
        txtProvince.Text = ""
        txtPostal.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        cboModel.Text = ""
        txtLot.Text = ""
        txtBlock.Text = ""
        txtLotPlan.Text = ""
        cboStatus.ListIndex = 0
        cboCommunity.ListIndex = 0
    Else
        SetCtrlFocus txtDescription
        txtJob.Text = "" & rs("Job_No")
        Me.Caption = "Job Setup - " & txtJob.Text
        txtDescription.Text = "" & rs("Description")
        txtNotes.Text = "" & rs("Notes")
                       
        Call SetComboBoxListIndex(cboGLPrefix, , "" & rs("gl_prefix"))
        txtGLPrefix.Text = "" & rs("gl_prefix")
        txtAddress.Text = "" & rs("municipal_address")
        txtCity.Text = "" & rs("city")
        txtProvince.Text = "" & rs("province")
        txtPostal.Text = "" & rs("zip")
        txtPhone.Text = FormatPhone("" & rs("sitephone"))
        txtFax.Text = FormatPhone("" & rs("sitefax"))
        cboModel.Text = "" & rs("Model")
        txtLot.Text = "" & rs("Lot")
        txtBlock.Text = "" & rs("Block")
        txtLotPlan.Text = "" & rs("LotPlan")
        
        cboStatus.ListIndex = IIf("" & rs("Inactive") <> "True", 0, 1)
        Call SetComboBoxListIndex(cboCommunity, , "" & rs("Community"))
        cboPhase.Text = "" & rs("CommunityPhase")

    End If

    'load properties
    Call LoadProperties
        
    'load items
    Call LoadItems
    
    
    
    Dirty = False
    
End Sub



Private Sub txtBlock_Change()
    Dirty = True
End Sub

Private Sub txtBlock_GotFocus()
    SelectAll txtBlock
End Sub

Private Sub txtFax_Validate(Cancel As Boolean)
    txtFax.Text = FormatPhone(txtFax.Text)
End Sub

Private Sub txtGLPrefix_Change()
    Dirty = True
End Sub

Private Sub txtGLPrefix_GotFocus()
    SelectAll txtGLPrefix
End Sub

Private Sub txtJob_Change()
    Dirty = True
End Sub
Private Sub txtDescription_Change()
    Dirty = True
End Sub
Private Sub cboGLPrefix_Change()
    Dirty = True
End Sub
Private Sub txtAddress_Change()
    Dirty = True
End Sub
Private Sub txtCity_Change()
    Dirty = True
End Sub

Private Sub txtLot_Change()
    Dirty = True
End Sub

Private Sub txtLot_GotFocus()
    SelectAll txtLot
End Sub

Private Sub txtLotPlan_Change()
    Dirty = True
End Sub

Private Sub txtLotPlan_GotFocus()
    SelectAll txtLotPlan
End Sub

Private Sub txtNotes_Change()
    Dirty = True
End Sub

Private Sub txtNotes_GotFocus()
    SelectAll txtNotes
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    txtPhone.Text = FormatPhone(txtPhone.Text)
End Sub

Private Sub txtProvince_Change()
    Dirty = True
End Sub
Private Sub txtPostal_Change()
    Dirty = True
End Sub
Private Sub txtPhone_Change()
    Dirty = True
End Sub
Private Sub txtFax_Change()
    Dirty = True
End Sub


Private Sub txtJob_GotFocus()
    SelectAll txtJob
End Sub
Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub
Private Sub txtAddress_GotFocus()
    SelectAll txtAddress
End Sub
Private Sub txtCity_GotFocus()
    SelectAll txtCity
End Sub
Private Sub txtProvince_GotFocus()
    SelectAll txtProvince
End Sub
Private Sub txtPostal_GotFocus()
    SelectAll txtPostal
End Sub
Private Sub txtPhone_GotFocus()
    SelectAll txtPhone
End Sub
Private Sub txtFax_GotFocus()
    SelectAll txtFax
End Sub

Private Sub LoadCommunities()
    Call LoadComboBox(cboCommunity, HFApp.Databases(dbHomefront), "SELECT c.Description,c.Area,0 FROM tblLocality c ORDER BY 1")
End Sub



Private Sub LoadProperties()
On Error GoTo eh
    Dim rs As Recordset
    Dim r As Long
    Dim Category As String
    Dim FieldName As String
    Dim fields As Recordset
    Dim data   As Recordset
    Dim TaxGroups As String

    With gProperties
        r = -1
        .Rows = 0
        
        
        'add built in properties first
        Category = "Tax Groups"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = FMain.SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True
        
        
        TaxGroups = ""
        Set rs = HFApp.SqlExec("select taxgroup,description from TaxGroups order by 1")
        While Not rs.EOF
            TaxGroups = TaxGroups & "|" & rs(0) & vbTab & rs(1)
            rs.MoveNext
        Wend
        TaxGroups = Mid(TaxGroups, 2)
On Error Resume Next
        Set rs = HFApp.SqlExec("select * from tblJobs where job_no=" & DbQuote(Str, mJob))
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        .TextMatrix(r, .ColIndex("name")) = "Labour"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("LabourTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        .TextMatrix(r, .ColIndex("name")) = "Material"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("MaterialTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
                    
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        .TextMatrix(r, .ColIndex("name")) = "Subcontract"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("SubContractTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        .TextMatrix(r, .ColIndex("name")) = "Equipment"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("EquipmentTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        .TextMatrix(r, .ColIndex("name")) = "Overhead"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("OverheadTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        .TextMatrix(r, .ColIndex("name")) = "Other"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("OtherTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
On Error GoTo eh
            
            
        'now load user defined fields
        Set fields = HFApp.SqlExec("SELECT * FROM JobCustomFieldDefs ORDER BY Category,Name")
        Set data = HFApp.SqlExec("SELECT * FROM JobCustomFields WHERE Job_No=" & DbQuote(Str, mJob))
        
        While Not fields.EOF
            
            If Category <> "" & fields("category") Then
                Category = "" & fields("category")
                r = r + 1
                .AddItem ""
                .TextMatrix(r, .ColIndex("category")) = Category
                .TextMatrix(r, .ColIndex("name")) = IIf(Category = "", "unclassified", Category)
                Set .Cell(flexcpPicture, r, .ColIndex("name")) = FMain.SmallIcons.ListImages("category").Picture
                .RowOutlineLevel(r) = 0
                .IsSubtotal(r) = True
            End If
            
            r = r + 1
            .AddItem ""
            FieldName = "" & fields("Name")
            .TextMatrix(r, .ColIndex("category")) = Category
            .TextMatrix(r, .ColIndex("name")) = FieldName
            .TextMatrix(r, .ColIndex("PickList")) = "" & fields("PickList")
            Select Case data(FieldName).Type
                Case adInteger, adTinyInt, adSmallInt, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt: .TextMatrix(r, .ColIndex("DataType")) = NumInt
                Case adDouble, adSingle, adDecimal, adNumeric:                                                                           .TextMatrix(r, .ColIndex("DataType")) = Num
                Case adCurrency:                                                                                                         .TextMatrix(r, .ColIndex("DataType")) = Cur
                Case adBoolean:                                                                                                          .TextMatrix(r, .ColIndex("DataType")) = Bit
                Case adDate, adDBDate, adDBTime, adDBTimeStamp:                                                                          .TextMatrix(r, .ColIndex("DataType")) = DateTime
                Case adVarChar, adBSTR, adChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adVariant:                           .TextMatrix(r, .ColIndex("DataType")) = Str
            End Select
            .TextMatrix(r, .ColIndex("length")) = data(FieldName).DefinedSize
            On Error Resume Next
            .TextMatrix(r, .ColIndex("value")) = data(FieldName).Value
            On Error GoTo eh
            
            .RowOutlineLevel(r) = 1
            .IsSubtotal(r) = True
            
            fields.MoveNext
        Wend
        
        
        
        
        Call .AutoSize(0, .cols - 1)
        Call .Outline(0)
        
        

        
    End With
    
Exit Sub
eh: Call errHandler(srcFile & "LoadProperties")
End Sub

Public Sub EditQuote(quote As String)
    
    Dim f As FJob
    If quote = "" Then
        mDirty = False
       Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
    Else
        Set f = FindForm("FQuote", quote)
        If f Is Nothing Then
            mJob = quote
            Call LoadData
        End If
    End If
End Sub

Private Sub cboCommunity_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete, vbKeyBack
            cboCommunity.ListIndex = 0
    End Select
End Sub


Private Function LoadGLPrefixes()
    Dim s As String
    Dim rs As Recordset
    
    If HFApp.Options(AccountingSystem) = asTimberline And HFApp.Databases(dbAccounting).State = adStateOpen Then
        
        cboGLPrefix.Visible = True
        txtGLPrefix.Visible = False
        
        'which prefix table?
        s = "select account_prefix_a_length,account_prefix_ab_length,account_prefix_abc_length from glm_master__account_format"
        Set rs = HFApp.SqlExec(s, dbAccounting)
        Select Case True
            Case Val("" & rs(2)) <> 0: s = "select account_prefix_abc_description ,account_prefix_abc, 0   from glm_master__account_prefix_abc_1"
            Case Val("" & rs(1)) <> 0: s = "select account_prefix_ab_description ,account_prefix_ab, 0   from glm_master__account_prefix_ab_1"
            Case Val("" & rs(0)) <> 0: s = "select account_prefix_a_description ,account_prefix_a, 0   from glm_master__account_prefix_a_1"
        End Select
        Call LoadComboBox(cboGLPrefix, HFApp.Databases(dbAccounting), s)
    
    Else
        
        cboGLPrefix.Visible = False
        txtGLPrefix.Visible = True
    
    End If
    
End Function


Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal col As Long)
On Error Resume Next
    Dim r           As Long
    Dim rowBudgeted As Boolean
    Dim rowPOed     As Boolean
    Dim Rate        As Double
    Dim rs          As Recordset
    Dim s           As String
    Dim i           As Long
    
    
    With gItems
    
    For r = Min(Row, .RowSel) To Max(Row, .RowSel)
    
        If r > 0 Then
    
        rowBudgeted = .ValueMatrix(r, .ColIndex("BudgetGenerated")) <> 0
        rowPOed = Trim(.TextMatrix(r, .ColIndex("PONumber"))) <> ""
    
        Select Case .ColKey(col)
            'case "EstItemID":
            'case "EstAssemblyID":
            Case "Selected":
            'case "Job_No":
            'case "JobDesc":
            Case "JCExtra":
            Case "JCCostCode":
            Case "JCCostCodeDesc":
            'case "OriginalJCCategory":
            Case "JCCategory":
                If .TextMatrix(r, .ColIndex("BudgetTaxGroup")) = "" Then
                    s = ""
                    s = s & "SELECT * FROM TaxGroups WHERE TaxGroup=" & vbCrLf
                    s = s & "dbo.Purch_GetDefaultTaxGroup(" & DbQuote(Str, .TextMatrix(r, .ColIndex("Job_No"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                    s = s & "   ," & DbQuote(Str, cboPhase.Text) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("BudgetVendor"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("JCCategory"))) & ")"
                    Set rs = HFApp.SqlExec(s)
                    If Not rs.EOF Then
                        .TextMatrix(r, .ColIndex("BudgetTaxGroup")) = "" & rs("TaxGroup")
                        .TextMatrix(r, .ColIndex("BudgetJCTaxRate")) = "" & rs("JCRate")
                        .TextMatrix(r, .ColIndex("BudgetNJCTaxRate")) = "" & rs("NJCRate")
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    End If
                End If
                    
                If .TextMatrix(r, .ColIndex("POTaxGroup")) = "" Then
                    s = ""
                    s = s & "SELECT * FROM TaxGroups WHERE TaxGroup=" & vbCrLf
                    s = s & "dbo.Purch_GetDefaultTaxGroup(" & DbQuote(Str, .TextMatrix(r, .ColIndex("Job_No"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
                    s = s & "   ," & DbQuote(Str, cboPhase.Text) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("POVendor"))) & vbCrLf
                    s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("JCCategory"))) & ")"
                    Set rs = HFApp.SqlExec(s)
                    If Not rs.EOF Then
                        .TextMatrix(r, .ColIndex("POTaxGroup")) = "" & rs("TaxGroup")
                        .TextMatrix(r, .ColIndex("POJCTaxRate")) = "" & rs("JCRate")
                        .TextMatrix(r, .ColIndex("PONJCTaxRate")) = "" & rs("NJCRate")
                        .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    End If
                    
                End If
                
            Case "JCCategoryDesc":
            'case "HFDescription":
            'case "EstPhase":
            'case "EstItem":
            Case "ItemDesc":
            Case "ItemComments":
            
            Case "BudgetVendor", "BudgetVendorName":
                If GetVendorCost(r, .TextMatrix(r, .ColIndex("BudgetVendor")), Rate) Then
                    .TextMatrix(r, .ColIndex("BudgetOverridden")) = "False"
                    .TextMatrix(r, .ColIndex("BudgetRate")) = Rate
                    .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("POOverridden")) = "False"
                        .TextMatrix(r, .ColIndex("PORate")) = .TextMatrix(r, .ColIndex("BudgetRate"))
                        .TextMatrix(r, .ColIndex("POPretax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                        .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                        .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                    End If
                End If
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("POVendor")) = .TextMatrix(r, .ColIndex("BudgetVendor"))
                    .TextMatrix(r, .ColIndex("POVendorName")) = .TextMatrix(r, .ColIndex("BudgetVendorName"))
                End If
            
            Case "POVendor", "POVendorName":
                If GetVendorCost(r, .TextMatrix(r, .ColIndex("POVendor")), Rate) Then
                    .TextMatrix(r, .ColIndex("PORate")) = Rate
                    .TextMatrix(r, .ColIndex("POOverridden")) = "False"
                    .TextMatrix(r, .ColIndex("POPretax")) = Round(.ValueMatrix(r, .ColIndex("POQty")) * .ValueMatrix(r, .ColIndex("PORate")), 2)
                    .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    If Not rowBudgeted Then
                        .TextMatrix(r, .ColIndex("BudgetOverridden")) = "False"
                        .TextMatrix(r, .ColIndex("BudgetRate")) = .TextMatrix(r, .ColIndex("PORate"))
                        .TextMatrix(r, .ColIndex("BudgetPretax")) = .TextMatrix(r, .ColIndex("POPretax"))
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                    End If
                End If
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetVendor")) = .TextMatrix(r, .ColIndex("POVendor"))
                    .TextMatrix(r, .ColIndex("BudgetVendorName")) = .TextMatrix(r, .ColIndex("POVendorName"))
                End If
            
            Case "TakeoffQty":
                If mShowBudgets Then
                    '.TextMatrix(r, .ColIndex("BudgetQty")) = Roundto(.ValueMatrix(r, .ColIndex("TakeoffQty")) * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
                    .TextMatrix(r, .ColIndex("BudgetQty")) = RoundTo(.ValueMatrix(r, .ColIndex("TakeoffQty")) * (100 + .ValueMatrix(r, .ColIndex("WastePrecent"))) / 100 * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
                    .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    If Not rowPOed Then
                        .TextMatrix(r, .ColIndex("POQty")) = .TextMatrix(r, .ColIndex("BudgetQTY"))
                        .TextMatrix(r, .ColIndex("POPretax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                        .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                        .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                        .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                    End If
                Else
                    '.TextMatrix(r, .ColIndex("POQty")) = Roundto(.ValueMatrix(r, .ColIndex("TakeoffQty")) * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
                    .TextMatrix(r, .ColIndex("POQty")) = RoundTo(.ValueMatrix(r, .ColIndex("TakeoffQty")) * (100 + .ValueMatrix(r, .ColIndex("WastePercent"))) / 100 * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
                    .TextMatrix(r, .ColIndex("POPretax")) = Round(.ValueMatrix(r, .ColIndex("POQty")) * .ValueMatrix(r, .ColIndex("PORate")), 2)
                    .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    If Not rowBudgeted Then
                        .TextMatrix(r, .ColIndex("BudgetQty")) = .TextMatrix(r, .ColIndex("POQTY"))
                        .TextMatrix(r, .ColIndex("BudgetPretax")) = .TextMatrix(r, .ColIndex("POPretax"))
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                    End If
                End If
            
            Case "BudgetQty":
                .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("POQty")) = .TextMatrix(r, .ColIndex("BudgetQTY"))
                    .TextMatrix(r, .ColIndex("POPretax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                    .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                    .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                End If
                
            Case "POQty":
                .TextMatrix(r, .ColIndex("POPretax")) = Round(.ValueMatrix(r, .ColIndex("POQty")) * .ValueMatrix(r, .ColIndex("PORate")), 2)
                .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetQty")) = .TextMatrix(r, .ColIndex("POQTY"))
                    .TextMatrix(r, .ColIndex("BudgetPretax")) = .TextMatrix(r, .ColIndex("POPretax"))
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                End If
            
            Case "OrderUOM":
            
            Case "BudgetRate":
                .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                    .TextMatrix(r, .ColIndex("PORate")) = .TextMatrix(r, .ColIndex("BudgetRate"))
                    .TextMatrix(r, .ColIndex("POPretax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                    .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                    .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                End If
            
            Case "PORate":
                .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                .TextMatrix(r, .ColIndex("POPretax")) = Round(.ValueMatrix(r, .ColIndex("POQty")) * .ValueMatrix(r, .ColIndex("PORate")), 2)
                .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                    .TextMatrix(r, .ColIndex("BudgetRate")) = .TextMatrix(r, .ColIndex("PORate"))
                    .TextMatrix(r, .ColIndex("BudgetPretax")) = .TextMatrix(r, .ColIndex("POPretax"))
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                End If
            
            Case "BudgetPretax":
                If .ValueMatrix(r, .ColIndex("BudgetQty")) = 0 Then .TextMatrix(r, .ColIndex("BudgetQty")) = 1
                .TextMatrix(r, .ColIndex("BudgetRate")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) / .ValueMatrix(r, .ColIndex("BudgetQty")), 4)
                .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                    .TextMatrix(r, .ColIndex("POQty")) = .TextMatrix(r, .ColIndex("BudgetQty"))
                    .TextMatrix(r, .ColIndex("PORate")) = .TextMatrix(r, .ColIndex("BudgetRate"))
                    .TextMatrix(r, .ColIndex("POPreTax")) = .TextMatrix(r, .ColIndex("BudgetPretax"))
                    .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                    .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                End If
                
            Case "POPretax":
                If .ValueMatrix(r, .ColIndex("POQty")) = 0 Then .TextMatrix(r, .ColIndex("POQty")) = 1
                .TextMatrix(r, .ColIndex("POOverridden")) = "True"
                .TextMatrix(r, .ColIndex("PORate")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) / .ValueMatrix(r, .ColIndex("POQty")), 4)
                .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                    .TextMatrix(r, .ColIndex("BudgetQty")) = .TextMatrix(r, .ColIndex("POQty"))
                    .TextMatrix(r, .ColIndex("BudgetRate")) = .TextMatrix(r, .ColIndex("PORate"))
                    .TextMatrix(r, .ColIndex("BudgetPreTax")) = .TextMatrix(r, .ColIndex("POPretax"))
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                End If
            
            
            
            
            Case "BudgetTaxGroup":
                s = "SELECT JCRate,NJCRate FROM TaxGroups WHERE TaxGroup=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("BudgetTaxGroup")))
                Set rs = HFApp.SqlExec(s)
                .TextMatrix(r, .ColIndex("BudgetJCTaxRate")) = Val("" & rs("JCRate"))
                .TextMatrix(r, .ColIndex("BudgetNJCTaxRate")) = Val("" & rs("NJCRate"))
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                If Not rowPOed Then
                    .Cell(flexcpText, r, .ColIndex("POTaxGroup")) = .Cell(flexcpText, r, .ColIndex("BudgetTaxGroup"))
                    .Cell(flexcpText, r, .ColIndex("POJCTaxRate")) = .Cell(flexcpText, r, .ColIndex("BudgetJCTaxRate"))
                    .Cell(flexcpText, r, .ColIndex("PONJCTaxRate")) = .Cell(flexcpText, r, .ColIndex("BudgetNJCTaxRate"))
                    .Cell(flexcpText, r, .ColIndex("POJCTax")) = .Cell(flexcpText, r, .ColIndex("BudgetJCTax"))
                    .Cell(flexcpText, r, .ColIndex("PONJCTax")) = .Cell(flexcpText, r, .ColIndex("BudgetNJCTax"))
                    .Cell(flexcpText, r, .ColIndex("POTax")) = .Cell(flexcpText, r, .ColIndex("BudgetTax"))
                End If
                
            Case "POTaxGroup":
                s = "SELECT JCRate,NJCRate FROM TaxGroups WHERE TaxGroup=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("POTaxGroup")))
                Set rs = HFApp.SqlExec(s)
                .TextMatrix(r, .ColIndex("POJCTaxRate")) = Val("" & rs("JCRate"))
                .TextMatrix(r, .ColIndex("PONJCTaxRate")) = Val("" & rs("NJCRate"))
                .TextMatrix(r, .ColIndex("POJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("POJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("PONJCTax")) = Round(.ValueMatrix(r, .ColIndex("POPretax")) * .ValueMatrix(r, .ColIndex("PONJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                If Not rowBudgeted Then
                    .Cell(flexcpText, r, .ColIndex("BudgetTaxGroup")) = .Cell(flexcpText, r, .ColIndex("POTaxGroup"))
                    .Cell(flexcpText, r, .ColIndex("BudgetJCTaxRate")) = .Cell(flexcpText, r, .ColIndex("POJCTaxRate"))
                    .Cell(flexcpText, r, .ColIndex("BudgetNJCTaxRate")) = .Cell(flexcpText, r, .ColIndex("PONJCTaxRate"))
                    .Cell(flexcpText, r, .ColIndex("BudgetJCTax")) = .Cell(flexcpText, r, .ColIndex("POJCTax"))
                    .Cell(flexcpText, r, .ColIndex("BudgetNJCTax")) = .Cell(flexcpText, r, .ColIndex("PONJCTax"))
                    .Cell(flexcpText, r, .ColIndex("BudgetTax")) = .Cell(flexcpText, r, .ColIndex("POTax"))
                End If
            
            
            Case "BudgetJCTax":
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("POJCTax")) = .TextMatrix(r, .ColIndex("BudgetJCTax"))
                    .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                End If
                
            'case "BudgetJCTaxRate":
            
            Case "POJCTax":
                .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = .TextMatrix(r, .ColIndex("POJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                End If
            
            'case "POJCTaxRate":
            Case "BudgetNJCTax":
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                If Not rowPOed Then
                    .TextMatrix(r, .ColIndex("PONJCTax")) = .TextMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("POTax")) = .TextMatrix(r, .ColIndex("BudgetTax"))
                End If
            
            'case "BudgetNJCTaxRate":
            Case "PONJCTax":
                .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                If Not rowBudgeted Then
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = .TextMatrix(r, .ColIndex("PONJCTax"))
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .TextMatrix(r, .ColIndex("POTax"))
                End If
            
            'case "PONJCTaxRate":
            'case "BudgetTax":
            'case "POTax":
            Case "POIndex":
            'case "SalesWorksheet":
            'case "PONumber":
            'case "BudgetGenerated":
            Case "ExcludeFromPO":
        End Select
        
        
        If .ColKey(col) <> "Selected" Then
            Dirty = True
            For i = r To .RowSel
                If .RowData(i) <> "NEW" Then .RowData(i) = "DIRTY"
            Next
        End If
        
        End If
    Next
    End With
    Call ColorizeItems(-1)
End Sub








Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
Static inhere As Boolean
If inhere Then Exit Sub
inhere = True
    gItems.ColSel = gItems.col
inhere = False
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    Dim rowBudgeted As Boolean
    Dim rowPOed As Boolean
    Dim rowEither As Boolean
    
    With gItems
        .ComboList = ""
        .EditMaxLength = 0
        .AutoSearch = flexSearchNone
        
        rowBudgeted = .ValueMatrix(Row, .ColIndex("BudgetGenerated")) <> 0
        rowPOed = Trim(.TextMatrix(Row, .ColIndex("PONumber"))) <> ""
        rowEither = rowBudgeted Or rowPOed
        
        If ((rowBudgeted And mShowBudgets) Or (rowPOed And Not mShowBudgets)) And .ColKey(col) <> "ExcludeFromPO" Then
            Cancel = True
            Exit Sub
        End If
        
        Select Case .ColKey(col)
            'Case "EstItemID":
            'Case "EstAssemblyID":
            Case "Selected":           Cancel = .TextMatrix(Row, .ColIndex("EstItemID")) = ""
            Case "BudgetApproved":     Cancel = .TextMatrix(Row, .ColIndex("EstItemID")) = ""
            'Case "Job_No":
            'Case "JobDesc":
            Case "JCExtra":            Cancel = rowEither:         .ComboList = IIf(HFApp.Options(Use_Timberline), "|...", ""):   .EditMaxLength = 10
            Case "JCCostCode":         Cancel = rowEither:         .ComboList = "|..."
            Case "JCCostCodeDesc":     Cancel = rowEither:         .ComboList = "..."
            'Case "OriginalJCCategory"
            Case "JCCategory":         Cancel = rowEither:         .ComboList = "|..."
            Case "JCCategoryDesc":     Cancel = rowEither:         .ComboList = "..."
            'Case "HFDescription":
            'Case "EstPhase":
            'Case "EstItem":
            Case "ItemDesc":
                .ComboList = "|...":        .EditMaxLength = 200
            Case "ItemComments":                                   .ComboList = "|...":
            Case "BudgetVendor":       Cancel = rowBudgeted:       .ComboList = "|..."
            Case "POVendor":           Cancel = rowPOed:           .ComboList = "|..."
            Case "BudgetVendorName":   Cancel = rowBudgeted:       .ComboList = "..."
            Case "POVendorName":       Cancel = rowPOed:           .ComboList = "..."
            
            Case "TakeoffQty":         Cancel = (rowBudgeted And mShowBudgets) Or (rowPOed And Not mShowBudgets)
            Case "BudgetQty":          Cancel = rowBudgeted
            Case "POQty":              Cancel = rowPOed
            Case "BudgetRate":         Cancel = rowBudgeted
            Case "PORate":             Cancel = rowPOed
            Case "OrderUOM":                                       .ComboList = "|...":        .EditMaxLength = 10
            Case "BudgetPretax":       Cancel = rowBudgeted
            Case "POPretax":           Cancel = rowPOed
            Case "BudgetTaxGroup":     Cancel = rowBudgeted:       .ComboList = "|..."
            Case "POTaxGroup":         Cancel = rowPOed:           .ComboList = "|..."
            Case "BudgetJCTax":        Cancel = rowBudgeted
            'Case "BudgetJCTaxRate":
            Case "POJCTax":            Cancel = rowPOed
            'Case "POJCTaxRate":
            Case "BudgetNJCTax":       Cancel = rowBudgeted
            'Case "BudgetNJCTaxRate":
            Case "PONJCTax":           Cancel = rowPOed
            'Case "PONJCTaxRate":
            'Case "BudgetTax":
            'Case "POTax":
            Case "POIndex":            Cancel = rowEither:         .ComboList = "|..."
            Case "POIndexDescription": Cancel = rowEither:         .ComboList = "..."
            'Case "PONumber":
            'Case "BudgetGenerated":
            Case "ExcludeFromPO":
            'case "SalesWorksheet":
            Case Else:                 Cancel = True
                
        End Select
        
        .AutoSearch = IIf(Cancel, flexSearchFromCursor, flexSearchNone)
        
    End With
End Sub



Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    Dim r As Long
    If Button = vbRightButton Then
        If gItems.MouseRow = 0 Then
            Cancel = True
            Call FMain.ShowColumnMenu(gItems)
        Else
            If gItems.Row < 1 And gItems.MouseRow > 1 Then gItems.Row = gItems.MouseRow
            If gItems.Row > 0 Then
            Set MouseCtrl = gItems
            MouseCol = gItems.MouseCol
            FMain.mnuEstimateItemsGridSub(mcITEM_CANCELBUDGETS).Enabled = mShowBudgets
            
            
            
            With Me.gItems
                     
                'can do this only if item is onetime
                FMain.mnuEstimateItemsGridSub(mcITEM_SAVEONETIMETODB).Enabled = .TextMatrix(.Row, .ColIndex("EstPhase")) = ""
            
            
                'can do this if budgeted not generated and po not generated
                FMain.mnuEstimateItemsGridSub(mcITEM_CHANGEITEM).Enabled = Not (.ValueMatrix(.Row, .ColIndex("BudgetGenerated")) <> 0 Or Trim(.TextMatrix(.Row, .ColIndex("PONumber"))) <> "")
            End With
            
            
            FMain.mnuEstimateItemsGridSub(mcITEM_CANCELBUDGETS).Enabled = False
            
            PopupMenu FMain.mnuEstimateItemsGrid
            End If
        End If
    End If
End Sub






Private Sub gItems_BeforeUserResize(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    Select Case gItems.ColKey(col)
        Case "WarningMessages", "Selected"
            Cancel = True
    End Select
End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal col As Long)
    Dim s As String
    Dim i As Long
    With gItems
        Select Case .ColKey(col)
        
            Case "JCExtra"
                s = "select Extra,Description from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No"))))
                If FPickList.Choose(HFApp.Databases(dbAccounting), "Extra", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCExtra"), .RowSel, .ColIndex("JCExtra")) = FPickList.SelectedItem("Extra")
                End If
                
            Case "JCCostCode"
                s = "SELECT CostCode,Description FROM StandardCostCodes"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCostCode"), .RowSel, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, Row, .ColIndex("JCCostCodeDesc"), .RowSel, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "JCCostCodeDesc"
                s = "SELECT Description,CostCode FROM StandardCostCodes"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCostCode"), .RowSel, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, Row, .ColIndex("JCCostCodeDesc"), .RowSel, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "JCCategory"
                s = "SELECT Category,Description FROM StandardCategories"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCategory"), .RowSel, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, Row, .ColIndex("JCCategoryDesc"), .RowSel, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "JCCategoryDesc"
                s = "SELECT Description,Category FROM StandardCategories"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCategory"), .RowSel, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, Row, .ColIndex("JCCategoryDesc"), .RowSel, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "OrderUOM"
                s = "SELECT DISTINCT OrderUOM Unit FROM tblPhaseItem"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Units", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("OrderUOM"), .RowSel, .ColIndex("OrderUOM")) = FPickList.SelectedItem("OrderUOM")
                End If
            
            Case "ItemComments"
                s = gItems.Text
                If FComments.Edit(s, gItems) Then
                    gItems.Text = s
                End If
                
            Case "ItemDesc":
                s = gItems.Text
                If FComments.Edit(s, gItems, , "Description", 200) Then
                    gItems.Text = s
                End If
                
            Case "BudgetVendor"
                s = "Approved Vendors" & Chr(1) & "SELECT DISTINCT v.Vendor_ID Vendor, v.Vendor_Name Company,v.City,v.Phone FROM tblVendors v JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor) WHERE (ISNULL(Community,'')='' OR " & " Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & ") AND (ISNULL(Assembly,'')='' OR " & " Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ") AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & Chr(0) & _
                    "All Vendors" & Chr(1) & "SELECT Vendor_ID Vendor, Vendor_Name Company,City,Phone FROM tblVendors"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , FMain.SmallIcons.ListImages("vendor").Picture) Then
                    .Cell(flexcpText, Row, .ColIndex("BudgetVendor"), .RowSel, .ColIndex("BudgetVendor")) = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Row, .ColIndex("BudgetVendorName"), .RowSel, .ColIndex("BudgetVendorName")) = FPickList.SelectedItem("Company")
                End If
                
            Case "BudgetVendorName"
                s = "Approved Vendors" & Chr(1) & "SELECT DISTINCT v.Vendor_Name Company, v.Vendor_ID Vendor,v.City,v.Phone FROM tblVendors v JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor) WHERE (ISNULL(Community,'')='' OR " & " Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & ") AND (ISNULL(Assembly,'')='' OR " & " Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ") AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & Chr(0) & _
                    "All Vendors" & Chr(1) & "SELECT Vendor_Name Company,Vendor_ID Vendor,City,Phone FROM tblVendors"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , FMain.SmallIcons.ListImages("vendor").Picture) Then
                    .Cell(flexcpText, Row, .ColIndex("BudgetVendor"), .RowSel, .ColIndex("BudgetVendor")) = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Row, .ColIndex("BudgetVendorName"), .RowSel, .ColIndex("BudgetVendorName")) = FPickList.SelectedItem("Company")
                End If
            
            Case "POVendor"
                s = "Approved Vendors" & Chr(1) & "SELECT DISTINCT v.Vendor_ID Vendor, v.Vendor_Name Company,v.City,v.Phone FROM tblVendors v JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor) WHERE (ISNULL(Community,'')='' OR " & " Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & ") AND (ISNULL(Assembly,'')='' OR " & " Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ") AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & Chr(0) & _
                    "All Vendors" & Chr(1) & "SELECT Vendor_ID Vendor, Vendor_Name Company,City,Phone FROM tblVendors"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , FMain.SmallIcons.ListImages("vendor").Picture) Then
                    .Cell(flexcpText, Row, .ColIndex("POVendor"), .RowSel, .ColIndex("POVendor")) = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Row, .ColIndex("POVendorName"), .RowSel, .ColIndex("POVendorName")) = FPickList.SelectedItem("Company")
                End If
            
            Case "POVendorName"
                s = "Approved Vendors" & Chr(1) & "SELECT DISTINCT v.Vendor_Name Company, v.Vendor_ID Vendor,v.City,v.Phone FROM tblVendors v JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor) WHERE (ISNULL(Community,'')='' OR " & " Community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & ") AND (ISNULL(Assembly,'')='' OR " & " Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ") AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & Chr(0) & _
                    "All Vendors" & Chr(1) & "SELECT Vendor_Name Company,Vendor_ID Vendor,City,Phone FROM tblVendors"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , FMain.SmallIcons.ListImages("vendor").Picture) Then
                    .Cell(flexcpText, Row, .ColIndex("POVendor"), .RowSel, .ColIndex("POVendor")) = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Row, .ColIndex("POVendorName"), .RowSel, .ColIndex("POVendorName")) = FPickList.SelectedItem("Company")
                End If

            Case "BudgetTaxGroup"
                s = "SELECT TaxGroup,Description FROM TaxGroups"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gItems) Then
                    .Cell(flexcpText, Row, col, .RowSel, col) = FPickList.SelectedItem("TaxGroup")
                End If
        
            Case "POTaxGroup"
                s = "SELECT TaxGroup,Description FROM TaxGroups"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gItems) Then
                    .Cell(flexcpText, Row, col, .RowSel, col) = FPickList.SelectedItem("TaxGroup")
                End If
            
            Case "POIndex"
'                s = "SELECT POIndex,Description FROM tblPOIndex"
                s = "SELECT POIndex,case when poindex=description then '' else Description end Description FROM tblPOIndex"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "PO Index", s, gItems) Then
                    .Cell(flexcpText, Row, col, .RowSel, col) = FPickList.SelectedItem("POIndex")
                    .Cell(flexcpText, Row, .ColIndex("POIndexDescription"), .RowSel, .ColIndex("POIndexDescription")) = FPickList.SelectedItem("Description")
                End If

            Case "POIndexDescription"
                s = "SELECT Description,POIndex FROM tblPOIndex"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "PO Index", s, gItems) Then
                    .Cell(flexcpText, Row, col, .RowSel, col) = FPickList.SelectedItem("Description")
                    .Cell(flexcpText, Row, .ColIndex("POIndex"), .RowSel, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                End If

        End Select
        
        Call gItems_AfterEdit(Row, col)
        
    End With
End Sub






Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim r As Long
    Dim CanDeleteBudget As Boolean
    Dim CanDeletePO     As Boolean
    
    With gItems
        Select Case True
            Case KeyCode = vbKeyF And Shift = vbCtrlMask
                Call FFind.Showform(gItems)
                
            Case KeyCode = vbKeyDelete And Shift = vbCtrlMask
                For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    CanDeleteBudget = .TextMatrix(r, .ColIndex("BudgetGenerated")) <> "True"
                    CanDeletePO = .TextMatrix(r, .ColIndex("PONumber")) = ""
                    If (CanDeleteBudget And mShowBudgets) Or (CanDeletePO And Not mShowBudgets) Then
                        Dirty = True
                        .RowHidden(r) = True
                        .RowData(r) = "DELETE"
                    End If
                    .Row = -1
                Next
                
        End Select
    End With
    
End Sub

Private Sub gItems_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    If Button <> 0 Then Exit Sub
    With gItems
        If mFunnyFlag Then
            mFunnyFlag = False
            Exit Sub
        End If
        If .MouseRow >= Min(.Row, .RowSel) And .MouseRow <= Max(.Row, .RowSel) And .MouseCol = .col Then
            Call gItems_SelChange
        Else
            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315)
        End If
    End With
End Sub

Private Sub gItems_RowColChange()
On Error Resume Next
    With gItems
        mFunnyFlag = True
        Call ShowTip(.Row, .col, .colPos(.col) + 180, .RowPos(.Row) + .RowHeight(.Row) + 180)
    End With
End Sub

Private Sub gItems_SelChange()
    Dim r As Long
    Dim prefix As String
    
    Dim Count As Long
    Dim pretax As Double
    Dim tax As Double
    Dim tip As String
    
    If mShowBudgets Then
        prefix = "Budget"
    Else
        prefix = "PO"
    End If
    
    With gItems
        If .Row <> .RowSel Then
            For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                Count = Count + 1
                pretax = pretax + .ValueMatrix(r, .ColIndex(prefix & "Pretax"))
                tax = tax + .ValueMatrix(r, .ColIndex(prefix & "Tax"))
            Next
            
            tip = "" & _
                  "Pretax:" & vbTab & format(pretax, "#,##0.00") & vbCrLf & _
                  "Tax:" & vbTab & format(tax, "#,##0.00") & vbCrLf & _
                  "Total:" & vbTab & format(pretax + tax, "#,##0.00")

            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315, tip)
            
        End If
    End With
End Sub


Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    Dim s As String
    Dim i As Long
    Dim Description As String
    Dim rs As Recordset
    
    With gItems
        s = .EditText
        Select Case .ColKey(col)
            'case "EstItemID":
            'case "EstAssemblyID":
            'case "Job_No":
            'case "JobDesc":
            
            Case "JCExtra"
                If HFApp.Options(Use_Timberline) And Trim(s) <> "" Then
                    Set rs = HFApp.SqlExec("select extra from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No")))) & " AND extra=" & DbQuote(Str, s), dbAccounting)
                    If rs.EOF Then
                        Description = InputBox(vbCrLf & "Extra '" & s & "' could not be found." & vbCrLf & "Do you want to add it?" & vbCrLf & vbCrLf & vbCrLf & "Enter a description for the extra. (30 characters or less)", App.ProductName)
                        If Description = "" Then
                            Cancel = True
                        Else
                            s = ""
                            s = s & "INSERT INTO jcm_master__extra(Job,Extra,Description,Status)" & vbCrLf
                            s = s & "VALUES(" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No")))) & vbCrLf
                            s = s & "      ," & DbQuote(Str, .EditText) & vbCrLf
                            s = s & "      ," & DbQuote(Str, left(Description, 30)) & vbCrLf
                            s = s & "      ,'In progress')" & vbCrLf
                            Set rs = HFApp.SqlExec(s, dbAccounting)
                            s = .EditText
                        End If
                    Else
                        s = "" & rs(0)
                    End If
                End If
            
                                   
            Case "JCCostCode":              Cancel = Not ValidateField(gItems, s, "Cost Code not found", "SELECT CostCode,Description FROM StandardCostCodes WHERE CostCode=" & DbQuote(Str, s), "JCCostCodeDesc")
            'Case "JCCostCodeDesc":
            'case "OriginalJCCategory":
            Case "JCCategory":              Cancel = Not ValidateField(gItems, s, "Category not found", "SELECT Category,Description FROM StandardCategories WHERE Category=" & DbQuote(Str, s), "JCCategoryDesc")
            'Case "JCCategoryDesc":
            'case "HFDescription":
            'case "EstPhase":
            'case "EstItem":
            Case "ItemDesc":
            Case "ItemComments":
            Case "BudgetVendor":            Cancel = Not ValidateField(gItems, s, "Vendor not found", "SELECT Vendor_ID,Vendor_Name FROM tblVendors WHERE Vendor_ID=" & DbQuote(Str, s), "BudgetVendorName")
            Case "POVendor":                Cancel = Not ValidateField(gItems, s, "Vendor not found", "SELECT Vendor_ID,Vendor_Name FROM tblVendors WHERE Vendor_ID=" & DbQuote(Str, s), "POVendorName")
            'Case "BudgetVendorName":
            'Case "POVendorName":
            'case "TakeoffQty":
            Case "BudgetQty":               Cancel = Not IsNumeric(s)
            Case "POQty":                   Cancel = Not IsNumeric(s)
            Case "OrderUOM":
            Case "BudgetRate":              Cancel = Not IsNumeric(s)
            Case "PORate":                  Cancel = Not IsNumeric(s)
            Case "BudgetPretax":            Cancel = Not IsNumeric(s)
            Case "POPretax":                Cancel = Not IsNumeric(s)
            
            Case "BudgetTaxGroup":          Cancel = Not ValidateField(gItems, s, "Tax Group not found", "SELECT TaxGroup FROM TaxGroups WHERE TaxGroup=" & DbQuote(Str, s))
                
            Case "POTaxGroup":              Cancel = Not ValidateField(gItems, s, "Tax Group not found", "SELECT TaxGroup FROM TaxGroups WHERE TaxGroup=" & DbQuote(Str, s))
            Case "BudgetJCTax":             Cancel = Not IsNumeric(s)
            'case "BudgetJCTaxRate":
            Case "POJCTax":                 Cancel = Not IsNumeric(s)
            'case "POJCTaxRate":
            Case "BudgetNJCTax":            Cancel = Not IsNumeric(s)
            'case "BudgetNJCTaxRate":
            Case "PONJCTax":                Cancel = Not IsNumeric(s)
            'case "PONJCTaxRate":
            'case "BudgetTax":
            'case "POTax":
            Case "POIndex":                 Cancel = Not ValidateField(gItems, s, "PO Index not found", "SELECT POIndex,Description POIndexDescription FROM tblPOIndex WHERE POIndex=" & DbQuote(Str, s), "POIndexDescription")
            'case "PONumber":
            'case "BudgetGenerated":
            Case "ExcludeFromPO":
            'case "SalesWorksheet":
        End Select
        .EditText = s
        If .ColKey(col) <> "Selected" Then
            Dirty = True
            For i = Min(Row, .RowSel) To Max(Row, .RowSel)
                If .RowData(i) <> "NEW" Then .RowData(i) = "DIRTY"
            Next
        End If
    End With
End Sub







Private Sub ShowTip(Row As Long, col As Long, x As Long, Y As Long, Optional tip As String)
    With gItems
        If Row < 0 Or col < 0 Or (.RowSel = .Row And tip <> "") Then
            picWarningMessages.Visible = False
        Else
            lblWarningMessages = IIf(tip <> "", tip, gItems.Cell(flexcpData, Row, col))
            
            imgTipIcon.Picture = IIf(tip <> "", imgInfo.Picture, imgWarning.Picture)
            
            
            picWarningMessages.Visible = lblWarningMessages <> ""
            
            Set Me.Font = lblWarningMessages.Font
            
            lblWarningMessages.Alignment = IIf(tip <> "", 1, 0)
            lblWarningMessages.Width = Me.TextWidth(lblWarningMessages) + lblWarningMessages.left
            lblWarningMessages.Height = Me.TextHeight(lblWarningMessages) + lblWarningMessages.Top
            
            picWarningMessages.Width = lblWarningMessages.Width + lblWarningMessages.left + lblWarningMessages.Top
            picWarningMessages.Height = lblWarningMessages.Height + 2 * lblWarningMessages.Top
            If .Height - Y - picWarningMessages.Height - 365 < 0 Then Y = .Height - picWarningMessages.Height - 365
            If .Width - x - picWarningMessages.Width - 365 < 0 Then x = .Width - picWarningMessages.Width - 365
            picWarningMessages.Top = Y
            picWarningMessages.left = x
            
            shpWarningMessages.Width = picWarningMessages.Width
            shpWarningMessages.Height = picWarningMessages.Height
        End If
    End With
End Sub

Public Sub AddItem(Assembly As String, _
                   AssemblyDescription As String, _
                   Model As String, _
                   Phase As String, _
                   item As String, _
                   Description As String, _
                   OrderQty As Double, _
                   OrderUOM As String, _
                   TakeoffQty As Double, _
                   TakeoffUOM As String, _
                   ConversionFactor As Double, RoundTo As Double, RoundDir As Long, WastePercent As Long, _
                   JCExtra As String, _
                   JCCostCode As String, _
                   JCCostCodeDesc As String, _
                   JCCategory As String, _
                   JCCategoryDesc As String, _
                   vendor As String, _
                   vendorName As String, _
                   Price As Double, _
                   TaxGroup As String, _
                   TaxGroupName As String, _
                   JCTaxRate As Double, _
                   NJCTaxRate As Double, _
                   POIndex As String, _
                   Comments As String)
On Error GoTo eh
    
    Dim i As Long
    Dim s As String
    
    Dim rs As Recordset
    Dim RndTo As Double
    Dim RndDir As Long
    
    
    Dirty = True
    With gItems
        .AddItem ""
        i = .Rows - 1
        
        On Error Resume Next
        Set rs = HFApp.SqlExec("select roundto,rounddir from tblphaseitem where phase=" & DbQuote(Str, Phase) & " and item=" & DbQuote(Str, item), dbHomefront)
        RndTo = Val("" & rs("RoundTo"))
        RndDir = Val("" & rs("RoundDir"))
        On Error GoTo eh
        
        .RowData(i) = "NEW"
        .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexNoCheckbox
        
        .TextMatrix(i, .ColIndex("Assembly")) = Assembly
        .TextMatrix(i, .ColIndex("AssemblyDescription")) = AssemblyDescription
        .TextMatrix(i, .ColIndex("EstAssemblyID")) = mEstAssemblyID
        .TextMatrix(i, .ColIndex("Job_No")) = mJob
        .TextMatrix(i, .ColIndex("JCExtra")) = JCExtra
        .TextMatrix(i, .ColIndex("EstPhase")) = Phase
        .TextMatrix(i, .ColIndex("EstItem")) = item
        .TextMatrix(i, .ColIndex("ItemDesc")) = Description
        .TextMatrix(i, .ColIndex("ConversionFactor")) = ConversionFactor
        .TextMatrix(i, .ColIndex("RoundTo")) = RoundTo
        .TextMatrix(i, .ColIndex("RoundDir")) = RoundDir
        .TextMatrix(i, .ColIndex("WastePercent")) = WastePercent
        .TextMatrix(i, .ColIndex("JCCostCode")) = JCCostCode
        .TextMatrix(i, .ColIndex("JCCostCodeDesc")) = JCCostCodeDesc
        .TextMatrix(i, .ColIndex("JCCategory")) = JCCategory
        .TextMatrix(i, .ColIndex("OriginalJCCategory")) = JCCategory
        .TextMatrix(i, .ColIndex("JCCategoryDesc")) = JCCategoryDesc
        .TextMatrix(i, .ColIndex("POIndex")) = POIndex
        .TextMatrix(i, .ColIndex("ItemComments")) = Comments
        .TextMatrix(i, .ColIndex("TakeoffQty")) = TakeoffQty
        .TextMatrix(i, .ColIndex("TakeoffUOM")) = TakeoffUOM
        .TextMatrix(i, .ColIndex("OrderUOM")) = OrderUOM
        .TextMatrix(i, .ColIndex("Model")) = Model
        .TextMatrix(i, .ColIndex("OptionID")) = ""
        
        .TextMatrix(i, .ColIndex("RoundTo")) = RndTo
        .TextMatrix(i, .ColIndex("RoundDir")) = RndDir
        .TextMatrix(i, .ColIndex("BudgetQty")) = OrderQty
        .TextMatrix(i, .ColIndex("BudgetVendor")) = vendor
        .TextMatrix(i, .ColIndex("BudgetVendorName")) = vendorName
        .TextMatrix(i, .ColIndex("BudgetRate")) = Price
        .TextMatrix(i, .ColIndex("BudgetPretax")) = Round(OrderQty * Price, 2)
        .TextMatrix(i, .ColIndex("BudgetTaxGroup")) = TaxGroup
        .TextMatrix(i, .ColIndex("BudgetJCTaxRate")) = JCTaxRate
        .TextMatrix(i, .ColIndex("BudgetNJCTaxRate")) = NJCTaxRate
        .TextMatrix(i, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(i, .ColIndex("BudgetPretax")) * JCTaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(i, .ColIndex("BudgetPretax")) * NJCTaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("BudgetTax")) = .ValueMatrix(i, .ColIndex("BudgetNJCTax")) + .ValueMatrix(i, .ColIndex("BudgetJCTax"))
        .TextMatrix(i, .ColIndex("POQty")) = .TextMatrix(i, .ColIndex("BudgetQty"))
        .TextMatrix(i, .ColIndex("POVendor")) = .TextMatrix(i, .ColIndex("BudgetVendor"))
        .TextMatrix(i, .ColIndex("POVendorName")) = .TextMatrix(i, .ColIndex("BudgetVendorName"))
        .TextMatrix(i, .ColIndex("PORate")) = .TextMatrix(i, .ColIndex("BudgetRate"))
        .TextMatrix(i, .ColIndex("POPretax")) = .TextMatrix(i, .ColIndex("BudgetPretax"))
        .TextMatrix(i, .ColIndex("POTaxGroup")) = .TextMatrix(i, .ColIndex("BudgetTaxGroup"))
        .TextMatrix(i, .ColIndex("POJCTaxRate")) = .TextMatrix(i, .ColIndex("BudgetJCTaxRate"))
        .TextMatrix(i, .ColIndex("PONJCTaxRate")) = .TextMatrix(i, .ColIndex("BudgetNJCTaxRate"))
        .TextMatrix(i, .ColIndex("POJCTax")) = .TextMatrix(i, .ColIndex("BudgetJCTax"))
        .TextMatrix(i, .ColIndex("PONJCTax")) = .TextMatrix(i, .ColIndex("BudgetNJCTax"))
        .TextMatrix(i, .ColIndex("POTax")) = .TextMatrix(i, .ColIndex("BudgetTax"))

    End With
Exit Sub
eh: Call errHandler(srcFile & "AddItem")
End Sub



Private Sub ClearWhereClause()
    Dim r As Long
    
    If gItems.Rows > 1 Then
        gItems.Cell(flexcpChecked, 1, gItems.ColIndex("Selected"), gItems.Rows - 1, gItems.ColIndex("Selected")) = flexUnchecked
    End If
End Sub

Public Sub ColorizeItems(r As Long)
    Const RateZeroColor = 16646111 'light cyan &HC0C0FF 'pink
    Dim s           As String
    Dim rowBudgeted As Boolean
    Dim rowPOed     As Boolean
    Dim rowLocked   As Boolean
    Dim startRow    As Long
    Dim EndRow      As Long
        
    With gItems
        If r > 0 Then
            startRow = r
            EndRow = r
        Else
            startRow = 1
            EndRow = .Rows - 1
        End If
        For r = startRow To EndRow
        
            s = ""
            rowBudgeted = .ValueMatrix(r, .ColIndex("BudgetGenerated")) <> 0
            rowPOed = Trim(.TextMatrix(r, .ColIndex("PONumber"))) <> ""
            rowLocked = (rowBudgeted And mShowBudgets) Or (rowPOed And Not mShowBudgets)
            
            'clear all
            .Cell(flexcpForeColor, r, 0, r, .cols - 1) = vbWindowText
            .Cell(flexcpBackColor, r, 0, r, .cols - 1) = vbWindowBackground
            .Cell(flexcpFontItalic, r, 0, r, .cols - 1) = False
            .Cell(flexcpFontBold, r, 0, r, .cols - 1) = False
            
            
            
            .Cell(flexcpForeColor, r, 0, r, .cols - 1) = vbWindowText
            
            If rowLocked Then
                .Cell(flexcpForeColor, r, 0, r, .cols - 1) = vbGrayText
            Else
                
                ' less than zero
                If .ValueMatrix(r, .ColIndex(IIf(mShowBudgets, "BudgetRate", "PORate"))) < 0 Or .ValueMatrix(r, .ColIndex(IIf(mShowBudgets, "BudgetQty", "POQty"))) < 0 Then
                    .Cell(flexcpForeColor, r, 0, r, .cols - 1) = HFApp.Options(Format_QtyRateLTZero_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .cols - 1) = HFApp.Options(Format_QtyRateLTZero_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_QtyRateLTZero_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_QtyRateLTZero_FontStyle), "Bold")
                End If
                
                ' equal to zero
                If .ValueMatrix(r, .ColIndex(IIf(mShowBudgets, "BudgetPretax", "POPretax"))) = 0 Then
                    .Cell(flexcpForeColor, r, 0, r, .cols - 1) = HFApp.Options(Format_QtyRateEQZero_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .cols - 1) = HFApp.Options(Format_QtyRateEQZero_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_QtyRateEQZero_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_QtyRateEQZero_FontStyle), "Bold")
                End If
                
                
                ' if overridden value
                If .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True" Then
                    .Cell(flexcpForeColor, r, 0, r, .cols - 1) = HFApp.Options(Format_NonSysRate_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .cols - 1) = HFApp.Options(Format_NonSysRate_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Bold")
                End If
                
                If .TextMatrix(r, .ColIndex("POOverridden")) = "True" Then
                    .Cell(flexcpForeColor, r, 0, r, .cols - 1) = HFApp.Options(Format_NonSysRate_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .cols - 1) = HFApp.Options(Format_NonSysRate_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Bold")
                End If
                
                
                ' figure out if there are errors
                If mShowBudgets Then
                    If Trim(.TextMatrix(r, .ColIndex("POIndex"))) = "" Then s = s & vbCrLf & "PO Index not specified"
                    If Trim(.TextMatrix(r, .ColIndex("JCCostCode"))) = "" Then s = s & vbCrLf & "Cost Code not specified"
                    If Trim(.TextMatrix(r, .ColIndex("JCCategory"))) = "" Then s = s & vbCrLf & "Category not specified"
                    If Trim(.TextMatrix(r, .ColIndex("BudgetVendor"))) = "" Then s = s & vbCrLf & "Vendor not specified"
                    If HFApp.Options(TaxGroupRequired) Then If Trim(.TextMatrix(r, .ColIndex("BudgetTaxGroup"))) = "" Then s = s & vbCrLf & "Tax Group not specified"
                Else
                    If Trim(.TextMatrix(r, .ColIndex("POIndex"))) = "" Then s = s & vbCrLf & "PO Index not specified"
                    If Trim(.TextMatrix(r, .ColIndex("JCCostCode"))) = "" Then s = s & vbCrLf & "Cost code not specified"
                    If Trim(.TextMatrix(r, .ColIndex("JCCategory"))) = "" Then s = s & vbCrLf & "Category not specified"
                    If Trim(.TextMatrix(r, .ColIndex("POVendor"))) = "" Then s = s & vbCrLf & "Vendor not specified"
                    If HFApp.Options(TaxGroupRequired) Then If Trim(.TextMatrix(r, .ColIndex("POTaxGroup"))) = "" Then s = s & vbCrLf & "Tax Group not specified"
                End If
                gItems.Cell(flexcpData, r, 0, r, .cols - 1) = Mid(s, 3)
                If s = "" Then
                    .Cell(flexcpPicture, r, .ColIndex("WarningMessages")) = Nothing
                Else
                    .Cell(flexcpPicture, r, .ColIndex("WarningMessages")) = imgWarning.Picture
                    .Cell(flexcpForeColor, r, 0, r, .cols - 1) = HFApp.Options(Format_InvalidData_ForeColor)
                    .Cell(flexcpBackColor, r, 0, r, .cols - 1) = HFApp.Options(Format_InvalidData_BackColor)
                    .Cell(flexcpFontItalic, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_InvalidData_FontStyle), "Italic")
                    .Cell(flexcpFontBold, r, 0, r, .cols - 1) = InStr(1, HFApp.Options(Format_InvalidData_FontStyle), "Bold")
                End If
                
            End If
        Next
    End With
End Sub


Private Function GetVendorCost(Row As Long, vendor As String, Rate As Double) As Boolean
    Dim s As String
    Dim r As Double
    With gItems
    
        s = ""
        s = s & "SELECT dbo.Purch_GetItemRate(" & vbCrLf
        s = s & "       0" & vbCrLf
        s = s & "      ,0" & vbCrLf
        s = s & "      ," & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
        s = s & "      ," & DbQuote(Str, cboPhase.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Model"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("OptionID"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & vbCrLf
        s = s & "      ," & DbQuote(Num, .TextMatrix(Row, .ColIndex("Seq"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, vendor) & vbCrLf
        s = s & "      ,GETDATE())"
        On Error Resume Next
        r = HFApp.SqlExec(s)(0)
        
        If r <> 0 Or HFApp.Options(ZeroRateOnChangeVendor) Then
            GetVendorCost = True
            Rate = r
        End If
        
    
    End With
End Function

Public Sub mnuEstimateItemsGridSub_Click(Index As Integer)
    Dim rc As Long
    Dim rc2 As Long
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim newrow As Long
    Dim Phase As String
    Dim item As String
    Dim Description As String
    
    With gItems
        If .Row <= 0 Then Exit Sub
        
        Select Case Index
        
            Case mcITEM_VIEWFILES
                'Call FAttachments.Showform(ObjectID, .TextMatrix(.Row, .ColIndex("Description")), "")
                'this doesn't work. the popup menus on FAttachments don't work if the screen is displayed from inside this procedure
                'to get around it we have to do something weird.
                ' 1.) turn on a timer then exit sub
                ' 2.) when timer goes disable it and launch the attachments window
                Timer1.Interval = 10
                Timer1.Enabled = True
                
                
            Case mcITEM_INSERTFILE, mcITEM_LINKTOFILE
                Call FAttachments.AddFile("EST~" & .TextMatrix(.Row, .ColIndex("EstAssemblyID")) & "~" & .TextMatrix(.Row, .ColIndex("EstPhase")) & "~" & .TextMatrix(.Row, .ColIndex("EstItem")), Index = mcITEM_INSERTFILE)
        
            Case mcITEM_REMOVEITEMS
                Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)
        
            Case mcITEM_SAVEONETIMETODB
                Description = .TextMatrix(.Row, .ColIndex("itemdesc"))
                If FItem.Add(Phase, item, Description) Then
                    .TextMatrix(.Row, .ColIndex("estphase")) = Phase
                    .TextMatrix(.Row, .ColIndex("estitem")) = item
                    .TextMatrix(.Row, .ColIndex("itemdesc")) = Description
                    If .RowData(.Row) <> "NEW" Then .RowData(.Row) = "DIRTY"
                    
                    s = ""
                    s = s & "update tblphaseitem" & vbCrLf
                    s = s & "set takeoffuom=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("takeoffuom"))) & vbCrLf
                    s = s & "   ,orderuom=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("orderuom"))) & vbCrLf
                    s = s & "   ,price=" & DbQuote(Num, .TextMatrix(.Row, .ColIndex(IIf(mShowBudgets, "budget", "po") & "rate"))) & vbCrLf
                    s = s & "   ,poindex=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("poindex"))) & vbCrLf
                    s = s & "   ,jccostcode=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("jccostcode"))) & vbCrLf
                    s = s & "   ,jccategory=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("jccategory"))) & vbCrLf
                    s = s & "   ,ustmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
                    s = s & "   ,tstmp=getdate()" & vbCrLf
                    s = s & "   ,taxgroup=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex(IIf(mShowBudgets, "budget", "po") & "taxgroup"))) & vbCrLf
                    s = s & "where phase=" & DbQuote(Str, Phase) & vbCrLf
                    s = s & "and item=" & DbQuote(Str, item) & vbCrLf
                    Call HFApp.SqlExec(s, dbHomefront)
                    
                    Call SaveData(False)
                    
                End If
            
            
            Case mcITEM_CHANGEITEM
'                s = ""
'                s = s & "select isnull(phase,'')+char(1)+isnull(item,'') phaseitem " & vbCrLf
'                s = s & "      ,poindex" & vbCrLf
'                s = s & "      ,Phase " & vbCrLf
'                s = s & "      ,Item " & vbCrLf
'                s = s & "      ,Description " & vbCrLf
'                s = s & "      ,v.vendor_id vendor" & vbCrLf
'                s = s & "      ,v.Vendor_name DefaultVendor" & vbCrLf
'                s = s & "  from tblphaseitem" & vbCrLf
'                s = s & "       left outer join tblvendors v on(vendor_id=dbo.Purch_GetCommunityVendor(" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & ",poindex))" & vbCrLf
'                s = s & " where phase<>" & DbQuote(Str, HFApp.Options(SelectAtTakeoffPhase)) & vbCrLf
'                s = s & "   and item<>" & DbQuote(Str, HFApp.Options(SelectAtTakeoffItem)) & vbCrLf
'                If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s, , , , , "phaseitem,poindex,vendor") Then
'                    'this can only happen when neither pos or budgets have been generated so we only need to set budget values. the afteredit event will set the po values
'                    .TextMatrix(.Row, .ColIndex("EstPhase")) = FPickList.SelectedItem("Phase")
'                    .TextMatrix(.Row, .ColIndex("EstItem")) = FPickList.SelectedItem("Item")
'                    .TextMatrix(.Row, .ColIndex("ItemDesc")) = FPickList.SelectedItem("Description")
'                    .TextMatrix(.Row, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
'                    .TextMatrix(.Row, .ColIndex("BudgetVendor")) = FPickList.SelectedItem("Vendor")
'                    .TextMatrix(.Row, .ColIndex("BudgetVendorName")) = FPickList.SelectedItem("DefaultVendor")
'                    .TextMatrix(.Row, .ColIndex("POOverridden")) = "False"
'                    .TextMatrix(.Row, .ColIndex("BudgetOverridden")) = "False"
'                    Call gItems_AfterEdit(.Row, .ColIndex("BudgetVendor"))
'                End If
                s = ""
                s = s & "select isnull(phase,'')+char(1)+isnull(item,'') phaseitem " & vbCrLf
                s = s & "      ,poindex" & vbCrLf
                s = s & "      ,Phase " & vbCrLf
                s = s & "      ,Item " & vbCrLf
                s = s & "      ,Description " & vbCrLf
                s = s & "  from tblphaseitem" & vbCrLf
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s, .TextMatrix(.Row, .ColIndex("EstPhase")) & Chr(1) & .TextMatrix(.Row, .ColIndex("EstItem")), , , , "phaseitem,poindex") Then
                    'this can only happen when neither pos or budgets have been generated so we only need to set budget values. the afteredit event will set the po values
                    .TextMatrix(.Row, .ColIndex("EstPhase")) = FPickList.SelectedItem("Phase")
                    .TextMatrix(.Row, .ColIndex("EstItem")) = FPickList.SelectedItem("Item")
                    .TextMatrix(.Row, .ColIndex("ItemDesc")) = FPickList.SelectedItem("Description")
                    .TextMatrix(.Row, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                    .TextMatrix(.Row, .ColIndex("POOverridden")) = "False"
                    .TextMatrix(.Row, .ColIndex("BudgetOverridden")) = "False"
                End If
                
            
            Case mcITEM_COPYITEMS
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    Dirty = True
                    
                    newrow = Max(.Row, .RowSel) + 1
                    
                    .AddItem "", newrow
                    
                    .RowData(newrow) = "NEW"
                    For c = 0 To .cols - 1
                        .TextMatrix(newrow, c) = .TextMatrix(r, c)
                    Next
                    .TextMatrix(newrow, .ColIndex("BudgetQty")) = 0
                    .TextMatrix(newrow, .ColIndex("POQty")) = 0
                    .TextMatrix(newrow, .ColIndex("PONumber")) = ""
                    .TextMatrix(newrow, .ColIndex("EstItemID")) = "0"
                    .TextMatrix(newrow, .ColIndex("BudgetGenerated")) = "False"
                    .TextMatrix(newrow, .ColIndex("BudgetPostingBatch")) = "0"
                    Call ColorizeItems(newrow)
                Next
                       
            
                
            Case mcITEM_COMPAREPRICES
                r = .Row
                If r < 1 Then Exit Sub
                Call FPriceComparison.Showform(.TextMatrix(r, .ColIndex(IIf(mShowBudgets, "BudgetVendor", "POVendor"))), _
                                               GetComboBoxListKey(cboCommunity), _
                                               cboPhase.Text, _
                                               .TextMatrix(r, .ColIndex("Assembly")), _
                                               .TextMatrix(r, .ColIndex("Model")), _
                                               .TextMatrix(r, .ColIndex("EstPhase")), _
                                               .TextMatrix(r, .ColIndex("EstItem")), _
                                               .TextMatrix(r, .ColIndex("ItemDesc")))
                        
            
            
            Case mcITEM_UPDATEPRICELIST
                r = .Row
                If r < 1 Then Exit Sub
                Call FPriceListUpdate.Showform(mShowBudgets, _
                                               .TextMatrix(r, .ColIndex("Assembly")), _
                                               GetComboBoxListKey(cboCommunity), _
                                               cboPhase.Text, _
                                               .TextMatrix(r, .ColIndex(IIf(mShowBudgets, "BudgetVendor", "POVendor"))), _
                                               .TextMatrix(r, .ColIndex(IIf(mShowBudgets, "BudgetVendorName", "POVendorName"))), _
                                               .TextMatrix(r, .ColIndex("EstPhase")), _
                                               .TextMatrix(r, .ColIndex("EstItem")), _
                                               .TextMatrix(r, .ColIndex("ItemDesc")), _
                                               .ValueMatrix(r, .ColIndex(IIf(mShowBudgets, "BudgetRate", "PORate"))), _
                                               .TextMatrix(r, .ColIndex("OrderUOM")))
            
        End Select
    End With
End Sub



Public Function SaveItems() As Boolean
On Error GoTo eh

    Dim i As Long
    Dim s As String
    
    Dim CanDeleteBudget As Boolean
    Dim CanDeletePO As Boolean
                        
    If Not mDirty Then
        SaveItems = True
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    With gItems
        For i = 1 To .Rows - 1
            If .RowData(i) = "NEW" Then
                s = ""
                s = s & "INSERT INTO EstimateItems(EstAssemblyID,Assembly,AssemblyDescription,POIndex,Model,Phase,Item,Job,JCExtra,JCCostCode,JCCategory,SortOrder,Description,Comments,TakeoffQty,TakeoffUOM,ConversionFactor,OrderUOM,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate,BudgetGenerated,POVendor,POQty,PORate,POPretax,POTaxGroup,POJCTax,POJCTaxRate,PONJCTax,PONJCTaxRate,PONumber,POOverridden,BudgetOverridden,ExcludeFromPO,OriginalJCCategory)" & vbCrLf
                s = s & "VALUES(" & DbQuote(Num, mEstAssemblyID) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("AssemblyDescription"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("POIndex"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Model"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstPhase"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstItem"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, mJob) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCExtra"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, i) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemDesc"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemComments"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TakeoffUOM"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("ConversionFactor"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetVendor"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetQty"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetPretax"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetTaxGroup"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTax"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTax"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetGenerated"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("POVendor"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("POQty"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("PORate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("POPretax"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("POTaxGroup"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("POJCTax"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("POJCTaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("PONJCTax"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("PONJCTaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("PONumber"))) & vbCrLf
                
                s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("POOverridden")) = "True") & vbCrLf
                s = s & "      ," & DbQuote(Bit, .TextMatrix(i, .ColIndex("BudgetOverridden")) = "True") & vbCrLf
                
                s = s & "      ," & IIf(.Cell(flexcpChecked, i, .ColIndex("ExcludeFromPO")) = flexChecked, 1, 0) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("OriginalJCCategory"))) & ")"
                HFApp.SqlExec s
                .TextMatrix(i, .ColIndex("EstItemID")) = HFApp.SqlIdentity("EstimateItems")
                .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexUnchecked
                .RowData(i) = ""
            End If
            
            If .RowData(i) = "DIRTY" Then
                s = ""
                s = s & "UPDATE EstimateItems" & vbCrLf
                s = s & "SET Assembly=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf
                s = s & "   ,AssemblyDescription=" & DbQuote(Str, .TextMatrix(i, .ColIndex("AssemblyDescription"))) & vbCrLf
                s = s & "   ,POIndex=" & DbQuote(Str, .TextMatrix(i, .ColIndex("POIndex"))) & vbCrLf
                s = s & "   ,JCExtra=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCExtra")), , True) & vbCrLf
                s = s & "   ,JCCostCode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
                s = s & "   ,JCCategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemDesc"))) & vbCrLf
                s = s & "   ,Comments=" & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemComments"))) & vbCrLf
                s = s & "   ,OrderUOM=" & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM"))) & vbCrLf
                s = s & "   ,BudgetVendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetVendor"))) & vbCrLf
                s = s & "   ,BudgetQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetQty"))) & vbCrLf
                s = s & "   ,BudgetRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetRate"))) & vbCrLf
                s = s & "   ,BudgetPretax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetPretax"))) & vbCrLf
                s = s & "   ,BudgetTaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetTaxGroup"))) & vbCrLf
                s = s & "   ,BudgetJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTax"))) & vbCrLf
                s = s & "   ,BudgetJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTaxRate"))) & vbCrLf
                s = s & "   ,BudgetNJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTax"))) & vbCrLf
                s = s & "   ,BudgetNJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTaxRate"))) & vbCrLf
                s = s & "   ,POVendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("POVendor"))) & vbCrLf
                s = s & "   ,POQty=" & DbQuote(Num, .TextMatrix(i, .ColIndex("POQty"))) & vbCrLf
                s = s & "   ,PORate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("PORate"))) & vbCrLf
                s = s & "   ,POPretax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("POPretax"))) & vbCrLf
                s = s & "   ,POTaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("POTaxGroup"))) & vbCrLf
                s = s & "   ,POJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("POJCTax"))) & vbCrLf
                s = s & "   ,POJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("POJCTaxRate"))) & vbCrLf
                s = s & "   ,PONJCTax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("PONJCTax"))) & vbCrLf
                s = s & "   ,PONJCTaxRate=" & DbQuote(Num, .TextMatrix(i, .ColIndex("PONJCTaxRate"))) & vbCrLf
                s = s & "   ,ExcludeFromPO=" & IIf(.Cell(flexcpChecked, i, .ColIndex("ExcludeFromPO")) = flexChecked, 1, 0) & vbCrLf
                s = s & "   ,BudgetOverridden=" & DbQuote(Bit, .TextMatrix(i, .ColIndex("BudgetOverridden")) = "True") & vbCrLf
                s = s & "   ,POOverridden=" & DbQuote(Bit, .TextMatrix(i, .ColIndex("POOverridden")) = "True") & vbCrLf
                s = s & "WHERE EstItemID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstItemID"))) & vbCrLf
                Call HFApp.SqlExec(s)
                .RowData(i) = ""
            End If
        Next
    
        For i = .Rows - 1 To 1 Step -1
            If .RowData(i) = "DELETE" Then
            
                CanDeleteBudget = .TextMatrix(i, .ColIndex("BudgetGenerated")) <> "True"
                CanDeletePO = .TextMatrix(i, .ColIndex("PONumber")) = ""
                Select Case True
                    Case CanDeleteBudget And CanDeletePO
                        s = ""
                        s = s & "DELETE FROM EstimateItems" & vbCrLf
                        s = s & "WHERE EstItemID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstItemID"))) & vbCrLf
                    Case CanDeleteBudget
                        s = ""
                        s = s & "UPDATE EstimateItems SET BudgetDeleted=1" & vbCrLf
                        s = s & "WHERE EstItemID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstItemID"))) & vbCrLf
                    Case CanDeletePO
                        s = ""
                        s = s & "UPDATE EstimateItems SET PODeleted=1" & vbCrLf
                        s = s & "WHERE EstItemID=" & DbQuote(Num, .TextMatrix(i, .ColIndex("EstItemID"))) & vbCrLf
                End Select
                
                
                Call HFApp.SqlExec(s)
                Call .RemoveItem(i)
                
            End If
        Next
    
    End With
    SaveItems = True
    Dirty = False
    Screen.MousePointer = vbDefault

Exit Function
eh: Call errHandler(srcFile & "SaveItems", s)
End Function




Private Function WhereClause(Optional prefix As String)
    Dim i As Long
    Dim s As String
    
    If prefix <> "" Then prefix = prefix & "."
    
    'add items to where clause
    With gItems
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked Then
                s = s & "," & .TextMatrix(i, .ColIndex("EstItemID"))
            End If
        Next
        If s <> "" Then s = " OR " & prefix & "EstItemID IN(" & Mid(s, 2) & ")"
    End With
    
    
    
    
    If s = "" Then
        WhereClause = "1=2"
    Else
        WhereClause = "(" & prefix & "EstItemID IS NOT NULL AND " & Mid(s, 5) & ")"
    End If


    
End Function

Private Sub LoadItems()
On Error GoTo eh
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    
    With gItems
        .Redraw = flexRDNone
        .Rows = 1
                
        s = "SELECT * FROM EstimatedItems WHERE Job_no=" & DbQuote(Str, mJob)
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
        
            mEstAssemblyID = Val("" & rs("EstAssemblyID"))
            If Val("" & rs("EstAssemblyID")) <> 0 Then
    
                'ignore deleted items
                If (mShowBudgets And Not rs("BudgetDeleted")) Or ((Not mShowBudgets) And Not rs("PODeleted")) Then
                    
                    .AddItem ""
                    r = .Rows - 1
                    For c = 0 To .cols - 1
                        Select Case .ColKey(c)
                        
                            Case "TakeoffQty"
                                'takeoff qty in db always shows original takeoff amounts
                                'on screen we want it to show the budget or po qty in takeoff uom
                                'If Val("" & rs("ConversionFactor")) <> 1 Then Stop
                                If Val("" & rs("ConversionFactor")) = 0 Then
                                    .Cell(flexcpText, r, c) = Val("" & rs(IIf(mShowBudgets, "BudgetQty", "POQty")))
                                Else
                                    .Cell(flexcpText, r, c) = Val("" & rs(IIf(mShowBudgets, "BudgetQty", "POQty"))) / Val("" & rs("ConversionFactor"))
                                End If
                                
                                
                            Case "ExcludeFromPO"
                                .Cell(flexcpChecked, r, c) = IIf(rs(.ColKey(c)), flexChecked, flexUnchecked)
                            Case Else
                                On Error Resume Next
                                .Cell(flexcpText, r, c) = "" & rs(.ColKey(c))
                                On Error GoTo 0
                        End Select
                    Next
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    .TextMatrix(r, .ColIndex("POTax")) = .ValueMatrix(r, .ColIndex("POJCTax")) + .ValueMatrix(r, .ColIndex("PONJCTax"))
                    
                End If
            End If
            rs.MoveNext
        Wend
 
        Call ColorizeItems(-1)
        .Redraw = flexRDBuffered
    End With
    Dirty = False
    
Exit Sub
eh: Call errHandler(srcFile & "LoadItems", s)
End Sub


Private Function CleanJob(FormatedJob As String) As String
    CleanJob = Trim(Replace(Replace(Replace(Replace(Replace(FormatedJob, "\", ""), ",", ""), "/", ""), ".", ""), "-", ""))
End Function


Private Function ValidateJobNumber(Job As String) As Boolean
    Select Case HFApp.Options(AccountingSystem)
    
        Case asTimberline
            ValidateJobNumber = IsBetween(Len(CleanJob(Job)), Val(HFApp.Options(Job_Section2)) + Val(HFApp.Options(Job_Section3)) + 1, Val(HFApp.Options(Job_Section1)) + Val(HFApp.Options(Job_Section2)) + Val(HFApp.Options(Job_Section3)))
            
        Case asMasterBuilder
            If IsNumeric(Job) And Val(Job) - Int(Val(Job)) = 0 Then
                ValidateJobNumber = True
                Job = Int(Val(Job))
            End If
            
        Case Else
            ValidateJobNumber = True
            
    End Select
End Function

