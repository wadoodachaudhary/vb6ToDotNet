VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FInboxJobs 
   Caption         =   "Prepare Estimates"
   ClientHeight    =   6135
   ClientLeft      =   1305
   ClientTop       =   4650
   ClientWidth     =   8400
   Icon            =   "FInboxJobs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdNav 
      Caption         =   "Send to &Pipeline"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   5655
      Width           =   1485
   End
   Begin HFEst.WizHead WizHead 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   1588
      Caption         =   "Prepare Estimates for sold items"
      Description     =   "Select jobs and contract items to be budgeted and purchased."
      Icon            =   "FInboxJobs.frx":000C
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   7245
      TabIndex        =   2
      Top             =   5655
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   6165
      TabIndex        =   1
      Top             =   5655
      Width           =   1035
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   8175
      _cx             =   14420
      _cy             =   7858
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FInboxJobs.frx":08E6
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   6
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Image imgShortcut 
         Height          =   240
         Index           =   0
         Left            =   1545
         Picture         =   "FInboxJobs.frx":0AD7
         Stretch         =   -1  'True
         ToolTipText     =   "Expand All"
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "FInboxJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const SRCFILE = "FInboxJobs::"
Private mDirty   As Boolean


Const c_Description = 0
Const c_Inverted = 1
Const c_Change_Order_No = 2
Const c_JCCat = 3
Const c_Warnings = 4
Const c_Customer_No = 5
Const c_Location = 6
Const c_seq = 7
Const c_Job = 8
Const c_Model = 9
Const c_community = 10
Const c_Quantity = 11
Const c_OptionID = 12
Const c_Notes = 13

Private Sub cmdNav_Click(Index As Integer)
    Dim JobNo As String
    Select Case Index
        
        Case 0, 2 'ok / send to pipeline
            JobNo = SaveData(Index = 2)
            If JobNo <> "" Then
                Unload Me
                Call FEstimateItems.ShowForm(fceBudget, JobNo)
            End If
            
        Case 1 'cancel
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    
    cmdNav(2).Visible = HFApp.Options.ValueByName("EstimatingSystem") = esPipeline
    
    Call LoadData
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    gData.Move margin, WizHead.Height + margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - WizHead.Height - margin * 3 - cmdNav(0).Height
    cmdNav(0).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
    cmdNav(1).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
    cmdNav(2).Move margin, Me.ScaleHeight - margin - cmdNav(0).Height
End Sub



Private Function SaveData(SendToPipeline As Boolean) As String
On Error GoTo eh
'used to return boolean
'now returns the first job number processed

    Dim r As Long
    Dim s As String
    Dim firstJob As String
    Dim lastCustomer  As String: lastCustomer = Chr(1)
    Dim EstimateIndex As Long
    Dim InboxBatchID As String:  InboxBatchID = CreateGUID()

    Dim X As New PipelineWrapper.PipelineWrapper
    Dim optLocation As Long
    Dim optNo As String
    Dim optDesc As String
    Dim optQty As String
    Dim optNotes As String
    Dim rs As Recordset
    
    
    
    Dim PipelineCommunitySpecific As Boolean
    
    If SendToPipeline Then
        PipelineCommunitySpecific = HFApp.Options.ValueByName("BIMCommunitySpecificPlansAndOptions") <> "False"
        Call X.Connect(HFApp.Options.ValueByName("CGVisionsRootURI"), HFApp.Options.ValueByName("CGVisionsUID"), HFApp.Options.ValueByName("CGVisionsPWD"))
    End If
    
    
    firstJob = ""
    With gData
        For r = 0 To .Rows - 1
            If .RowOutlineLevel(r) = 3 And .Cell(flexcpChecked, r, 0) = flexChecked Then

                If lastCustomer <> .TextMatrix(r, c_Customer_No) Then
                 
                    If SendToPipeline Then
                        If lastCustomer <> Chr(1) Then
                            Call X.SaveJob(EstimateIndex + 1)
                        End If
                        
                        If PipelineCommunitySpecific Then
                            Call X.AddJob(.TextMatrix(r, c_Job), .TextMatrix(r, c_Model), .TextMatrix(r, c_community))
                        Else
                            Call X.AddJob(.TextMatrix(r, c_Job), .TextMatrix(r, c_Model), "0000")
                        End If
                    End If
                 
                    lastCustomer = .TextMatrix(r, c_Customer_No)
                    EstimateIndex = HFApp.SqlExec("exec Purch_CreateCustomerEstimate " & DbQuote(Str, lastCustomer) & "," & DbQuote(Str, HFApp.LoginID))(0)
                                        
                End If

                If firstJob = "" Then firstJob = .TextMatrix(r, c_Job)
                                
                optLocation = .ValueMatrix(r, c_Location)
                If SendToPipeline And optLocation <> 0 Then
                    optNo = .TextMatrix(r, c_OptionID)
                    optDesc = .TextMatrix(r, c_Description)
                    optQty = .TextMatrix(r, c_Quantity)
                    optNotes = .TextMatrix(r, c_Notes)
                End If
                
                s = "exec dbo.Purch_PutEstimateAssembly " & DbQuote(Str, .TextMatrix(r, c_Customer_No)) & _
                                                      "," & DbQuote(Num, EstimateIndex) & _
                                                      "," & DbQuote(Num, .ValueMatrix(r, c_Location)) & _
                                                      "," & DbQuote(Str, .TextMatrix(r, c_Change_Order_No)) & _
                                                      "," & DbQuote(Num, .ValueMatrix(r, c_seq)) & _
                                                      "," & DbQuote(Str, "" & .TextMatrix(r, c_JCCat)) & _
                                                      "," & DbQuote(Str, HFApp.LoginID) & _
                                                      "," & DbQuote(Num, HFApp.DivisionID) & _
                                                      "," & DbQuote(Bit, Not SendToPipeline) & _
                                                      "," & DbQuote(Str, InboxBatchID)

                Set rs = HFApp.SqlExec(s)



                If SendToPipeline And optLocation <> 0 Then
                    Call X.AddJobOption(Val("" & rs("EstAssemblyID")), optNo, optDesc, optQty, optNotes)
                End If
            
            End If
        Next
    End With
    
    If SendToPipeline And lastCustomer <> Chr(1) Then
        Call X.SaveJob(EstimateIndex + 1)
    End If
    SaveData = firstJob


    'set estimateindex in crm
    Call Sales_SetEstimateIndex(InboxBatchID)



Exit Function
eh: Call errHandler(SRCFILE & "SaveData")
End Function

Private Sub LoadData()

Dim rs As Recordset
Dim r As Long
Dim s As String
Dim lastCommunity As String:    lastCommunity = Chr(1)
Dim lastJob       As String:    lastJob = Chr(1)
Dim lastCustomer  As String:    lastCustomer = Chr(1)
    
    
    With gData
        .Redraw = flexRDNone
        .Rows = 1
        
        r = 0
        
        s = ""
        s = s & "SELECT o.FullCommunityDesc" & vbCrLf
        s = s & "      ,o.Job JobNo" & vbCrLf
        s = s & "      ,ISNULL(o.Job,'') + ISNULL(' - ' + o.JobDesc,'') Job" & vbCrLf
        s = s & "      ,ISNULL(o.CustomerDesc,Customer_No) + '   (' + ISNULL(o.Model,'') + ' ' + ISNULL(o.Elevation,'') + ' ' + ISNULL(o.Series,'') + ')' Customer" & vbCrLf
        s = s & "      ,ISNULL(o.Municipal_Address,'address not entered') Address" & vbCrLf
        s = s & "      ,'Base Model' AssemblyDesc" & vbCrLf
        s = s & "      ,o.Customer_No" & vbCrLf
        s = s & "      ,0 Location" & vbCrLf
        s = s & "      ,'' Change_order_No" & vbCrLf
        s = s & "      ,0 Seq" & vbCrLf
        s = s & "      ,o.modelassembly assembly" & vbCrLf
        s = s & "      ,case when isnull(o.modelassembly,'')='' then 1" & vbCrLf
        s = s & "            when isnull(a.assembly,'')='' then 2" & vbCrLf
        s = s & "            else 0 end Quality" & vbCrLf
        'these are for pipeline only...
        s = s & ",o.Model,o.community,null optionid,null quantity,null notes" & vbCrLf
        s = s & ",o.InvertedFloorPlan" & vbCrLf
        s = s & "  FROM UnEstimatedCustomers o " & vbCrLf
        s = s & "       join system_setup s ON(s.id = " & HFApp.DivisionID & ")" & vbCrLf
        s = s & "       left outer join tbldbassemblymaster a " & vbCrLf
        s = s & "          on (a.optionid=''" & vbCrLf
        s = s & "          and a.assembly=o.modelassembly " & vbCrLf
        s = s & "          and a.DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & "          and isnull(a.Community,'')=isnull(dbo.Purch_GetAssemblyCommunity(o.modelassembly,o.model,'',o.community," & HFApp.DivisionID & "),'')" & vbCrLf
        s = s & "          and a.Model=o.Model)" & vbCrLf
        s = s & "       join divisioncommunities d on d.Community=o.Community where d.DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & " and s.ID = " & HFApp.DivisionID & vbCrLf
        s = s & "UNION All" & vbCrLf
        s = s & "SELECT o.FullCommunityDesc" & vbCrLf
        s = s & "      ,o.Job JobNo" & vbCrLf
        s = s & "      ,ISNULL(o.Job,'') + ISNULL(' - ' + o.JobDesc,'') Job" & vbCrLf
        s = s & "      ,ISNULL(o.CustomerDesc,o.Customer_No) + '   (' + ISNULL(o.Model,'') + ' ' + ISNULL(o.Elevation,'') + ' ' + ISNULL(o.Series,'') + ')' Customer" & vbCrLf
        s = s & "      ,ISNULL(o.Municipal_Address,'address not entered') Address" & vbCrLf
        s = s & "      ,left(o.Description,50) AssemblyDesc" & vbCrLf
        s = s & "      ,o.Customer_No" & vbCrLf
        s = s & "      ,o.Location" & vbCrLf
        s = s & "      ,o.Change_order_No" & vbCrLf
        s = s & "      ,o.Seq" & vbCrLf
        s = s & "      ,o.assembly" & vbCrLf
        s = s & "      ,case when isnull(o.assembly,'')='' then   1" & vbCrLf
        s = s & "            when isnull(a.assembly,'')='' then   2" & vbCrLf
        s = s & "            when isnull(o.assembly,'')='CUSTOM' then   3" & vbCrLf
        s = s & "            else 0 end Quality" & vbCrLf
        'these are for pipeline only...
        s = s & ",o.Model,o.community,o.opt optionid,o.qty quantity,o.estimatornotes notes" & vbCrLf
        s = s & ",o.InvertedFloorPlan" & vbCrLf
        s = s & "  FROM UnEstimatedOptions o" & vbCrLf
        s = s & "       join system_setup s ON(s.id = " & HFApp.DivisionID & ")" & vbCrLf
        If HFApp.DivisionID <> "" Then
            s = s & "       join divisioncommunities d on d.Community=o.Community" & vbCrLf
        End If
        s = s & "       left outer join tbldbassemblymaster a " & vbCrLf
        s = s & "          on (a.optionid=o.opt" & vbCrLf
        s = s & "          and a.assembly=o.assembly " & vbCrLf
        s = s & "          and a.DivisionID = " & HFApp.DivisionID
        s = s & "          and ISNULL(a.Model,'') = CASE WHEN (o.AssemblyType=3 OR o.AssemblyType=4) THEN '' ELSE o.Model END" & vbCrLf
        s = s & "          and isnull(a.Community,'')=isnull(dbo.Purch_GetAssemblyCommunity(o.assembly,CASE WHEN (o.AssemblyType=3 OR o.AssemblyType=4) THEN '' ELSE o.Model END,o.opt,o.community," & HFApp.DivisionID & "),''))" & vbCrLf
        s = s & " where s.id = " & HFApp.DivisionID & " and d.DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & "ORDER BY 1,3,4,8,6" 'Community, Job, Customer, o.Location, assemblydesc" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        
        While Not rs.EOF
            
            If lastCommunity <> "" & rs("Community") Then
                lastCommunity = "" & rs("Community")
                r = r + 1
                .AddItem "" & rs("FullCommunityDesc")
                Set .Cell(flexcpPicture, r, 0) = FMain.SmallIcons.ListImages("area").Picture
                .RowOutlineLevel(r) = 0
                .IsSubtotal(r) = True
            End If
            
            If lastJob <> "" & rs("JobNo") Then
                lastJob = "" & rs("JobNo")
                r = r + 1
                .AddItem HFApp.FormatJob(lastJob)
                Set .Cell(flexcpPicture, r, 0) = FMain.SmallIcons.ListImages("job").Picture
                .RowOutlineLevel(r) = 1
                .IsSubtotal(r) = True
            End If
            
            If lastCustomer <> "" & rs("Customer_No") Then
                lastCustomer = "" & rs("Customer_No")
                r = r + 1
                .AddItem "" & rs("Customer")
                
                .TextMatrix(r, c_Inverted) = IIf("" & rs("InvertedFloorPlan") = "True", "Yes", "No")
                .TextMatrix(r, c_Customer_No) = Trim("" & rs("Customer_No"))

                .RowOutlineLevel(r) = 2
                .IsSubtotal(r) = True
            End If
            

            r = r + 1
            .AddItem ""
            
            .TextMatrix(.Rows - 1, c_Description) = Trim("" & rs("AssemblyDesc"))
            .TextMatrix(.Rows - 1, c_Change_Order_No) = Trim("" & rs("Change_Order_No"))
            .TextMatrix(.Rows - 1, c_Customer_No) = Trim("" & rs("Customer_No"))
            .TextMatrix(.Rows - 1, c_Location) = Trim("" & rs("Location"))
            .TextMatrix(.Rows - 1, c_seq) = Trim("" & rs("Seq"))
            .TextMatrix(.Rows - 1, c_Job) = Trim("" & rs("JobNo"))
            
            Set .Cell(flexcpPicture, r, 0) = FMain.SmallIcons.ListImages("option").Picture
            
            Select Case Val("" & rs("quality"))
                
                Case 3: Set .Cell(flexcpPicture, r, c_Warnings) = FMain.SmallIcons.ListImages("custom").Picture
                            .TextMatrix(r, c_Warnings) = "custom option"
                
                Case 2: Set .Cell(flexcpPicture, r, c_Warnings) = FMain.SmallIcons.ListImages("warning").Picture
                            .TextMatrix(r, c_Warnings) = "assembly " & vbQuote & rs("assembly") & vbQuote & " could not be found"
                            
                Case 1: Set .Cell(flexcpPicture, r, c_Warnings) = FMain.SmallIcons.ListImages("question").Picture
                        .TextMatrix(r, c_Warnings) = "assembly not specified"

            End Select
            
            'for pipeline
            .TextMatrix(.Rows - 1, c_Model) = "" & rs("model")
            .TextMatrix(.Rows - 1, c_community) = "" & rs("community")
            .TextMatrix(.Rows - 1, c_Quantity) = "" & rs("quantity")
            .TextMatrix(.Rows - 1, c_OptionID) = "" & rs("optionid")
            .TextMatrix(.Rows - 1, c_Notes) = "" & rs("notes")
            
            .RowOutlineLevel(r) = 3
            .IsSubtotal(r) = True
            
            
            rs.MoveNext
        Wend
        
        Call .AutoSize(0, c_JCCat)
        On Error Resume Next
        .Cell(flexcpChecked, 0, 0, .Rows - 1, 0) = flexUnchecked
        .Outline 2
        .Outline 1
        .Outline 0
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub gData_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
    Call gData.AutoSize(0, 2)
End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    
    With gData
        If Row < 0 Then Exit Sub
        If Row = 0 Then
            .Cell(flexcpChecked, Row, 0, .Rows - 1, 0) = .Cell(flexcpChecked, Row, 0)
        Else
            If .Cell(flexcpChecked, Row, 0) = flexTSGrayed Then .Cell(flexcpChecked, Row, 0) = flexUnchecked
            Call ResetChildrenChecked(Row)
            Call ResetParentChecked(Row)
        End If
        
        For r = 1 To .Rows - 1
            If .RowOutlineLevel(r) <> 3 Then
                .TextMatrix(r, 2) = ""
            End If
        Next
    End With
End Sub

Private Sub ResetChildrenChecked(ByVal ParentRow As Long)
    Dim r As Long
    If ParentRow < 0 Then Exit Sub
    With gData
        'for all my children set checkbox to same as me
        r = .GetNodeRow(ParentRow, flexNTFirstChild)
        While r <> -1
            .Cell(flexcpChecked, r, 0) = .Cell(flexcpChecked, ParentRow, 0)
            Call ResetChildrenChecked(r)
            r = .GetNodeRow(r, flexNTNextSibling)
        Wend
    End With
End Sub

Private Sub ResetParentChecked(ByVal ChildRow As Long)
    Dim ParentRow As Long
    Dim r As Long
    With gData
        If ChildRow = -1 Then Exit Sub
        ParentRow = .GetNodeRow(ChildRow, flexNTParent)
        If ParentRow = -1 Then Exit Sub
        
        
        If .Cell(flexcpChecked, ParentRow, 0) = flexTSGrayed Then .Cell(flexcpChecked, ParentRow, 0) = flexUnchecked
        .Cell(flexcpChecked, ParentRow, 0) = .Cell(flexcpChecked, ChildRow, 0)
        
        r = .GetNodeRow(ChildRow, flexNTFirstSibling)
        While r <> -1
            If .Cell(flexcpChecked, r, 0) <> .Cell(flexcpChecked, ChildRow, 0) Then
                .Cell(flexcpChecked, ParentRow, 0) = flexTSGrayed
                r = -1
            Else
                r = .GetNodeRow(r, flexNTNextSibling)
            End If
        Wend
        Call ResetParentChecked(ParentRow)
    End With
End Sub

Private Sub gData_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
Static inHere As Boolean
If inHere Then Exit Sub
inHere = True
    gData.ColSel = gData.Col
inHere = False
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    Dim rs As Recordset
    With gData
        .AutoSearch = flexSearchFromCursor
        .ComboList = ""
        Select Case Col
            Case c_Description: ' checkbox
            
            Case c_Inverted:
                If .RowOutlineLevel(Row) = 2 Then
                    .AutoSearch = flexSearchNone
                    .ComboList = "Yes|No"
                Else
                    .ComboList = ""
                    Cancel = True
                End If
            
            Case c_JCCat: ' jc category
                If .RowOutlineLevel(Row) = 3 Then
                    s = "SELECT Category,Description FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                    Set rs = HFApp.SqlExec(s, dbHomefront)
                    .ComboList = " |" & .BuildComboList(rs, "Category,Description")
                    .AutoSearch = flexSearchNone
                Else
                    Cancel = True
                End If
                
                
            Case Else:
                Cancel = True
                
        End Select
    End With
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gData
        If .Row < 0 Then Exit Sub
        If .Col <> 0 And .Col <> 2 Then Exit Sub
        Select Case KeyCode
            
            Case vbKeyDelete
                If .Col = 2 Then
                    .Text = ""
                End If
                
            Case vbKeySpace
                .IsCollapsed(.Row) = IIf(.IsCollapsed(.Row) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)

            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If

            Case vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                        .Col = 0
                        KeyCode = 0
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If
                
        End Select
    End With
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    With gData
        Select Case Col
            Case c_Inverted:
                If .RowOutlineLevel(Row) = 2 Then
                    s = "update tblcustomers set InvertedFloorPlan=" & DbQuote(Bit, .EditText) & " where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and customer_no=" & DbQuote(Str, .TextMatrix(Row, c_Customer_No))
                    HFApp.SqlExec s
                End If
        End Select
    End With
End Sub

Private Sub imgShortcut_Click(Index As Integer)
    Call GridExpandALL(Me.gData)
End Sub

