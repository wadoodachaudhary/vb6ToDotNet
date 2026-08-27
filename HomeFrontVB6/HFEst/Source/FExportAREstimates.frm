VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FExportAREstimates 
   Caption         =   "AR Estimate Posting Wizard"
   ClientHeight    =   4785
   ClientLeft      =   6915
   ClientTop       =   2595
   ClientWidth     =   6330
   ControlBox      =   0   'False
   Icon            =   "FExportAREstimates.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6330
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3270
      Index           =   2
      Left            =   8145
      TabIndex        =   13
      Top             =   4965
      Width           =   6315
      Begin VB.Frame frmReview 
         BorderStyle     =   0  'None
         Height          =   2835
         Left            =   1140
         TabIndex        =   14
         Top             =   420
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label lblBatchTask 
            AutoSize        =   -1  'True
            Caption         =   "Open the posting summary report"
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
            Index           =   0
            Left            =   420
            TabIndex        =   15
            Top             =   210
            Width           =   2325
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FExportAREstimates.frx":000C
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         Caption         =   "The posting wizard is ready to proceed."
         Height          =   195
         Left            =   1140
         TabIndex        =   16
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   3360
      Index           =   1
      Left            =   3570
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   2670
         Left            =   1140
         TabIndex        =   12
         Top             =   420
         Width           =   4905
         _cx             =   1986929100
         _cy             =   1986925158
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
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   -2147483635
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FExportAREstimates.frx":08D6
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   3
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
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
      Begin VB.Image imgError 
         Height          =   480
         Left            =   300
         Picture         =   "FExportAREstimates.frx":0912
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblListDescription 
         Caption         =   "Select the budgets you would like to post."
         Height          =   255
         Left            =   1140
         TabIndex        =   1
         Top             =   120
         Width           =   3915
      End
   End
   Begin HFEst.WizHead WizHead 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1588
      Caption         =   "Post AR Estimates to Accounting"
      Description     =   "The posting wizard will send AR estimates to accounting"
      Icon            =   "FExportAREstimates.frx":11DC
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6330
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4200
      Width           =   6330
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1590
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2790
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   3930
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5130
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   0
         X2              =   26540
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3360
      Index           =   0
      Left            =   765
      TabIndex        =   8
      Top             =   1380
      Width           =   6315
      Begin VB.OptionButton chkReview 
         Caption         =   "Review a previous posting"
         Height          =   255
         Left            =   1140
         TabIndex        =   11
         Top             =   660
         Width           =   4455
      End
      Begin VB.OptionButton chkPostNew 
         Caption         =   "Post new estimates"
         Height          =   255
         Left            =   1140
         TabIndex        =   10
         Top             =   420
         Value           =   -1  'True
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Choose the task you would like to perform"
         Height          =   255
         Index           =   0
         Left            =   1140
         TabIndex        =   9
         Top             =   120
         Width           =   3915
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FExportAREstimates.frx":1AB6
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "FExportAREstimates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FExportAREstimates::"


Private mBatch As Long
Private WhereClause As String


Private Sub LoadTasks()
On Error GoTo eh

    If chkReview.value Then
        frmReview.Visible = True
        mBatch = gData.ValueMatrix(gData.Row, 0)
    Else
        frmReview.Visible = False
        cmdNav(3).Enabled = True
    End If
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadData")
End Sub


Private Sub LoadData()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As ADODB.Recordset
    Dim ra As Long
    Dim IsUSVersion As Boolean
    
    IsUSVersion = HFApp.Options.ValueByName("AccountingVersion") <> "CA"
    
    Screen.MousePointer = vbHourglass
    
    
    Select Case True
        Case chkReview.value = True
            s = ""
            s = s & "SELECT '' EstAssemblyID" & vbCrLf 'hidden column used by posting grid
            s = s & "      ,Batch" & vbCrLf
            s = s & "      ,TStmp Posted" & vbCrLf
            s = s & "      ,UStmp PostedBy" & vbCrLf
            s = s & "      ,''" & vbCrLf 'used by posting grid
            s = s & "  FROM Batches" & vbCrLf
            s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and BatchType='AR Estimate Posting'" & vbCrLf
            s = s & "ORDER BY Batch DESC" & vbCrLf
            lblListDescription.Caption = "Select the batch you would like to review."
            gData.AllowSelection = False
        Case Else
            s = ""
            s = s & "select" & vbCrLf
            s = s & "  a.EstAssemblyID" & vbCrLf
            s = s & "  ,c.area Community,j.CommunityPhase Phase,c.Description,j.job_no Job,j.Description [Job Description],a.HFDescription Item" & vbCrLf
            s = s & "  ,a.SalesQty Qty,a.SalesRate*m.BillingFactor Rate,a.SalesQty*a.SalesRate*m.BillingFactor Amount,cc.description BillingCode,cat.description BillingCategory" & vbCrLf
            s = s & "  ,case when cc.costcode is null then 'invalid billing code, ' else '' end" & vbCrLf
            s = s & "   +case when cat.category is null then 'invalid billing category, ' else '' end" & vbCrLf
            If IsUSVersion = False Then
                s = s & "   +case when isnull(t.taxgroup,'')='' then 'invalid AR taxgroup, ' else '' end" & vbCrLf
            End If
            s = s & "    ValidationErrors" & vbCrLf
            s = s & "from estimateassemblies a" & vbCrLf
            s = s & "join tbljobs j on a.job=j.job_no and a.divisionid=j.divisionid" & vbCrLf
            s = s & "join tbllocality c on c.area=j.community" & vbCrLf
            s = s & "join tblDBAssemblyMaster m ON" & vbCrLf
            s = s & "  m.DivisionID = a.DivisionID and" & vbCrLf
            s = s & "  m.Community=dbo.Purch_GetAssemblyCommunity(a.Assembly,case when(a.AssemblyType=3 OR a.AssemblyType=4) then '' else a.Model end,a.OptionID,j.Community,a.DivisionID) and" & vbCrLf
            s = s & "  m.Assembly=a.Assembly and" & vbCrLf
            s = s & "  m.OptionID=a.OptionID and" & vbCrLf
            s = s & "  m.Model=case when a.AssemblyType=0 OR a.AssemblyType=2 then a.Model else '' end" & vbCrLf
            s = s & "left outer join standardCostcodes cc on m.divisionid=cc.divisionid and m.BillingCode=cc.costcode" & vbCrLf
            s = s & "left outer join standardcategories cat on m.divisionid=cat.divisionid and m.BillingCategory=cat.category" & vbCrLf
            s = s & "left outer join taxgroups t on m.divisionid=t.divisionid and j.artaxgroup=t.taxgroup" & vbCrLf
            s = s & "where isnull(a.AREstimatePostingBatch,0)=0" & vbCrLf
            s = s & " and a.Divisionid = " & HFApp.DivisionID & vbCrLf
            s = s & "order by c.area,j.communityphase,j.job_no" & vbCrLf
            lblListDescription.Caption = "Select the estimates you would like to post."
            gData.AllowSelection = True
    End Select
    Set rs = HFApp.SqlExec(s, dbHomefront, ra)
    Set gData.DataSource = rs
    'hide EstAssemblyID
    With gData
        .ColHidden(0) = True
        'hide validationErrors if reviewing a batch
        If chkReview.value = True Then .ColHidden(.Cols - 1) = True
        .Refresh
        On Error Resume Next
        .Cell(flexcpForeColor, 1, .Cols - 1, .Rows - 1, .Cols - 1) = vbRed
        Call .Select(1, 0)
    End With
    cmdNav(2).Enabled = SelectionValid()
    Screen.MousePointer = vbDefault

Exit Sub
eh: Call errHandler(SRCFILE & "LoadData")
End Sub

Private Function SelectionValid() As Boolean
    Dim i As Long
    
    'check that validationerrors is blank in all selected rows
    With gData
    For i = 0 To .SelectedRows - 1
        If .TextMatrix(.SelectedRow(i), .Cols - 1) <> "" Then
            SelectionValid = False
            Exit Function
        End If
    Next
    End With
    
    SelectionValid = gData.SelectedRows > 0
    
End Function
    
Private Sub CreateBatch()
On Error GoTo eh

    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    cmdNav(0).Enabled = False
    cmdNav(1).Enabled = False
    
    'create new batch
    Call HFApp.SqlExec("INSERT INTO Batches(BatchType,UStmp,TStmp,DivisionID) VALUES('AR Estimate Posting'," & DbQuote(Str, HFApp.LoginID) & ",GETDATE()," & HFApp.DivisionID & ")")
    mBatch = HFApp.SqlIdentity("Batches")
    
    'write batch number to selected items
    With gData
    s = ""
    WhereClause = ""
    For i = 0 To .SelectedRows - 1
        s = s & "," & DbQuote(Num, .TextMatrix(.SelectedRow(i), 0))
    Next
    If s = "" Then Exit Sub
    WhereClause = " estassemblyid in(" & Mid(s, 2) & ")"
    End With
    
    s = ""
    s = s & "UPDATE EstimateAssemblies SET AREstimatePostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
    s = s & "WHERE ISNULL(AREstimatePostingBatch,0)=0 and " & WhereClause
    Call HFApp.SqlExec(s)
    Call SendBatch
    
Exit Sub
eh: Call errHandler(SRCFILE & "CreateBatch")
End Sub


Private Sub SendBatch()
    
    'import batch
    Select Case HFApp.Options(AccountingSystem)
        Case asQuickBooks:    Call SendBatchToQB
    End Select

End Sub



Private Sub SendBatchToQB()
    Dim s As String
    Dim i As Integer
    Dim rs As Recordset
    Dim lastJob As String
    Dim ThisJob As String
    Dim HFJobNumber As String
    Dim IsUSVersion As Boolean
    
    
    IsUSVersion = HFApp.Options.ValueByName("AccountingVersion") <> "CA"
    
    s = "select j.ExternalJobID,j.Job_no from estimateassemblies a join tbljobs j on a.job=j.job_no where isnull(j.ExternalJobID,'')='' and " & WhereClause
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        Call HFApp.WriteJobToAccounting("" & rs("job_no"))
        rs.MoveNext
    Wend
        
    
    s = ""
    s = s & "select " & vbCrLf
    s = s & "  cc.ExternalID JCCostCode" & vbCrLf
    s = s & " ,cat.ExternalID JCCategory" & vbCrLf
    s = s & " ,a.SalesQty*m.BillingFactor Qty" & vbCrLf
    s = s & " ,a.SalesRate Rate" & vbCrLf
    s = s & " ,j.ARTaxGroup TaxGroup" & vbCrLf
    s = s & " ,a.HFDescription ItemDesc" & vbCrLf
    s = s & " ,j.ExternalJobID" & vbCrLf
    s = s & " ,j.job_no" & vbCrLf
    s = s & " ,j.notes" & vbCrLf
    s = s & "from estimateassemblies a" & vbCrLf
    s = s & "join tbljobs j on a.job=j.job_no and a.divisionid=j.divisionid" & vbCrLf
    s = s & "join tbllocality c on c.area=j.community" & vbCrLf
    s = s & "join tblDBAssemblyMaster m ON" & vbCrLf
    s = s & "  m.DivisionID = a.DivisionID and" & vbCrLf
    s = s & "  m.Community=dbo.Purch_GetAssemblyCommunity(a.Assembly,case when(a.AssemblyType=3 OR a.AssemblyType=4) then '' else a.Model end,a.OptionID,j.Community,a.DivisionID) and" & vbCrLf
    s = s & "  m.Assembly=a.Assembly and" & vbCrLf
    s = s & "  m.OptionID=a.OptionID and" & vbCrLf
    s = s & "  m.Model=case when a.AssemblyType=0 OR a.AssemblyType=2 then a.Model else '' end" & vbCrLf
    s = s & "left outer join standardCostcodes cc on m.divisionid=cc.divisionid and m.BillingCode=cc.costcode" & vbCrLf
    s = s & "left outer join standardcategories cat on m.divisionid=cat.divisionid and m.BillingCategory=cat.category" & vbCrLf
    s = s & "where isnull(a.AREstimatePostingBatch,0)=" & DbQuote(Num, mBatch) & vbCrLf
    s = s & "ORDER BY 6,5" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    s = HFApp.XmlQBStart()
    While Not rs.EOF
        ThisJob = "" & rs("ExternalJobID")
        
        HFJobNumber = "" & rs("Job_no")
        If ThisJob <> lastJob Then
            If lastJob <> "" Then
                s = s & "</EstimateAdd></EstimateAddRq>" & vbCrLf
                s = s & HFApp.XmlQBEnd()
                s = HFApp.XmlQBSubmit(s)
                s = HFApp.XmlQBStart()
            End If
            i = i + 1
            lastJob = ThisJob
            s = s & "<EstimateAddRq requestID=""1"">" & vbCrLf
            s = s & "<EstimateAdd>" & vbCrLf
            s = s & "<CustomerRef>" & vbCrLf
            s = s & HFApp.XmlQBAdd(c, 41, "ListID", "" & ThisJob)
            s = s & "</CustomerRef>" & vbCrLf
            s = s & HFApp.XmlQBAdd(d, 41, "TxnDate", Now())
            
            s = s & HFApp.XmlQBAdd(c, 4095, "Memo", "" & rs("notes"))
            
        End If
        s = s & "<EstimateLineAdd>" & vbCrLf
        s = s & "<ItemRef>" & HFApp.XmlQBAdd(c, 41, "ListID", "" & rs("JCCostCode")) & "</ItemRef>" & vbCrLf
        s = s & HFApp.XmlQBAdd(c, 4095, "Desc", "" & rs("ItemDesc"))
        s = s & HFApp.XmlQBAdd(n, 8.5, "Quantity", "" & rs("Qty"))
        s = s & HFApp.XmlQBAdd(n, 8.2, "Rate", "" & rs("Rate"))
        s = s & "<ClassRef>" & HFApp.XmlQBAdd(c, 41, "ListID", "" & rs("JCCategory")) & "</ClassRef>" & vbCrLf
        If IsUSVersion = False And "" & rs("TaxGroup") <> "" Then
            s = s & "<SalesTaxCodeRef>" & HFApp.XmlQBAdd(c, 3, "FullName", "" & rs("TaxGroup")) & "</SalesTaxCodeRef>" & vbCrLf
        End If
        s = s & "</EstimateLineAdd>" & vbCrLf
        
        rs.MoveNext
    Wend
    If HFJobNumber <> "" Then
    s = s & "</EstimateAdd></EstimateAddRq>" & vbCrLf
    s = s & HFApp.XmlQBEnd()
    s = HFApp.XmlQBSubmit(s)
    End If

    
Exit Sub
eh: Call errHandler(SRCFILE & "SendBatchToQB", s)
End Sub



    

Private Sub Form_Load()
    Dim s As String

    Call IniGetForm(Me)
On Error Resume Next
    s = Choose(HFApp.Options(AccountingSystem), "Timberline", "MasterBuilder", "Quickbooks", "Sage50", "MYOB", "Peachtree", "Xero", "Spectrum")
    Image2(0).Picture = FMain.LargeIcons.ListImages(s).Picture
    imgError.Picture = FMain.LargeIcons.ListImages(s).Picture
    Set WizHead.Picture = FMain.LargeIcons.ListImages(s).Picture
On Error GoTo 0

  

    CurrentFrame = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Dim i As Long
    Const margin = 60
    For i = 0 To WizFrame.UBound
        WizFrame(i).BorderStyle = 0
        WizFrame(i).Move 0, WizHead.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead.Height - WizFoot.Height
    Next
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
    gData.Move gData.Left, gData.Top, WizFrame(0).Width - gData.Left - 2 * margin, WizFrame(0).Height - gData.Top - 2 * margin

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3
            If chkReview.value Then
                Unload Me
            Else
                Call CreateBatch
                Unload Me
            End If

    End Select
End Sub

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
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
    Select Case RHS
        Case 0:
        Case 1: Call LoadData
        Case 2: Call LoadTasks
    End Select
End Property


Private Sub gData_SelChange()
    cmdNav(2).Enabled = SelectionValid()
End Sub


Private Sub lblBatchTask_Click(Index As Integer)
    Dim s As String
    
    Select Case Index
        Case 0 'open summary report
            s = HFApp.SystemFolder & "System\Reports\Estimating\ExportAREstimateBatch.rpt"
            
            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            Call c.ParameterValue("Batch", mBatch)
            On Error GoTo 0
            Call c.PrintPreview("Print Preview")
        
    End Select

End Sub
