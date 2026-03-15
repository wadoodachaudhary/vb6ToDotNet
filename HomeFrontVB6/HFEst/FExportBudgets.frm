VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FExportBudgets 
   Caption         =   "Budget Posting Wizard"
   ClientHeight    =   4785
   ClientLeft      =   7005
   ClientTop       =   1410
   ClientWidth     =   6330
   ControlBox      =   0   'False
   Icon            =   "FExportBudgets.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6330
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3270
      Index           =   2
      Left            =   0
      TabIndex        =   13
      Top             =   870
      Width           =   6315
      Begin VB.Frame frmPost 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2835
         Left            =   1140
         TabIndex        =   20
         Top             =   420
         Width           =   4575
         Begin VB.CheckBox chkShowProcessingJournal 
            Caption         =   "Show processing journal when finished."
            Height          =   285
            Left            =   420
            TabIndex        =   22
            Top             =   795
            Width           =   4455
         End
         Begin VB.CheckBox chkShowBatchReport 
            Caption         =   "Show posting summary report when finished."
            Height          =   285
            Left            =   420
            TabIndex        =   21
            Top             =   540
            Width           =   4455
         End
      End
      Begin VB.Frame frmReview 
         BorderStyle     =   0  'None
         Height          =   2835
         Left            =   1140
         TabIndex        =   15
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
            TabIndex        =   19
            Top             =   210
            Width           =   2325
         End
         Begin VB.Label lblBatchTask 
            AutoSize        =   -1  'True
            Caption         =   "Open the processing journal"
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
            Index           =   3
            Left            =   420
            TabIndex        =   18
            Top             =   1380
            Width           =   1980
         End
         Begin VB.Label lblBatchTask 
            AutoSize        =   -1  'True
            Caption         =   "Review the rejected records"
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
            Index           =   2
            Left            =   420
            TabIndex        =   17
            Top             =   1140
            Width           =   1995
         End
         Begin VB.Label lblBatchTask 
            AutoSize        =   -1  'True
            Caption         =   "Review the posting file"
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
            Left            =   420
            TabIndex        =   16
            Top             =   900
            Width           =   1605
         End
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         Caption         =   "The posting wizard is ready to proceed."
         Height          =   195
         Left            =   1140
         TabIndex        =   14
         Top             =   120
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FExportBudgets.frx":000C
         Top             =   240
         Width           =   480
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
      Caption         =   "Post Budgets to Accounting"
      Description     =   "The posting wizard will send budgets to accounting"
      Icon            =   "FExportBudgets.frx":08D6
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
      Left            =   10350
      TabIndex        =   8
      Top             =   5520
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
         Caption         =   "Post new budgets"
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
         Picture         =   "FExportBudgets.frx":11B0
         Top             =   240
         Width           =   480
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
         FormatString    =   $"FExportBudgets.frx":1A7A
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
         Picture         =   "FExportBudgets.frx":1AB6
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
End
Attribute VB_Name = "FExportBudgets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FExportBudgets::"

Private mBatch As Long
Private mCOBatch As Long
Private WhereClause As String


Private Sub LoadTasks()
On Error GoTo eh
    Select Case True
        Case chkReview.value = True
            frmReview.Visible = True
            frmPost.Visible = False
            mBatch = gData.ValueMatrix(gData.Row, 0)
        
        Case Else
            cmdNav(3).Enabled = False
            frmReview.Visible = False
            frmPost.Visible = False
            chkShowProcessingJournal.Visible = HFApp.Options(AccountingSystem) = asTimberline
    
    
            If HFApp.Options(AccountingSystem) = asMasterBuilder Then
                lblMessage.Caption = ""
                If MbApiIsRunning Then
                    lblMessage.Caption = "The posting wizard is ready to proceed."
                    frmPost.Visible = True
                    cmdNav(3).Enabled = True
                Else
                    lblMessage.Caption = "The posting wizard is not able to process these budgets." & vbCrLf & vbCrLf & "The Master Builder API is not accepting requests."
                    frmPost.Visible = False
                    cmdNav(3).Enabled = False
                End If
            Else
                frmPost.Visible = True
                cmdNav(3).Enabled = True
            End If
    
    End Select
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadData")
End Sub


Private Sub LoadData()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As ADODB.Recordset
    Dim ra As Long
    Screen.MousePointer = vbHourglass
    
    
    Select Case True
        Case chkReview.value = True
            s = ""
            s = s & "SELECT Batch" & vbCrLf
            s = s & "      ,TStmp Posted" & vbCrLf
            s = s & "      ,UStmp PostedBy" & vbCrLf
            s = s & "  FROM Batches" & vbCrLf
            s = s & " WHERE (DivisionID = " & HFApp.DivisionID & " or isnull(DivisionID,0)=0) and BatchType='Budget Posting'" & vbCrLf
            s = s & "ORDER BY Batch DESC" & vbCrLf
            lblListDescription.Caption = "Select the batch you would like to review."
            gData.AllowSelection = False
        Case Else
           s = ""
            s = s & "SELECT i.Job,j.Description ""Job Description""" & vbCrLf
            s = s & "  FROM EstimateItems i Join tbljobs j on i.Job = j.Job_No and i.DivisionID = j.DivisionID" & vbCrLf
            s = s & " WHERE BudgetGenerated = 1" & vbCrLf
            s = s & "   AND BudgetPostingBatch=0" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and i.Divisionid = " & HFApp.DivisionID & vbCrLf
            End If
            s = s & "UNION" & vbCrLf
            s = s & "SELECT cb.Job,j.Description ""Job Description""" & vbCrLf
            s = s & "  FROM cancelledbudgets cb join tbljobs j on cb.Job = j.Job_No and cb.DivisionID = j.DivisionID" & vbCrLf
            s = s & " WHERE ISNULL(cb.PostBatch,0)=0" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and cb.Divisionid = " & HFApp.DivisionID & vbCrLf
            End If
             lblListDescription.Caption = "Select the budgets you would like to post."
            gData.AllowSelection = True
    End Select
    Set rs = HFApp.SqlExec(s, dbHomefront, ra)
    Set gData.DataSource = rs
    gData.Refresh
    On Error Resume Next
    Call gData.Select(1, 0)
    cmdNav(2).Enabled = gData.SelectedRows > 0
    Screen.MousePointer = vbDefault
Exit Sub
eh: Call errHandler(SRCFILE & "LoadData")
End Sub


Private Sub CreateBatch()
On Error GoTo eh

    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    cmdNav(0).Enabled = False
    cmdNav(1).Enabled = False
    
    'create new batch
    s = "INSERT INTO Batches(BatchType,UStmp,TStmp,DivisionID) VALUES('Budget Posting'," & DbQuote(Str, HFApp.LoginID) & ",GETDATE()," & HFApp.DivisionID & ")"
    Call HFApp.SqlExec(s)
    mBatch = HFApp.SqlIdentity("Batches")
    
    'write batch number to selected items
    With gData
    s = ""
    WhereClause = ""
    For i = 0 To .SelectedRows - 1
        s = s & " OR (Job_No=" & DbQuote(Str, .TextMatrix(.SelectedRow(i), 0)) & ")"
        WhereClause = WhereClause & " OR (e.Job_No=" & DbQuote(Str, .TextMatrix(.SelectedRow(i), 0)) & ")"
    Next
    s = Mid(s, 5)
    WhereClause = Mid(WhereClause, 5)
    End With
    Select Case HFApp.Options(AccountingSystem)
        Case asTimberline, asMasterBuilder, asIntacct
            Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingBatch=" & DbQuote(Num, mBatch) & " WHERE BudgetGenerated=1 AND ISNULL(BudgetPostingBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE " & s & ")")
            Call HFApp.SqlExec("UPDATE CancelledBudgets SET PostBatch=" & DbQuote(Num, mBatch) & " WHERE ISNULL(PostBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM CancelledBudgets WHERE " & Replace(s, "job_no=", "job=") & ")")
            Call SendBatch
        Case asQuickBooks, asQuickbooksOnline
            Call SendBatch
    End Select
    
Exit Sub
eh: Call errHandler(SRCFILE & "CreateBatch")
End Sub


Private Sub SendBatch()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    
    Select Case HFApp.Options(AccountingSystem)
    Case asMasterBuilder, asIntacct
        'ensure jobs exist
        Set rs = HFApp.SqlExec("select distinct job from estimateitems  where BudgetPostingBatch=" & DbQuote(Num, mBatch))
        While Not rs.EOF
            Call HFApp.WriteJobToAccounting("" & rs("job"))
            rs.MoveNext
        Wend
    End Select
    'import batch
    Select Case HFApp.Options(AccountingSystem)
        Case asIntacct:           Call SendEstimatesToIntacct(mBatch)
        Case asTimberline:        Call SendBatchToSage300
        Case asMasterBuilder:     Call SendBatchToSage100
        Case asQuickBooks:        Call SendBatchToQB
        Case asQuickbooksOnline:  Call SendBatchToQBOnline
    End Select


    'show summary file
    If chkShowBatchReport.value = vbChecked Then
        s = HFApp.SystemFolder & "System\Reports\Estimating\ExportBudgetBatch.rpt"
    
        Dim c As New ZybUtil.Crystal
        Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
        On Error Resume Next
        Call c.ParameterValue("DivisionID", HFApp.DivisionID)
        
        If mCOBatch <> 0 Then Call c.ParameterValue("Batch", mCOBatch)
        Call c.ParameterValue("Batch", mBatch)
        
        Call c.ParameterValue("PostSummaries", HFApp.Options(PostSummarizedPOs))
        On Error GoTo eh
        Call c.PrintPreview("Print Preview")
    
    End If
Exit Sub
eh: Call errHandler(SRCFILE & "SendBatch")
End Sub

Private Sub SendBatchToSage100()
On Error GoTo eh

    Dim budgetcreated As Boolean
    Dim co As Long
    
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    Dim lastJob As String
    Dim ThisJob As String
    Dim ThisCO As String
    Dim LastCo As String
    Dim Mode As String
    Dim COID  As Long
    
'set all posting dates to null, then post each job/co individually. set dates as rows are posted
'when job/co fails an error is thrown and processing is stopped. the rows's with a posting date are marked
'as posted, the rest are removed from the batch.
Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingDate=NULL WHERE BudgetPostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
    
    If HFApp.Options(PostSummarizedBudgets) Then
        s = ""
        s = s & "SELECT c.Job" & vbCrLf
        s = s & "      ,c.ChangeOrder" & vbCrLf
        s = s & "      ,c.JCExtra" & vbCrLf
        s = s & "      ,sc.Externalid CostCode" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=1 then c.Pretax+isnull(c.Tax,0) else 0 end) ct1" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=2 then c.Pretax+isnull(c.Tax,0) else 0 end) ct2" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=3 then c.Pretax+isnull(c.Tax,0) else 0 end) ct3" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=4 then c.Pretax+isnull(c.Tax,0) else 0 end) ct4" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=5 then c.Pretax+isnull(c.Tax,0) else 0 end) ct5" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=6 then c.Pretax+isnull(c.Tax,0) else 0 end) ct6" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=7 then c.Pretax+isnull(c.Tax,0) else 0 end) ct7" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=8 then c.Pretax+isnull(c.Tax,0) else 0 end) ct8" & vbCrLf
        s = s & "      ,sum(case when c.JCCategory=9 then c.Pretax+isnull(c.Tax,0) else 0 end) ct9" & vbCrLf
        s = s & "      ,sum(c.Qty) Qty" & vbCrLf
        s = s & "      ,'Summarized' Description" & vbCrLf
        s = s & "  FROM CancelledBudgets c" & vbCrLf
        s = s & "       join standardcostcodes sc on c.divisionid=sc.divisionid and c.jccostcode=sc.costcode" & vbCrLf
        s = s & " WHERE PostBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "GROUP BY c.ChangeOrder,c.Job,c.JCExtra,sc.externalid" & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "SELECT Job_No" & vbCrLf
        s = s & "      ,ChangeOrder" & vbCrLf
        s = s & "      ,JCExtra" & vbCrLf
        s = s & "      ,ExternalCostCode JCCostCode" & vbCrLf
        s = s & "      ,sum(case when JCCategory=1 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct1" & vbCrLf
        s = s & "      ,sum(case when JCCategory=2 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct2" & vbCrLf
        s = s & "      ,sum(case when JCCategory=3 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct3" & vbCrLf
        s = s & "      ,sum(case when JCCategory=4 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct4" & vbCrLf
        s = s & "      ,sum(case when JCCategory=5 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct5" & vbCrLf
        s = s & "      ,sum(case when JCCategory=6 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct6" & vbCrLf
        s = s & "      ,sum(case when JCCategory=7 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct7" & vbCrLf
        s = s & "      ,sum(case when JCCategory=8 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct8" & vbCrLf
        s = s & "      ,sum(case when JCCategory=9 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end) ct9" & vbCrLf
        s = s & "      ,sum(BudgetQty) Qty" & vbCrLf
        s = s & "      ,'Summarized' Description" & vbCrLf
        s = s & "  FROM EstimatedItems" & vbCrLf
        s = s & " WHERE BudgetPostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "GROUP BY ChangeOrder,Job_No,JCExtra,ExternalCostCode" & vbCrLf
        s = s & "ORDER BY 1,2,3,4" & vbCrLf
    Else
        s = ""
        s = s & "SELECT c.Job" & vbCrLf
        s = s & "      ,c.ChangeOrder" & vbCrLf
        s = s & "      ,c.JCExtra" & vbCrLf
        s = s & "      ,sc.Externalid CostCode" & vbCrLf
        s = s & "      ,case when cc.externalid=1 then c.Pretax+isnull(c.Tax,0) else 0 end ct1" & vbCrLf
        s = s & "      ,case when cc.externalid=2 then c.Pretax+isnull(c.Tax,0) else 0 end ct2" & vbCrLf
        s = s & "      ,case when cc.externalid=3 then c.Pretax+isnull(c.Tax,0) else 0 end ct3" & vbCrLf
        s = s & "      ,case when cc.externalid=4 then c.Pretax+isnull(c.Tax,0) else 0 end ct4" & vbCrLf
        s = s & "      ,case when cc.externalid=5 then c.Pretax+isnull(c.Tax,0) else 0 end ct5" & vbCrLf
        s = s & "      ,case when cc.externalid=6 then c.Pretax+isnull(c.Tax,0) else 0 end ct6" & vbCrLf
        s = s & "      ,case when cc.externalid=7 then c.Pretax+isnull(c.Tax,0) else 0 end ct7" & vbCrLf
        s = s & "      ,case when cc.externalid=8 then c.Pretax+isnull(c.Tax,0) else 0 end ct8" & vbCrLf
        s = s & "      ,case when cc.externalid=9 then c.Pretax+isnull(c.Tax,0) else 0 end ct9" & vbCrLf
        s = s & "      ,c.Qty Qty" & vbCrLf
        s = s & "      ,isnull(nullif(i.Description,''),'?') Description" & vbCrLf
        s = s & "  FROM CancelledBudgets c" & vbCrLf
        s = s & "       left outer join EstimateItems i on c.estitemid=i.estitemid" & vbCrLf
        s = s & "       join standardcostcodes sc on c.divisionid=sc.divisionid and c.jccostcode=sc.costcode" & vbCrLf
        s = s & "       join standardcategories cc on c.divisionid=cc.divisionid and c.jccategory=cc.category" & vbCrLf
        s = s & " WHERE PostBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "SELECT Job_No" & vbCrLf
        s = s & "      ,ChangeOrder" & vbCrLf
        s = s & "      ,JCExtra" & vbCrLf
        s = s & "      ,ExternalCostCode JCCostCode" & vbCrLf
        s = s & "      ,case when externalcategory=1 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct1" & vbCrLf
        s = s & "      ,case when externalcategory=2 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct2" & vbCrLf
        s = s & "      ,case when externalcategory=3 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct3" & vbCrLf
        s = s & "      ,case when externalcategory=4 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct4" & vbCrLf
        s = s & "      ,case when externalcategory=5 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct5" & vbCrLf
        s = s & "      ,case when externalcategory=6 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct6" & vbCrLf
        s = s & "      ,case when externalcategory=7 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct7" & vbCrLf
        s = s & "      ,case when externalcategory=8 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct8" & vbCrLf
        s = s & "      ,case when externalcategory=9 then BudgetPretax+isnull(BudgetJCTax,0) else 0 end ct9" & vbCrLf
        s = s & "      ,BudgetQty" & vbCrLf
        s = s & "      ,isnull(nullif(ItemDesc,''),'?')" & vbCrLf
        s = s & "  FROM EstimatedItems" & vbCrLf
        s = s & " WHERE BudgetPostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "ORDER BY 1,2,3,4" & vbCrLf
    End If
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    
    lastJob = ""
    LastCo = ""
    While Not rs.EOF
        ThisJob = "" & rs("Job")
        ThisCO = "" & rs("ChangeOrder")
        If ThisJob <> lastJob Or ThisCO <> LastCo Then
            If LastCo <> "" Then
                'end co
                s = s & "</ChangeOrder" & Mode & "Rq>" & vbCrLf
                s = s & HFApp.XmlMBEnd
                Call HFApp.XmlMbSubmit(s, HFApp.Options(MasterBuilderPWD))
                Call HFApp.SqlExec("UPDATE EstimatedItems SET BudgetPostingDate=GETDATE() WHERE job_no=" & DbQuote(Str, lastJob) & " and changeorder=" & DbQuote(Str, LastCo) & " and BudgetPostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
            ElseIf lastJob <> "" Then
                'end job
                s = s & "</Budget" & Mode & "Rq>" & vbCrLf
                s = s & HFApp.XmlMBEnd
                Call HFApp.XmlMbSubmit(s, HFApp.Options(MasterBuilderPWD))
                Call HFApp.SqlExec("UPDATE EstimatedItems SET BudgetPostingDate=GETDATE() WHERE job_no=" & DbQuote(Str, lastJob) & " and changeorder='' and BudgetPostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
            End If
                
            If ThisCO = "" Then
                'start job
                Mode = "Add"
                On Error Resume Next
                Mode = HFApp.SqlExec("select 'Mod' from budget where recnum=" & DbQuote(Num, ThisJob), dbAccounting)(0)
                On Error GoTo eh
                
                s = HFApp.XmlMbStart(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
                s = s & "<Budget" & Mode & "Rq requestID=""1"">" & vbCrLf
                s = s & HFApp.XmlMBAdd(n, 10, "ObjectRef", ThisJob)
                lastJob = ThisJob
                LastCo = ThisCO
            Else
                'start co
                On Error Resume Next
                COID = 0
                COID = HFApp.SqlExec("select max(recnum) from prmchg where jobnum=" & DbQuote(Num, ThisJob) & " and chgnum=" & DbQuote(Str, ThisCO), dbAccounting)(0)
                Mode = IIf(COID = 0, "Add", "Mod")
                On Error GoTo eh
                
                s = HFApp.XmlMbStart(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
                s = s & "<ChangeOrder" & Mode & "Rq requestID=""1"">" & vbCrLf
                If COID <> 0 Then
                    s = s & HFApp.XmlMBAdd(n, 10, "ObjectRef", COID)
                End If
                s = s & HFApp.XmlMBAdd(c, 8, "ChangeOrderNumber", "" & ThisCO)
                s = s & HFApp.XmlMBAdd(d, 0, "ChangeOrderDate", Now)
                If COID = 0 Then
                    s = s & HFApp.XmlMBAdd(n, 10, "JobRef", ThisJob)
                End If
                s = s & HFApp.XmlMBAdd(c, 50, "Desc", "Change order")
                s = s & HFApp.XmlMBAdd(d, 0, "SubmittedDate", Now)
                s = s & HFApp.XmlMBAdd(d, 0, "ApprovedDate", Now)
                s = s & "<PrimeChangeStatus>1</PrimeChangeStatus>" & vbCrLf
                lastJob = ThisJob
                LastCo = ThisCO
            End If
        End If
        
        If ThisCO = "" Then
            'add job line
            s = s & "<BudgetLineAdd>" & vbCrLf
            s = s & "<ObjectRef>" & HFApp.XmlMBAdd(n, 10, "PhaseID", Val("" & rs("JCExtra"))) & "</ObjectRef>" & vbCrLf
            s = s & HFApp.XmlMBAdd(n, 15.3, "CostCodeRef", "" & rs("CostCode"))
            s = s & HFApp.XmlMBAdd(n, 12.2, "MaterialAmount", "" & rs("ct1"))
            
            If HFApp.Options.ValueByName("MasterBuilderEdition") = "CA" Then
                s = s & HFApp.XmlMBAdd(n, 12.2, "LabourAmount", "" & rs("ct2"))
            Else
                s = s & HFApp.XmlMBAdd(n, 12.2, "LaborAmount", "" & rs("ct2"))
            End If
            s = s & HFApp.XmlMBAdd(n, 12.2, "EquipmentAmount", "" & rs("ct3"))
            s = s & HFApp.XmlMBAdd(n, 12.2, "SubcontractAmount", "" & rs("ct4"))
            s = s & HFApp.XmlMBAdd(n, 12.2, "OtherAmount", "" & rs("ct5"))
            s = s & HFApp.XmlMBAdd(n, 12.2, "UserDefined6Amount", "" & rs("ct6"))
            s = s & HFApp.XmlMBAdd(n, 12.2, "UserDefined7Amount", "" & rs("ct7"))
            s = s & HFApp.XmlMBAdd(n, 12.2, "UserDefined8Amount", "" & rs("ct8"))
            s = s & HFApp.XmlMBAdd(n, 12.2, "UserDefined9Amount", "" & rs("ct9"))
            s = s & HFApp.XmlMBAdd(n, 12.4, "Quantity", "" & rs("Qty"))
            s = s & HFApp.XmlMBAdd(m, 250, "Memo", HFApp.LoginID & " -- " & format(Now, "Mmm d/yy") & " -- " & rs("Description"))
            s = s & "</BudgetLineAdd>" & vbCrLf
        Else
            'add co line
            For i = 1 To 9
            If Val("" & rs("ct" & i)) <> 0 Then
                s = s & "<BudgetSubChangeDetailAdd>" & vbCrLf
                s = s & HFApp.XmlMBAdd(c, 50, "Desc", "" & rs("Description"))
                s = s & HFApp.XmlMBAdd(n, 12.2, "BudgetChangeAmount", "" & rs("ct" & i))
                s = s & HFApp.XmlMBAdd(n, 15.3, "CostCodeRef", "" & rs("CostCode"))
                s = s & HFApp.XmlMBAdd(n, 2, "CostTypeRef", i)
                s = s & "</BudgetSubChangeDetailAdd>" & vbCrLf
            End If
            Next
        End If
        rs.MoveNext
    Wend
    If LastCo <> "" Then
        'end co
        s = s & "</ChangeOrder" & Mode & "Rq>" & vbCrLf
        s = s & HFApp.XmlMBEnd
        Call HFApp.XmlMbSubmit(s, HFApp.Options(MasterBuilderPWD))
        Call HFApp.SqlExec("UPDATE EstimatedItems SET BudgetPostingDate=GETDATE() WHERE job_no=" & DbQuote(Str, lastJob) & " and changeorder=" & DbQuote(Str, LastCo) & " and BudgetPostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
    ElseIf lastJob <> "" Then
        'end job
        s = s & "</Budget" & Mode & "Rq>" & vbCrLf
        s = s & HFApp.XmlMBEnd
        Call HFApp.XmlMbSubmit(s, HFApp.Options(MasterBuilderPWD))
        Call HFApp.SqlExec("UPDATE EstimatedItems SET BudgetPostingDate=GETDATE() WHERE job_no=" & DbQuote(Str, lastJob) & " and changeorder='' and BudgetPostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
    End If
    
    
    
    
Exit Sub
eh: Call errHandler(SRCFILE & "SendBatchToSage100", s)
    Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingBatch=0 WHERE BudgetPostingDate is null and BudgetPostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
End Sub


Private Sub SendBatchToQBOnline()
    Dim s As String
    Dim i As Integer
    Dim rs As Recordset
    Dim lastJob As String
    Dim ThisJob As String
    Dim HFJobNumber As String
    Dim IsUSVersion As Boolean
    Dim Response As String
    
    IsUSVersion = HFApp.Options.ValueByName("AccountingVersion") <> "CA"
    
    s = "select e.ExternalJobID,e.Job_no from tbljobs e where e.DivisionID = " & HFApp.DivisionID & " and isnull(e.ExternalJobID,'')='' and (" & WhereClause & ")"
    Set rs = HFApp.SqlExec(s)
    
    While Not rs.EOF
    If "" & rs("ExternalJobID") = "" Then
        'Need to modify the WriteJobToAccounting code
        Call HFApp.WriteJobToAccounting("" & rs("job_no"))
    End If
        rs.MoveNext
    Wend
        
    s = ""
    s = s & "SELECT cc.ExternalID JCCostCode,cat.ExternalID JCCategory,sum(c.Pretax)+sum(c.Tax) Amount,c.taxgroup budgettaxgroup,cc.Description ItemDesc,e.ExternalJobID,e.job_no" & vbCrLf
    s = s & "  FROM CancelledBudgets c " & vbCrLf
    s = s & "  join tbljobs e on(c.job=e.job_no)" & vbCrLf
    s = s & "  join StandardCostCodes cc on c.jccostcode=cc.costcode" & vbCrLf
    s = s & "  join StandardCategories cat on c.jccategory=cat.category" & vbCrLf
    s = s & " WHERE (" & WhereClause & ") and e.DivisionID = " & HFApp.DivisionID & " and c.PostBatch=0 and c.JCCostCode<>''" & vbCrLf
    s = s & "GROUP BY cc.externalid,cat.externalid,c.taxgroup,cc.Description,e.ExternalJobID,e.job_no" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT e.ExternalCostCode,e.ExternalCategory,sum(e.BudgetPretax)+sum(e.BudgetJCTax) Amount,e.budgettaxgroup,e.JCCostCodeDesc ItemDesc,j.ExternalJobID,e.job_no" & vbCrLf
    s = s & "  FROM EstimatedItems e" & vbCrLf
    s = s & "  join tbljobs j on(e.job_no=j.job_no)" & vbCrLf
    s = s & " WHERE (" & WhereClause & ") and e.DivisionID = " & HFApp.DivisionID & " and e.BudgetPostingBatch=0 and e.BudgetGenerated=1 and e.JCCostCode<>''" & vbCrLf
    s = s & "GROUP BY e.ExternalCostCode,e.ExternalCategory,e.budgettaxgroup,e.JCCostCodeDesc,j.ExternalJobID,e.job_no" & vbCrLf
    s = s & "ORDER BY 6,5" & vbCrLf
    Set rs = HFApp.SqlExec(s)
ThisJob = ""
lastJob = ""

    s = ""
    s = s & "<?xml version=""" & "1.0""" & " encoding=""" & "utf-8""" & "?>" & vbCrLf
    s = s & "<Estimate xmlns:xsi=""" & "http://www.w3.org/2001/XMLSchema-instance""" & " xmlns:xsd=""" & "http://www.w3.org/2001/XMLSchema""" & ">" & vbCrLf
    While Not rs.EOF
        ThisJob = "" & rs("ExternalJobID")
        HFJobNumber = "" & rs("Job_no")
        If ThisJob <> lastJob Then
            If lastJob <> "" Then
                s = s & "</Estimate>" & vbCrLf
                Response = SubmitQBOXml("PostEstimate", s)
                
                
                If Response <> "" Then
                    Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingBatch=" & DbQuote(Num, mBatch) & " WHERE BudgetGenerated=1 AND ISNULL(BudgetPostingBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE Job_No=" & DbQuote(Str, HFJobNumber) & ")")
                    Call HFApp.SqlExec("UPDATE CancelledBudgets SET PostBatch=" & DbQuote(Num, mBatch) & " WHERE ISNULL(PostBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE Job_no=" & DbQuote(Str, HFJobNumber) & ")")
                End If
                s = ""
                s = s & "<?xml version=""" & "1.0""" & " encoding=""" & "utf-8""" & "?>" & vbCrLf
                s = s & "<Estimate xmlns:xsi=""" & "http://www.w3.org/2001/XMLSchema-instance""" & " xmlns:xsd=""" & "http://www.w3.org/2001/XMLSchema""" & ">" & vbCrLf
            End If
            i = i + 1
            lastJob = ThisJob
            s = s & HFApp.XmlQBAdd(c, 41, "Customer", "" & ThisJob)
        End If
        s = s & "<Line>" & vbCrLf
        s = s & vbTab & HFApp.XmlQBAdd(c, 41, "JCCostCode", "" & rs("JCCostCode"))
        s = s & vbTab & HFApp.XmlQBAdd(c, 4095, "Description", CleanXML("" & rs("ItemDesc")))
        s = s & vbTab & HFApp.XmlQBAdd(n, 8.2, "UnitPrice", "" & rs("Amount"))
        s = s & vbTab & HFApp.XmlQBAdd(n, 8.2, "Amount", "" & rs("Amount"))
        s = s & vbTab & HFApp.XmlQBAdd(n, 8.2, "Qty", "1")
        
        s = s & vbTab & HFApp.XmlQBAdd(c, 41, "JCCategory", "" & rs("JCCategory"))
        If IsUSVersion = False And "" & rs("BudgetTaxGroup") <> "" Then
            s = s & vbTab & HFApp.XmlQBAdd(c, 3, "TaxGroup", "" & rs("BudgetTaxGroup"))
        End If
        s = s & "</Line>" & vbCrLf
        
        rs.MoveNext
    Wend
    If HFJobNumber <> "" Then
    s = s & "</Estimate>" & vbCrLf
    Response = SubmitQBOXml("PostEstimate", s)
  
         If Response <> "" Then
            Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingBatch=" & DbQuote(Num, mBatch) & " WHERE BudgetGenerated=1 AND ISNULL(BudgetPostingBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE Job_No=" & DbQuote(Str, HFJobNumber) & ")")
            Call HFApp.SqlExec("UPDATE CancelledBudgets SET PostBatch=" & DbQuote(Num, mBatch) & " WHERE ISNULL(PostBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE Job_no=" & DbQuote(Str, HFJobNumber) & ")")
         End If
    End If

    
Exit Sub
eh: Call errHandler(SRCFILE & "SendBatchToQBOnline", s)
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
    
    s = "select e.ExternalJobID,e.Job_no from tbljobs e where e.DivisionID = " & HFApp.DivisionID & " and isnull(e.ExternalJobID,'')='' and (" & WhereClause & ")"
    Set rs = HFApp.SqlExec(s)
    
    While Not rs.EOF
    If "" & rs("ExternalJobID") = "" Then
        Call HFApp.WriteJobToAccounting("" & rs("job_no"))
    End If
        rs.MoveNext
    Wend
        
    s = ""
    s = s & "SELECT cc.ExternalID JCCostCode,cat.ExternalID JCCategory,sum(c.Pretax)+sum(c.Tax) Amount,c.taxgroup budgettaxgroup,cc.Description ItemDesc,e.ExternalJobID,e.job_no" & vbCrLf
    s = s & "  FROM CancelledBudgets c " & vbCrLf
    s = s & "  join tbljobs e on(c.job=e.job_no)" & vbCrLf
    s = s & "  join StandardCostCodes cc on c.jccostcode=cc.costcode" & vbCrLf
    s = s & "  join StandardCategories cat on c.jccategory=cat.category" & vbCrLf
    s = s & " WHERE (" & WhereClause & ") and e.DivisionID = " & HFApp.DivisionID & " and c.PostBatch=0 and c.JCCostCode<>''" & vbCrLf
    s = s & "GROUP BY cc.externalid,cat.externalid,c.taxgroup,cc.Description,e.ExternalJobID,e.job_no" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT e.ExternalCostCode,e.ExternalCategory,sum(e.BudgetPretax)+sum(e.BudgetJCTax) Amount,e.budgettaxgroup,e.JCCostCodeDesc ItemDesc,j.ExternalJobID,e.job_no" & vbCrLf
    s = s & "  FROM EstimatedItems e" & vbCrLf
    s = s & "  join tbljobs j on(e.job_no=j.job_no)" & vbCrLf
    s = s & " WHERE (" & WhereClause & ") and e.DivisionID = " & HFApp.DivisionID & " and e.BudgetPostingBatch=0 and e.BudgetGenerated=1 and e.JCCostCode<>''" & vbCrLf
    s = s & "GROUP BY e.ExternalCostCode,e.ExternalCategory,e.budgettaxgroup,e.JCCostCodeDesc,j.ExternalJobID,e.job_no" & vbCrLf
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
                If s <> "" Then
                    Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingBatch=" & DbQuote(Num, mBatch) & " WHERE BudgetGenerated=1 AND ISNULL(BudgetPostingBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE Job_No=" & DbQuote(Str, HFJobNumber) & ")")
                    Call HFApp.SqlExec("UPDATE CancelledBudgets SET PostBatch=" & DbQuote(Num, mBatch) & " WHERE ISNULL(PostBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE Job_no=" & DbQuote(Str, HFJobNumber) & ")")
                End If
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
        End If
        s = s & "<EstimateLineAdd>" & vbCrLf
        s = s & "<ItemRef>" & HFApp.XmlQBAdd(c, 41, "ListID", "" & rs("JCCostCode")) & "</ItemRef>" & vbCrLf
        s = s & HFApp.XmlQBAdd(c, 4095, "Desc", "" & rs("ItemDesc"))
        s = s & HFApp.XmlQBAdd(n, 8.2, "Rate", "" & rs("Amount"))
        s = s & "<ClassRef>" & HFApp.XmlQBAdd(c, 41, "ListID", "" & rs("JCCategory")) & "</ClassRef>" & vbCrLf
        If IsUSVersion = False And "" & rs("BudgetTaxGroup") <> "" Then
            s = s & "<SalesTaxCodeRef>" & HFApp.XmlQBAdd(c, 3, "FullName", "" & rs("BudgetTaxGroup")) & "</SalesTaxCodeRef>" & vbCrLf
        End If
        s = s & "</EstimateLineAdd>" & vbCrLf
        
        rs.MoveNext
    Wend
    If HFJobNumber <> "" Then
    s = s & "</EstimateAdd></EstimateAddRq>" & vbCrLf
    s = s & HFApp.XmlQBEnd()
    s = HFApp.XmlQBSubmit(s)
  
    If s <> "" Then
       Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingBatch=" & DbQuote(Num, mBatch) & " WHERE BudgetGenerated=1 AND ISNULL(BudgetPostingBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE Job_No=" & DbQuote(Str, HFJobNumber) & ")")
       Call HFApp.SqlExec("UPDATE CancelledBudgets SET PostBatch=" & DbQuote(Num, mBatch) & " WHERE ISNULL(PostBatch,0)=0 AND EstItemID IN(SELECT EstItemID FROM EstimatedItems WHERE Job_no=" & DbQuote(Str, HFJobNumber) & ")")
    End If
         
    End If

    
Exit Sub
eh: Call errHandler(SRCFILE & "SendBatchToQB", s)
End Sub




Private Sub SendBatchToSage300()
    'everything is relative to root of timberline datafolder
    Const MacroFile = "HomeFrontMacros\Estimates.mac"
    Const JCEFile = "HomeFrontMacros\Estimates.jce"
    Const RejectFile = "HomeFrontMacros\EstimatesReject.jce"
    Const ImportPrintFile = "HomeFrontMacros\EstimatesImportPrint.prn"
    Const PostPrintFile = "HomeFrontMacros\EstimatesPostPrint.prn"
    
    Const COMacroFile = "HomeFrontMacros\EstimatesCO.mac"
    Const COJCEFile = "HomeFrontMacros\EstimatesCO.jce"
    Const CORejectFile = "HomeFrontMacros\EstimatesRejectCO.jce"
    Const COImportPrintFile = "HomeFrontMacros\EstimatesImportPrintCO.prn"
    Const COPostPrintFile = "HomeFrontMacros\EstimatesPostPrintCO.prn"
    
    
    Dim s As String
    
    
    
    'if a changeorder mac file exists the post changeorder budgets to separate file
    If FileExists(PathAppend(HFApp.Options(Timberline_Data_Path), COMacroFile)) Then
                          
        'create batch for changeorder items
        Call HFApp.SqlExec("INSERT INTO Batches(BatchType,UStmp,TStmp,DivisionID) VALUES('Budget Posting'," & DbQuote(Str, HFApp.LoginID) & ",GETDATE()," & HFApp.DivisionID & ")")
        mCOBatch = HFApp.SqlIdentity("Batches")
        
        'move changeorder items to new batch
        s = ""
        s = s & "update cancelledbudgets" & vbCrLf
        s = s & "set postbatch=" & DbQuote(Num, mCOBatch) & vbCrLf
        s = s & "where isnull(changeorder,'')<>'' " & vbCrLf
        s = s & "and postbatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & vbCrLf
        s = s & "update estimateitems" & vbCrLf
        s = s & "set budgetpostingbatch=" & DbQuote(Num, mCOBatch) & vbCrLf
        s = s & "from estimateitems i" & vbCrLf
        s = s & "join estimateassemblies a on i.estassemblyid=a.estassemblyid" & vbCrLf
        s = s & "where isnull(a.changeorder,'')<>''" & vbCrLf
        s = s & "and i.budgetpostingbatch=" & DbQuote(Num, mBatch) & vbCrLf
        Call HFApp.SqlExec(s)

        
        Call SendBatchToTL_GO(mBatch, _
                              PathAppend(HFApp.Options(Timberline_Data_Path), JCEFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), MacroFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), RejectFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), ImportPrintFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), PostPrintFile))
        
        Call SendBatchToTL_GO(mCOBatch, _
                              PathAppend(HFApp.Options(Timberline_Data_Path), COJCEFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), COMacroFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), CORejectFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), COImportPrintFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), COPostPrintFile))
    
    Else
        Call SendBatchToTL_GO(mBatch, _
                              PathAppend(HFApp.Options(Timberline_Data_Path), JCEFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), MacroFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), RejectFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), ImportPrintFile), _
                              PathAppend(HFApp.Options(Timberline_Data_Path), PostPrintFile))
    End If
    
End Sub



Private Sub SendBatchToTL_GO(Batch As Long, jce As String, mac As String, rej As String, Iprn As String, Pprn As String)
On Error GoTo eh
    Dim s As String
    Dim i As Integer
    
    Dim prevJob As String
    Dim prevExtra As String
    
    Dim rs As Recordset
'    Dim f As New FRptViewer
    Dim debugStr As String

    If HFApp.Options(PostSummarizedBudgets) Then
        s = ""
        s = s & "SELECT Job Job_No,'' JobDesc,JCExtra,null JCExtraDesc,JCCostCode,JCCategory,0 Qty,'' UOM,SUM(Pretax+Tax) Amount" & vbCrLf
        s = s & "  FROM CancelledBudgets" & vbCrLf
        s = s & " WHERE PostBatch=" & DbQuote(Num, Batch) & vbCrLf
        s = s & "GROUP BY Job,JCExtra,JCCostCode,JCCategory" & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "SELECT Job_No,JobDesc,JCExtra,left(AssemblyDescription,30) JCExtraDesc,JCCostCode,JCCategory,0 Qty,'' UOM,SUM(BudgetPretax+BudgetJCTax) Amount" & vbCrLf
        s = s & "  FROM EstimatedItems " & vbCrLf
        s = s & " WHERE BudgetPostingBatch=" & DbQuote(Num, Batch) & vbCrLf
        s = s & "GROUP BY Job_No,JobDesc,JCExtra,left(AssemblyDescription,30),JCCostCode,JCCategory" & vbCrLf
        s = s & "ORDER BY Job_No,JCExtra,JCCostCode,JCCategory" & vbCrLf
    Else
        s = ""
        s = s & "SELECT Job Job_No,'' JobDesc,JCExtra,null JCExtraDesc,JCCostCode,JCCategory,Qty,UOM,Pretax+Tax Amount" & vbCrLf
        s = s & "  FROM CancelledBudgets" & vbCrLf
        s = s & " WHERE PostBatch=" & DbQuote(Num, Batch) & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "SELECT Job_No,JobDesc,JCExtra,left(AssemblyDescription,30) JCExtraDesc,JCCostCode,JCCategory,BudgetQty Qty,OrderUOM UOM,BudgetPretax+BudgetJCTax Amount" & vbCrLf
        s = s & "  FROM EstimatedItems" & vbCrLf
        s = s & " WHERE BudgetPostingBatch=" & DbQuote(Num, Batch) & vbCrLf
        s = s & "ORDER BY Job_No,JCExtra,JCCostCode,JCCategory" & vbCrLf
    End If
    Set rs = HFApp.SqlExec(s)
    
    If rs.EOF Then Exit Sub
    
    Call CreatePath("export path", FilePath(jce))
    i = FreeFile
    debugStr = jce
    Open jce For Output As #i
    
    While Not rs.EOF
    
    
        If prevJob <> "" & rs("job_no") Then
            Print #i, "*," & Quote("" & rs("Job_No")) & "," & Quote("" & rs("JobDesc"))
        End If
        If prevExtra <> "" & rs("jcextra") Then
            If Trim("" & rs("JCExtra")) <> "" Then Print #i, "E," & Quote("" & rs("JCExtra")) & "," & Quote("" & rs("JCExtraDesc"))
        End If
        Print #i, "C," & Quote("" & rs("JCCostCode")) & ",," & Quote("" & rs("JCCategory")) & "," & format(Now, "mmddyyyy") & "," & rs("Qty") & "," & Quote("" & rs("UOM")) & "," & rs("Amount") & ",,,,," & Quote("" & rs("JCExtra"))
        
        prevJob = "" & rs("job_no")
        prevExtra = "" & rs("jcextra")
        rs.MoveNext
    Wend
    Close i
    
    
    'save datafile to batch table
    On Error Resume Next
    Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & Batch, "DataFile", jce)
    On Error GoTo eh
    
    debugStr = "FTSObject.Run()"
    Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), _
                       HFApp.Options(Timberline_UID), _
                       HFApp.Options(Timberline_PWD), _
                       mac, _
                       Iprn, _
                       "Processing Budgets...")
    DoEvents 'so fmain can paint over ftsobject image
    On Error Resume Next
    Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & Batch, "RejectFile", rej)
    Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & Batch, "ImportPrintFile", Iprn)
    Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & Batch, "PostingPrintFile", Pprn)

On Error GoTo ehNoRollBack
        
    'show processing journal
    If chkShowProcessingJournal.value = vbChecked Then
        debugStr = Iprn
        s = TempFile(FileExt(Iprn))
        Call FileCopy(Iprn, s)
        Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), HFApp.Options(Timberline_UID), HFApp.Options(Timberline_PWD), s)
    End If

    
Exit Sub
eh: Select Case Err.Number
        Case 70:   MsgBox Err.Description & vbCrLf & debugStr, vbInformation
        Case Else: Call errHandler(SRCFILE & "SendBatchToTL")
    
Stop: Resume
    End Select
    'clear batch if it didnt work.
    Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingBatch=0 WHERE BudgetPostingBatch=" & DbQuote(Num, Batch), dbHomefront)
    Call HFApp.SqlExec("UPDATE cancelledbudgets SET PostBatch=0 WHERE PostBatch=" & DbQuote(Num, Batch), dbHomefront)
    Call HFApp.SqlExec("delete Batches where batch=" & DbQuote(Num, Batch), dbHomefront)

ehNoRollBack:
    Select Case Err.Number
        Case 70:   MsgBox Err.Description & vbCrLf & debugStr, vbInformation
        Case Else: Call errHandler(SRCFILE & "SendBatchToTL")
    End Select

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

    chkShowBatchReport.value = IIf(HFApp.Options(ShowBudgetExportReport), vbChecked, vbUnchecked)
    chkShowProcessingJournal.value = IIf(HFApp.Options(ShowBudgetProcessingJournal), vbChecked, vbUnchecked)
    
    If HFApp.Options(AccountingSystem) <> asTimberline Then
        chkShowProcessingJournal.value = vbUnchecked
        Me.lblBatchTask(1).Visible = False
        Me.lblBatchTask(2).Visible = False
        Me.lblBatchTask(3).Visible = False
    End If

    CurrentFrame = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
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
    HFApp.Options.value(ShowBudgetExportReport) = chkShowBatchReport.value = vbChecked
    HFApp.Options.value(ShowBudgetProcessingJournal) = chkShowProcessingJournal.value = vbChecked
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
    cmdNav(2).Enabled = gData.SelectedRows > 0
End Sub

Private Sub lblBatchTask_Click(Index As Integer)
    Dim s As String
    'Dim f As New FRptViewer
    
    Select Case Index
        Case 0 'open summary report
            s = HFApp.SystemFolder & "System\Reports\Estimating\ExportBudgetBatch.rpt"
            'Call f.ShowReport(s, True, True, "Batch", mBatch, "PostSummaries", HFApp.Options(PostSummarizedBudgets))
            
            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            Call c.ParameterValue("Batch", mBatch)
            Call c.ParameterValue("PostSummaries", HFApp.Options(PostSummarizedBudgets))
            On Error GoTo 0
            Call c.PrintPreview("Print Preview")
        
        Case 1 'review posting file
            s = TempFile("txt")
            Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "DataFile")
            Call ShellFile(Me.hwnd, s)
            
        Case 2 'review rejected records
            s = TempFile("txt")
            Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "RejectFile")
            Call ShellFile(Me.hwnd, s)
        
        Case 3 'open import print file
            s = TempFile("prn")
            Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "ImportPrintFile")
            Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), HFApp.Options(Timberline_UID), HFApp.Options(Timberline_PWD), s)
        
'        Case 4 'repost rejects
'            'extract rejects to datafile location
'            Call CreatePath("export path", FilePath(PathAppend(HFApp.Options(Timberline_Data_Path), DataFile)))
'            Call DBGetFile(PathAppend(HFApp.Options(Timberline_Data_Path), DataFile), , HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "RejectFile")
'            'launch macro
'            Call TsObject(HFApp.Options(Timberline_Data_Path), _
'                          PathAppend(HFApp.Options(Timberline_Data_Path), MacroFile), _
'                          HFApp.Options(Timberline_UID), _
'                          HFApp.Options(Timberline_PWD))
    
    End Select
End Sub





