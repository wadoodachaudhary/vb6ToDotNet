VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FExportPOs 
   Caption         =   "Purchase Order Posting Wizard"
   ClientHeight    =   4965
   ClientLeft      =   9390
   ClientTop       =   2055
   ClientWidth     =   6465
   ControlBox      =   0   'False
   Icon            =   "FExportPOs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6465
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6465
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4380
      Width           =   6465
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
   Begin HFEst.WizHead WizHead 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   1588
      Caption         =   "Post Purchase Orders to Accounting"
      Description     =   "The posting wizard will send Purchase Orders to accounting"
      Icon            =   "FExportPOs.frx":000C
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   3360
      Index           =   1
      Left            =   7680
      TabIndex        =   0
      Top             =   1590
      Visible         =   0   'False
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   3120
         Left            =   1140
         TabIndex        =   12
         Top             =   240
         Width           =   5175
         _cx             =   1986929576
         _cy             =   1986925951
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
         ExtendLastCol   =   0   'False
         FormatString    =   $"FExportPOs.frx":08E6
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
         Picture         =   "FExportPOs.frx":0922
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblListDescription 
         Caption         =   "Select the purchase orders you would like to post."
         Height          =   255
         Left            =   1140
         TabIndex        =   1
         Top             =   0
         Width           =   3915
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3360
      Index           =   2
      Left            =   -30
      TabIndex        =   13
      Top             =   870
      Width           =   6315
      Begin VB.Frame frmReview 
         BorderStyle     =   0  'None
         Height          =   1845
         Left            =   1020
         TabIndex        =   15
         Top             =   420
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label lblBatchTask 
            AutoSize        =   -1  'True
            Caption         =   "Re-post batch"
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
            Index           =   4
            Left            =   420
            TabIndex        =   23
            Top             =   540
            Width           =   1005
         End
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
      Begin VB.Frame frmPost 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1845
         Left            =   1080
         TabIndex        =   20
         Top             =   930
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CheckBox chkShowProcessingJournal 
            Caption         =   "Show processing journal when finished."
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   795
            Width           =   4455
         End
         Begin VB.CheckBox chkShowBatchReport 
            Caption         =   "Show posting summary report when finished."
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   540
            Width           =   4455
         End
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         Caption         =   "The posting wizard is ready to proceed"
         Height          =   195
         Left            =   1140
         TabIndex        =   14
         Top             =   120
         Width           =   2730
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FExportPOs.frx":11EC
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3360
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   900
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
         Caption         =   "Post purchase orders"
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
         Picture         =   "FExportPOs.frx":1AB6
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "FExportPOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FExportPOs::"

'everything is relative to root of timberline datafolder
Const MacroFile = "HomeFrontMacros\Commitments.mac"
Const JCCFile = "HomeFrontMacros\Commitments.jcc"
Const JCEFile = "HomeFrontMacros\Commitments.jce"
Const RejectFile = "HomeFrontMacros\CommitmentsReject.jcc"
Const ImportPrintFile = "HomeFrontMacros\CommitmentsImportPrint.prn"
Const PostPrintFile = "HomeFrontMacros\CommitmentsPostPrint.prn"

Private mBatch As Long

Private Sub LoadTasks()
On Error GoTo eh
    
    If chkReview.value Then
        frmReview.Visible = True
        frmPost.Visible = False
        mBatch = gData.ValueMatrix(gData.Row, 0)
        lblBatchTask(4).Enabled = gData.TextMatrix(gData.Row, 3) = ""
    Else
        cmdNav(3).Enabled = False
        frmReview.Visible = False
        frmPost.Visible = False
        chkShowProcessingJournal.Visible = HFApp.Options(AccountingSystem) = asTimberline
    End If
    
   
    If HFApp.Options(AccountingSystem) = asMasterBuilder Then
        lblMessage.Caption = ""
        If MbApiIsRunning Then
            lblMessage.Caption = "The posting wizard is ready to proceed."
            frmPost.Visible = True
            cmdNav(3).Enabled = True
        Else
            lblMessage.Caption = "Unable to post. The Master Builder API is not accepting requests."
            frmPost.Visible = False
            cmdNav(3).Enabled = False
        End If
    Else
        frmPost.Visible = True
        cmdNav(3).Enabled = True
    End If
    
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadTasks")
End Sub


Private Sub LoadData()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim ra As Long
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Set cn = New ADODB.Connection
    cn.Open HFApp.ConnectionString(dbHomefront), HFApp.LoginDBUID, HFApp.LoginDBPWD
    
    Select Case True
        Case chkReview.value = True
            s = ""
            s = s & "SELECT Batch" & vbCrLf
            s = s & "      ,UStmp PostedBy" & vbCrLf
            s = s & "      ,TStmp Posted" & vbCrLf
            s = s & "      ,CASE when repostbatch<>0 then 'Reposted as batch ' + cast(repostbatch as varchar) else '' end" & vbCrLf
            s = s & "  FROM Batches" & vbCrLf
            s = s & " WHERE (DivisionID = " & HFApp.DivisionID & " or isnull(DivisionID,0)=0) and BatchType IN('Repost PO Batch','PO Posting')" & vbCrLf
            s = s & "ORDER BY Batch DESC" & vbCrLf
            lblListDescription.Caption = "Select the batch you would like to review."
            gData.AllowSelection = False
        Case Else
            s = ""
            s = s & "SELECT p.PONumber ""PO Number""" & vbCrLf
            s = s & "      ,'' CO" & vbCrLf
            s = s & "      ,p.Job + ISNULL(' - ' + j.Description,'') Job" & vbCrLf
            s = s & "      ,ISNULL(v.Vendor_Name,p.Vendor) Vendor" & vbCrLf
            s = s & "      ,p.POIndex ""Purchase Order""" & vbCrLf
            s = s & "      ,p.Description" & vbCrLf
            s = s & "  FROM POMaster p " & vbCrLf
            s = s & "       LEFT OUTER JOIN tblVendors v ON (v.Vendor_id=p.Vendor and v.DivisionID = p.DivisionID)" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblJobs j ON (j.Job_No=p.Job and j.DivisionID = p.DivisionID)" & vbCrLf
            s = s & " WHERE p.status<>'Pending' and Cancelled=0 AND PostingBatch=0" & vbCrLf
            s = s & "   and p.DivisionID =" & HFApp.DivisionID & vbCrLf
            s = s & "   and isnull(v.isTBD,0)=0" & vbCrLf
            s = s & "   and (isnull(p.ismpo,0)=0 or isnull(p.mporeceived,0)=1)" & vbCrLf
            s = s & "UNION ALL" & vbCrLf
            s = s & "SELECT c.PONumber" & vbCrLf
            s = s & "      ,c.ChangeOrder CO" & vbCrLf
            s = s & "      ,p.Job + ISNULL(' - ' + j.Description,'') Job" & vbCrLf
            s = s & "      ,ISNULL(v.Vendor_Name,p.Vendor) Vendor" & vbCrLf
            s = s & "      ,p.POIndex" & vbCrLf
            s = s & "      ,p.Description" & vbCrLf
            s = s & "  FROM POMaster p JOIN POChangeOrders c ON(p.PONumber=c.PONumber and p.DivisionID = c.DivisionID)" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblVendors v ON (v.Vendor_id=p.Vendor and v.DivisionID = p.DivisionID)" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblJobs j ON (j.Job_No=p.Job and j.DivisionID = p.DivisionID)" & vbCrLf
            s = s & " WHERE p.status<>'Pending' and c.PostingBatch=0" & vbCrLf
            s = s & "   and p.DivisionID =" & HFApp.DivisionID & vbCrLf
            s = s & "   and isnull(v.isTBD,0)=0" & vbCrLf
            s = s & "   and (isnull(p.ismpo,0)=0 or isnull(p.mporeceived,0)=1)" & vbCrLf
            s = s & "ORDER BY 1,2,3,4"
            lblListDescription.Caption = "Select the purchase orders you would like to post."
            gData.AllowSelection = True
    End Select
    Set rs = cn.Execute(s, ra)
    'Set rs = HFApp.SqlExec(s, dbHomeFront, ra)
    Set gData.DataSource = rs
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
    Dim POs As String
    Dim s As String
    Dim rs As Recordset
    
    cmdNav(0).Enabled = False
    cmdNav(1).Enabled = False
    
    'create new batch
    Call HFApp.SqlExec("INSERT INTO Batches(BatchType,UStmp,TStmp,DivisionID) VALUES('PO Posting'," & DbQuote(Str, HFApp.LoginID) & ",GETDATE()" & "," & HFApp.DivisionID & ")")
    mBatch = HFApp.SqlIdentity("Batches")
    
    'write batch number to selected items
    With gData
        'save POs
        POs = ""
        For i = 0 To .SelectedRows - 1
            'if co is blank then this is a po row
            If .TextMatrix(.SelectedRow(i), 1) = "" Then
                POs = POs & "," & DbQuote(Str, .TextMatrix(.SelectedRow(i), 0))
            End If
        Next
        If POs <> "" Then
            s = ""
            s = s & "update pomaster" & vbCrLf
            s = s & "set PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
            s = s & "from pomaster p" & vbCrLf
            s = s & "join tblvendors v on p.divisionid=v.divisionid and p.vendor=v.vendor_id" & vbCrLf
            s = s & "where p.DivisionID = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "and p.PONumber IN(" & Mid(POs, 2) & ")" & vbCrLf
            s = s & "and isnull(v.isTBD,0)=0" & vbCrLf
            s = s & "and (isnull(p.isMPO,0)=0 or isnull(p.MPOReceived,0)=1)" & vbCrLf
            Call HFApp.SqlExec(s)
        End If
        'save COs
        s = ""
        For i = 0 To .SelectedRows - 1
            If .TextMatrix(.SelectedRow(i), 1) <> "" Then
                s = s & "," & DbQuote(Str, .TextMatrix(.SelectedRow(i), 0) & Chr(1) & .TextMatrix(.SelectedRow(i), 1))
            End If
        Next
        If s <> "" Then
            s = "UPDATE POChangeOrders SET PostingDate=GETDATE(),PostingBatch=" & DbQuote(Num, mBatch) & " WHERE DivisionID = " & HFApp.DivisionID & " and PONumber+char(1)+ChangeOrder IN(" & Mid(s, 2) & ")"
            Call HFApp.SqlExec(s)
        End If
    End With

    Call SendBatch
    
Exit Sub
eh: Call errHandler(SRCFILE & "CreateBatch")
End Sub


Private Function SendBatch() As Boolean
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim b  As Boolean
    Dim i As Long
    
    Dim LastPO As String
    
    Screen.MousePointer = vbHourglass
    
    'ensure jobs exist if not using Timberline Accounting and not None
    If HFApp.Options(AccountingSystem) <> asTimberline And HFApp.Options(AccountingSystem) <> 0 Then
        Set rs = HFApp.SqlExec("select distinct p.job,j.ExternalJobID from pomaster p join tbljobs j on (p.Job=j.Job_No and p.DivisionID = j.DivisionID) where  p.PostingBatch=" & DbQuote(Num, mBatch))
        While Not rs.EOF
            If "" & rs("ExternalJobID") = "" Then
                Call HFApp.WriteJobToAccounting("" & rs("job"))
            End If
            rs.MoveNext
        Wend
    End If
    
    'set summarized flag on pomaster
    s = ""
    s = s & "update pomaster" & vbCrLf
    s = s & "set summarized=" & DbQuote(Bit, HFApp.Options(PostSummarizedPOs)) & vbCrLf
    s = s & "where PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
    Call HFApp.SqlExec(s)
    
    'set line item numbers and descriptions on poitems
    If HFApp.Options(PostSummarizedPOs) Then
        s = ""
        s = s & "update poitems" & vbCrLf
        s = s & "set linedescription=isnull(c.description,a.Description)" & vbCrLf
'        s = s & "   ,linenumber=(select LineNumber" & vbCrLf
'        s = s & "                from (select row_number() over (PARTITION BY ponumber order by ponumber,job,jcextra,jccostcode,jccategory,taxgroup) LineNumber,ponumber,job,jcextra,jccostcode,jccategory,taxgroup " & vbCrLf
'        s = s & "                      from poitems x where x.ponumber=a.ponumber" & vbCrLf
'        s = s & "                      group by ponumber,job,jcextra,jccostcode,jccategory,taxgroup) b " & vbCrLf
'        s = s & "                where a.ponumber=b.ponumber" & vbCrLf
'        s = s & "                and a.taxgroup=b.taxgroup" & vbCrLf
'        s = s & "                and a.job=b.job" & vbCrLf
'        s = s & "                and isnull(a.jcextra,'')=isnull(b.jcextra,'')" & vbCrLf
'        s = s & "                and a.jccostcode=b.jccostcode" & vbCrLf
'        s = s & "                and a.jccategory=b.jccategory)" & vbCrLf
        s = s & "   ,linenumber=(select LineNumber" & vbCrLf
        s = s & "                from (select row_number() over (PARTITION BY ponumber order by ponumber,job,jcextra,jccostcode,jccategory,isnull(taxgroup,'')) LineNumber,ponumber,job,jcextra,jccostcode,jccategory,isnull(taxgroup,'') TaxGroup" & vbCrLf
        s = s & "                      from poitems x where x.ponumber=a.ponumber" & vbCrLf
        s = s & "                      group by ponumber,job,jcextra,jccostcode,jccategory,isnull(taxgroup,'')) b " & vbCrLf
        s = s & "                where a.ponumber=b.ponumber" & vbCrLf
        s = s & "                and isnull(a.taxgroup,'')=isnull(b.taxgroup,'')" & vbCrLf
        s = s & "                and a.job=b.job" & vbCrLf
        s = s & "                and isnull(a.jcextra,'')=isnull(b.jcextra,'')" & vbCrLf
        s = s & "                and a.jccostcode=b.jccostcode" & vbCrLf
        s = s & "                and a.jccategory=b.jccategory)" & vbCrLf
        
        s = s & "from poitems a" & vbCrLf
        s = s & "left outer join standardcostcodes c on a.DivisionID = c.DivisionID and a.jccostcode=c.costcode" & vbCrLf
        s = s & "where a.DivisionID =" & HFApp.DivisionID & " and a.genBatch=" & DbQuote(Num, mBatch) & vbCrLf
        Call HFApp.SqlExec(s)
    Else
        s = ""
        s = s & "update poitems" & vbCrLf
        s = s & "set linedescription=a.description" & vbCrLf
        s = s & "   ,linenumber=(select LineNumber" & vbCrLf
        s = s & "                from (select row_number() over (PARTITION BY DivisionID,ponumber order by ponumber,ItemSeq) LineNumber,ponumber,ItemSeq" & vbCrLf
        s = s & "                      from poitems x where x.ponumber=a.ponumber and x.DivisionID=a.DivisionID" & vbCrLf
        s = s & "                      group by DivisionID,ponumber,ItemSeq) b " & vbCrLf
        s = s & "                where a.ponumber=b.ponumber" & vbCrLf
        s = s & "                and a.itemseq=b.itemseq)" & vbCrLf
        s = s & "from poitems a" & vbCrLf
        s = s & "where a.DivisionID = " & HFApp.DivisionID & " and a.genBatch=" & DbQuote(Num, mBatch) & vbCrLf
        Call HFApp.SqlExec(s)
    End If
    
    'export batch
    Select Case HFApp.Options(AccountingSystem)
        Case asIntacct:             b = True: Call SendPOsToIntacct(mBatch)
        Case asTimberline:          b = True: Call SendBatchToSage300
        Case asMasterBuilder:       b = True: Call SendBatchToSage100
        Case asQuickBooks:          b = True   'SendBatchToQuickbooks
        Case asQuickbooksOnline:    b = True   'SendBatchToQuickBooksOnline
    End Select


    s = "update pomaster set postingdate=getdate() where postingbatch=" & DbQuote(Num, mBatch)
    Call HFApp.SqlExec(s)

    On Error Resume Next
    Unload FProgress
    On Error GoTo eh
    
    'show summary file
    If chkShowBatchReport.value = vbChecked Then
        s = HFApp.SystemFolder & "System\Reports\Estimating\ExportPOBatch.rpt"
        'Call FRptViewer.ShowReport(s, True, True, "Batch", mBatch, "PostSummaries", HFApp.Options(PostSummarizedPOs))
        
        Dim c As New ZybUtil.Crystal
        Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
        On Error Resume Next
        Call c.ParameterValue("DivisionID", HFApp.DivisionID)
        Call c.ParameterValue("Batch", mBatch)
        Call c.ParameterValue("PostSummaries", HFApp.Options(PostSummarizedPOs))
        On Error GoTo eh
        Call c.PrintPreview("Print Preview")
        
    End If
    
    SendBatch = True
    Screen.MousePointer = vbDefault
    
Exit Function
eh: Call errHandler(SRCFILE & "SendBatch")
    Screen.MousePointer = vbDefault
End Function




    

Private Sub SendBatchToSage300()
On Error GoTo eh
    Dim s As String
    Dim i As Integer
    Dim rs As Recordset
    Dim LastPO As String
    Dim ThisPO As String
    Dim prevJob As String
    Dim prevExtra As String
    Dim r As Long
    'Dim f As New FRptViewer
    Dim TLDateFormat As String
    
    TLDateFormat = HFApp.Options.ValueByName("TimberlineDateFormat")
    If TLDateFormat = "" Then TLDateFormat = "mm\/dd\/yyyy"
   
    'write cost codes to jce
    Call CreatePath("export path", FilePath(PathAppend(HFApp.Options(Timberline_Data_Path), JCEFile)))
    i = FreeFile
    Open PathAppend(HFApp.Options(Timberline_Data_Path), JCEFile) For Output As #i
    s = ""
    s = s & "SELECT DISTINCT i.Job,j.description JobDesc,i.JCExtra,i.JCCostCode,i.JCCategory" & vbCrLf
    s = s & "FROM POMaster p" & vbCrLf
    s = s & "LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
    s = s & "LEFT OUTER JOIN tbljobs j ON(p.DivisionID=j.DivisionID and p.job=j.job_no)" & vbCrLf
    s = s & "WHERE PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
    s = s & "ORDER BY 1,2,3,4" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
    
        If prevJob <> "" & rs("Job") Then
            Print #i, "*," & Quote("" & rs("Job")) & "," & Quote("" & rs("JobDesc"))
        End If
        If prevExtra <> "" & rs("JCExtra") Then
            If Trim("" & rs("JCExtra")) <> "" Then Print #i, "E," & Quote("" & rs("JCExtra"))
        End If
        Print #i, "C," & Quote("" & rs("JCCostCode")) & ",," & Quote("" & rs("JCCategory")) & "," & format(Now, "mmddyyyy") & ",0,,0,,,,," & Quote("" & rs("JCExtra"))
        
        prevJob = "" & rs("Job")
        prevExtra = "" & rs("JCExtra")
        rs.MoveNext
    
    Wend
    Close i
    
    'write purchase orders to jcc
    Call CreatePath("export path", FilePath(PathAppend(HFApp.Options(Timberline_Data_Path), JCCFile)))
    i = FreeFile
    Open PathAppend(HFApp.Options(Timberline_Data_Path), JCCFile) For Output As #i
    
    s = ""
    s = s & "select po,line,podate,isnull(nullif(podesc,''),poindex) podesc,vendor,orderedby,terms,shipvia,fob,retainagerate,linedesc,linejob,lineextra,linecostcode,linecategory,linetaxgroup" & vbCrLf
    s = s & ",linetax,linecommittedquantity,linecommittedunitprice,linecommitteduom,potype,Linepretax+linetax linetotal" & vbCrLf
    s = s & "from dbo.JCPODetails" & vbCrLf
    s = s & "WHERE PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
    s = s & "order by 1,2" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        ThisPO = "" & rs("PO")
        If ThisPO <> LastPO Then
            LastPO = ThisPO
            
            s = ""
            s = s & "C"                                         'record id
            s = s & "," & Quote(ThisPO)                         'commitment id
            s = s & "," & IIf("" & rs("potype") = "Subcontract", 1, 2) 'commitment type
            s = s & "," & Quote("" & rs("PODesc"), True, 30)    'Description
            s = s & "," & Quote("" & rs("Vendor"))              'Vendor ID
            s = s & "," & Quote(format("" & rs("PODate"), TLDateFormat)) 'Date
            s = s & "," & rs("RetainageRate")                   'retainage percent
            s = s & ",1"                                        'committed to JC
            s = s & ",0"                                        'closed
            s = s & ",0"                                        'printed
            s = s & ","                                         'to address 1
            s = s & ","                                         'to address 2
            s = s & ","                                         'to city
            s = s & ","                                         'to state
            s = s & ","                                         'to ZIP code
            s = s & ","                                         'ship to address 1
            s = s & ","                                         'ship to address 2
            s = s & ","                                         'ship to city
            s = s & ","                                         'ship to state
            s = s & ","                                         'ship to ZIP code
            
            If "" & rs("potype") = "Subcontract" Then
            Else
            s = s & "," & Quote("" & rs("OrderedBy"), True, 15) 'ordered by
            s = s & "," & Quote("" & rs("Terms"), True, 15)     'terms
            s = s & "," & Quote("" & rs("ShipVia"), True, 15)   'ship via
            s = s & "," & Quote("" & rs("FOB"), True, 15)       'free on board
            End If
            Print #i, s
            r = 0
        End If
        r = r + 1
        s = ""
        s = s & "CI"                                           'record id
        s = s & "," & Quote(ThisPO)                            'commitment id
        s = s & "," & "" & rs("Line")                          'Item Number
        s = s & "," & Quote("" & rs("lineDesc"), True, 30)     'Description
        s = s & ","                                            'retainage percent
        s = s & ","                                            'delivery Date
        s = s & ","                                            'scope of work
        s = s & "," & Quote("" & rs("lineJob"))                'Job
        s = s & "," & Quote("" & rs("lineExtra"))              'extra
        s = s & "," & Quote("" & rs("lineCostCode"))           'Cost Code
        s = s & "," & Quote("" & rs("lineCategory"))           'Category
        s = s & "," & Quote("" & rs("lineTaxGroup"))           'tax Group
        s = s & "," & Round(Val("" & rs("lineTax")), 2)        'tax
        
        s = s & "," & Val("" & rs("linecommittedquantity"))    'Units
        
        If Val("" & rs("linecommittedunitprice")) > 999999 Or Val("" & rs("linetotal")) = 0 Then
            s = s & ",0"                                           'Unit Cost
        Else
            s = s & "," & Quote("" & rs("linecommittedunitprice")) 'Unit Cost
        End If
        
        s = s & "," & Quote("" & rs("linecommitteduom"), , 6)  'Unit Description
        
        
        s = s & "," & Round(Val("" & rs("linetotal")), 2)      'amount
        
        Print #i, s
        rs.MoveNext
    Wend
    
    
    'write change orders to jcc
    s = ""
    s = s & "SELECT p.PONumber+','+p.ChangeOrder ID" & vbCrLf
    s = s & "      ,p.PONumber" & vbCrLf
    s = s & "      ,p.ChangeOrder" & vbCrLf
    s = s & "      ,p.Description" & vbCrLf
    s = s & "      ,p.CODate" & vbCrLf
    s = s & "      ,i.Description ItemDesc" & vbCrLf
    s = s & "      ,i.Linenumber" & vbCrLf
    s = s & "      ,i.Qty" & vbCrLf
    s = s & "      ,isnull(i.Pretax,0)+isnull(i.jctax,0)+isnull(i.njctax,0) Amount" & vbCrLf
    s = s & "  FROM POChangeOrders p LEFT OUTER JOIN POChangeOrderItems i ON(p.PONumber=i.PONumber and p.changeorder=i.changeorder and p.DivisionID = i.DivisionID)" & vbCrLf
    s = s & " WHERE p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
    s = s & "ORDER BY p.PONumber,p.Changeorder,i.LineNumber" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        ThisPO = "" & rs("ID")
        If ThisPO <> LastPO Then
            LastPO = ThisPO
            s = ""
            s = s & "CCO"                                         'record id
            s = s & "," & Quote("" & rs("PONumber"))              'commitment id
            s = s & "," & Quote("" & rs("ChangeOrder"))           'commitment co id
            s = s & "," & Quote("" & rs("Description"), True, 30) 'Description
            s = s & ","                                           'printed
            s = s & "," & Quote(format("" & rs("CODate"), TLDateFormat)) 'Date
            s = s & ","                                           'type
            s = s & ","                                           'scope
            s = s & ",5"                                          'status
            Print #i, s
            r = 0
        End If
        r = r + 1
        s = ""
        s = s & "CCOI"                                     'Record ID
        s = s & "," & Quote("" & rs("PONumber"))           'commitment id
        s = s & "," & Quote("" & rs("ChangeOrder"))        'commitment co id
        s = s & "," & Val("" & rs("LineNumber"))           'line number
        s = s & "," & Quote("" & rs("ItemDesc"), True, 30) 'Description
        s = s & "," & Val("" & rs("Qty"))                  'Units
        s = s & "," & Round(Val("" & rs("Amount")), 2)     'amount
        Print #i, s
        rs.MoveNext
    Wend
    Close i
    
    'save datafile to batch table
    On Error Resume Next
    Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "DataFile", PathAppend(HFApp.Options(Timberline_Data_Path), JCCFile))
    On Error GoTo eh
    
    'launch macro
    Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), _
                       HFApp.Options(Timberline_UID), _
                       HFApp.Options(Timberline_PWD), _
                       PathAppend(HFApp.Options(Timberline_Data_Path), MacroFile), _
                       PathAppend(HFApp.Options(Timberline_Data_Path), ImportPrintFile), _
                       "Processing PO's...")
    DoEvents 'so fmain can paint over ftsobject image
    On Error Resume Next
    Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "RejectFile", PathAppend(HFApp.Options(Timberline_Data_Path), RejectFile))
    If FileExists(PathAppend(HFApp.Options(Timberline_Data_Path), ForceExt(ImportPrintFile, "pdf"))) Then
        Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "ImportPrintFile", PathAppend(HFApp.Options(Timberline_Data_Path), ForceExt(ImportPrintFile, "pdf")))
        Call HFApp.SqlExec("update Batches set importprintfiletype = 'pdf' WHERE Batch=" & mBatch)
    Else
        Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "ImportPrintFile", PathAppend(HFApp.Options(Timberline_Data_Path), ImportPrintFile))
        Call HFApp.SqlExec("update Batches set importprintfiletype = 'prn' WHERE Batch=" & mBatch)
    End If
    If FileExists(PathAppend(HFApp.Options(Timberline_Data_Path), ForceExt(PostPrintFile, "pdf"))) Then
        Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "PostPrintFile", PathAppend(HFApp.Options(Timberline_Data_Path), ForceExt(PostPrintFile, "pdf")))
        Call HFApp.SqlExec("update Batches set postprintfiletype = 'pdf' WHERE Batch=" & mBatch)
    Else
        Call DBPutFile(HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "PostPrintFile", PathAppend(HFApp.Options(Timberline_Data_Path), PostPrintFile))
        Call HFApp.SqlExec("update Batches set postprintfiletype = 'prn' WHERE Batch=" & mBatch)
    End If
    On Error GoTo eh

    'show processing journal
    If chkShowProcessingJournal.value = vbChecked Then
        s = TempFile(FileExt(ImportPrintFile))
        Call FileCopy(PathAppend(HFApp.Options(Timberline_Data_Path), ImportPrintFile), s)
        Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), HFApp.Options(Timberline_UID), HFApp.Options(Timberline_PWD), s)
    End If
    
    
Exit Sub
eh: Select Case True
        Case Err.Description Like ("*The text, ntext, or image pointer value conflicts with the column name specified*")
            If vbYes = MsgBox("The posting files are too large to be saved with the batch. If you continue you won't be able to review them later. Do you want to continue?", vbYesNo + vbQuestion, App.ProductName) Then
                Resume Next
            End If
            
        Case Err.Number = 70
            MsgBox "Unable to display the posting journal. It may already be open.", vbInformation
            
        Case Else
            Call errHandler(SRCFILE & "SendBatchToSage300")
    End Select
    Close
End Sub





Private Sub Form_Load()
    Dim s As String

    Call IniGetForm(Me)
    
On Error Resume Next
    s = Choose(HFApp.Options(AccountingSystem), "Timberline", "Timberline", "MasterBuilder", "Quickbooks", "Sage50", "MYOB", "Peachtree", "Xero", "Spectrum")
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
    
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        Me.lblBatchTask(4).Visible = False
    End If
    
    CurrentFrame = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    HFApp.Options.value(ShowBudgetExportReport) = chkShowBatchReport.value = vbChecked
    HFApp.Options.value(ShowBudgetProcessingJournal) = chkShowProcessingJournal.value = vbChecked
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
On Error GoTo eh
    Dim OldBatch As Long
    Dim NewBatch As Long
    
    Dim s As String
    'Dim f As New FRptViewer
    
    Select Case Index
        Case 0 'open summary report
            s = HFApp.SystemFolder & "System\Reports\Estimating\ExportPOBatch.rpt"
            'Call f.ShowReport(s, True, True, "Batch", mBatch, "PostSummaries", HFApp.Options(PostSummarizedPOs))
            
            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            Call c.ParameterValue("Batch", mBatch)
            Call c.ParameterValue("PostSummaries", HFApp.Options(PostSummarizedPOs))
            On Error GoTo eh
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
            If HFApp.SqlExec("select importprintfiletype from batches where batch=" & mBatch)(0) = "pdf" Then
                s = TempFile("pdf")
                Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "ImportPrintFile")
                Call ShellFile(Me.hwnd, s)
            Else
                s = TempFile("prn")
                Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Batches WHERE Batch=" & mBatch, "ImportPrintFile")
                Call FTSObject.Run(HFApp.Options(Timberline_Data_Path), HFApp.Options(Timberline_UID), HFApp.Options(Timberline_PWD), s)
            End If
        
        Case 4 'repost rejects
            
            OldBatch = mBatch
            Call HFApp.SqlExec("insert into batches(DivisionID,batchtype,ustmp,tstmp) values(" & HFApp.DivisionID & ",'Repost PO Batch'," & DbQuote(Str, HFApp.LoginID) & ",getdate())", dbHomefront)
            NewBatch = HFApp.SqlIdentity("batches", dbHomefront)
            s = ""
            s = s & "update batches set repostbatch=" & NewBatch & " where batch=" & OldBatch & vbCrLf
            s = s & "update pomaster set postingbatch=" & NewBatch & " where postingbatch=" & OldBatch & vbCrLf
            s = s & "update pochangeorders set postingbatch=" & NewBatch & " where postingbatch=" & OldBatch & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
            mBatch = NewBatch
            If SendBatch Then Unload Me
        
    
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "lblBatchTask_Click")
End Sub


Private Function SendBatchToQuickBooksOnline() As Boolean
On Error GoTo eh
    Dim s As String
    Dim r As String
    Dim i As Long
    Dim Count As Long
    Dim rs As Recordset
    Dim LastPO As String
    Dim ThisPO As String
    Dim IsUSVersion As Boolean
    Dim qty As Double
    Dim Amount As Double
    Dim rate As Double
    Dim Response As String
    Dim BrokerUAT As Boolean
    
    
    ' Set this to false when compiling for production use
    BrokerUAT = "" & HFApp.Options.ValueByName("QBOBrokerUAT") = "True"
    
    
    IsUSVersion = HFApp.Options.ValueByName("AccountingVersion") <> "CA"
    
    
'set posting date = null, then set each po's posting date as it is posted
    Call HFApp.SqlExec("update pomaster set postingdate=null where PostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
    
    Call FProgress.Progress("Posting purchase orders...", "querying...", 1, 2)
    s = "SELECT COUNT(*) FROM POMaster WHERE PostingBatch=" & DbQuote(Num, mBatch)
    Count = Val("" & HFApp.SqlExec(s)(0))
    
    If HFApp.Options(PostSummarizedPOs) Then
        s = ""
        'budget amounts
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,v.ExternalID as Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,min(isnull(c.Description,'')) ItemDesc" & vbCrLf
        s = s & "      ,j.ExternalJobID,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,t.externalid" & vbCrLf
        s = s & "      ,SUM(i.JCTax-i.variancejctax)  JCTax" & vbCrLf
        s = s & "      ,SUM(i.NJCTax-i.variancenjctax) NJCTax" & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        s = s & "      ,SUM(i.Pretax-i.variancepretax) Pretax" & vbCrLf
        s = s & "      ,'' PartNumber" & vbCrLf
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       JOIN tblvendors v on v.DivisionID = p.divisionId and v.Vendor_id = p.Vendor" & vbCrLf
        s = s & "       LEFT OUTER JOIN tbljobs j ON(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID = pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN taxgroups t ON(i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.jccategory=sc.category)" & vbCrLf
        s = s & " WHERE p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "GROUP BY pi.POIndex,j.ExternalJobID,p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex),p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB,p.RetainagePercent,i.Job,i.JCExtra" & vbCrLf
        s = s & "      ,i.JCCostCode" & vbCrLf
        s = s & "      ,i.JCCategory" & vbCrLf
        s = s & "      ,t.externalid,sc.externalid,c.externalid" & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        'variance amounts
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,v.ExternalID as Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,min(isnull(c.Description,'')) ItemDesc" & vbCrLf
        s = s & "      ,j.ExternalJobID,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,t.externalid" & vbCrLf
        s = s & "      ,SUM(i.varianceJCTax)  JCTax" & vbCrLf
        s = s & "      ,SUM(i.varianceNJCTax) NJCTax" & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        s = s & "      ,SUM(i.variancePretax) Pretax" & vbCrLf
        s = s & "      ,'' PartNumber" & vbCrLf
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       JOIN tblvendors v on v.DivisionID = p.divisionId and v.Vendor_id = p.Vendor" & vbCrLf
        s = s & "       LEFT OUTER JOIN tbljobs j ON(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID = pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN taxgroups t ON(i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.jccategory=sc.category)" & vbCrLf
        s = s & " WHERE p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "GROUP BY pi.POIndex,j.ExternalJobID,p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex),p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB,p.RetainagePercent,i.Job,i.JCExtra" & vbCrLf
        s = s & "      ,i.JCCostCode" & vbCrLf
        s = s & "      ,i.JCCategory" & vbCrLf
        s = s & "      ,t.externalid,c.externalid,sc.externalid" & vbCrLf
        s = s & "HAVING SUM(i.variancePretax)<>0" & vbCrLf
        s = s & "ORDER BY 1,10,11,12,13" & vbCrLf
    Else
        s = ""
        'budget amounts
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,v.ExternalID as Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,i.Description ItemDesc" & vbCrLf
        s = s & "      ,i.SortOrder,j.ExternalJobID,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,t.externalid" & vbCrLf
        s = s & "      ,i.JCTax-i.variancejctax jctax" & vbCrLf
        s = s & "      ,i.NJCTax-i.variancenjctax njctax" & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        s = s & "      ,i.Pretax-i.variancepretax pretax" & vbCrLf
        s = s & "      ,i.PartNumber" & vbCrLf
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       JOIN tblvendors v on v.DivisionID = p.divisionId and v.Vendor_id = p.Vendor" & vbCrLf
        s = s & "       LEFT OUTER JOIN tbljobs j ON(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID = pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN taxgroups t ON(i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.jccategory=sc.category)" & vbCrLf
        s = s & " WHERE p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        'variance amounts
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,v.ExternalID as Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,i.Description ItemDesc" & vbCrLf
        s = s & "      ,i.SortOrder,j.ExternalJobID,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,t.externalid" & vbCrLf
        s = s & "      ,i.varianceJCTax jctax" & vbCrLf
        s = s & "      ,i.varianceNJCTax njctax " & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        s = s & "      ,i.variancepretax Pretax" & vbCrLf
        s = s & "      ,'' PartNumber" & vbCrLf
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       JOIN tblvendors v on v.DivisionID = p.divisionId and v.Vendor_id = p.Vendor" & vbCrLf
        s = s & "       LEFT OUTER JOIN tbljobs j ON(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID =pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN taxgroups t ON(i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.jccategory=sc.category)" & vbCrLf
        s = s & " WHERE p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "   AND i.variancepretax <>0" & vbCrLf
        s = s & "ORDER BY 1,10,11,12,13,14" & vbCrLf
    End If
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Function
    
    s = ""
    s = s & "<?xml version=""" & "1.0""" & " encoding=""" & "utf-8""" & "?>" & vbCrLf
    s = s & "<PurchaseOrder xmlns:xsi=""" & "http://www.w3.org/2001/XMLSchema-instance""" & " xmlns:xsd=""" & "http://www.w3.org/2001/XMLSchema""" & ">" & vbCrLf
    i = 0
    While Not rs.EOF
        ThisPO = "" & rs("PONumber")
        If ThisPO <> LastPO Then
            If LastPO <> "" Then
                s = s & "</PurchaseOrder>" & vbCrLf
                Response = SubmitQBOXml("PostPurchaseOrder", s)
                If Response <> "" Then
                    Call HFApp.SqlExec("update pomaster set postingdate=getdate() where DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, LastPO), dbHomefront)
                End If
            End If
            
            i = i + 1
            Call FProgress.Progress(, "processing...", i, Count)
            s = ""
            s = s & "<?xml version=""" & "1.0""" & " encoding=""" & "utf-8""" & "?>" & vbCrLf
            s = s & "<PurchaseOrder xmlns:xsi=""" & "http://www.w3.org/2001/XMLSchema-instance""" & " xmlns:xsd=""" & "http://www.w3.org/2001/XMLSchema""" & ">" & vbCrLf
            s = s & "<APAccount>" & HFApp.Options.ValueByName("APAccount") & "</APAccount>" & vbCrLf
            s = s & HFApp.XmlQBAdd(c, 40, "Vendor", "" & rs("Vendor"))
            s = s & HFApp.XmlQBAdd(c, 40, "DocNumber", "" & ThisPO)
            s = s & HFApp.XmlQBAdd(d, 0, "PODate", "" & rs("PODate"))
            LastPO = ThisPO
        End If
        s = s & "<Line>" & vbCrLf
        s = s & vbTab & HFApp.XmlQBAdd(c, 40, "JCCostCode", "" & rs("JCCostCode"))
        s = s & vbTab & HFApp.XmlQBAdd(c, 4095, "Description", IIf("" & rs("ItemDesc") = "", "untitled", "" & rs("ItemDesc")))
        s = s & vbTab & HFApp.XmlQBAdd(c, 40, "Job", "" & rs("ExternalJobID"))
        If IsUSVersion Then
            qty = Val("" & rs("Qty"))
            If qty = 0 Then qty = 1
            Amount = (Val("" & rs("Pretax")) + Val("" & rs("JCTax")))
            rate = Amount / qty
            s = s & vbTab & HFApp.XmlQBAdd(n, 7.5, "Qty", qty)
            s = s & vbTab & HFApp.XmlQBAdd(c, 40, "JCCategory", "" & rs("JCCategory"))
            s = s & vbTab & HFApp.XmlQBAdd(n, 7.5, "UnitPrice", rate)
            s = s & vbTab & HFApp.XmlQBAdd(n, 8.2, "Amount", Amount)
        Else
            qty = Val("" & rs("Qty"))
            If qty = 0 Then qty = 1
            rate = Amount / qty
            Amount = Val("" & rs("Pretax"))
            s = s & vbTab & HFApp.XmlQBAdd(n, 7.5, "Qty", qty)
            s = s & vbTab & HFApp.XmlQBAdd(c, 40, "JCCategory", "" & rs("JCCategory"))
            s = s & vbTab & HFApp.XmlQBAdd(n, 7.5, "UnitPrice", rate)
            s = s & vbTab & HFApp.XmlQBAdd(n, 8.2, "Amount", Amount)
            s = s & vbTab & HFApp.XmlQBAdd(c, 40, "TaxGroup", "" & rs("ExternalID"))
        End If
        s = s & "</Line>" & vbCrLf
        rs.MoveNext
    Wend
    s = s & "</PurchaseOrder>" & vbCrLf
    
    Response = SubmitQBOXml("PostPurchaseOrder", s)
    If Response <> "" Then
        Call HFApp.SqlExec("update pomaster set postingdate=getdate() where DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, LastPO), dbHomefront)
    End If
    SendBatchToQuickBooksOnline = True
    
Exit Function
eh: Select Case Err.Source
    Case "SubmitQBXml":  MsgBox "Unable to post PO number """ & LastPO & """" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.ProductName
    Case Else:           Call errHandler(SRCFILE & "SendBatchToQuickBooksOnline", s)
    End Select
    
    Call HFApp.SqlExec("update pomaster set postingbatch=0 where PostingDate IS NULL and PostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
End Function
Private Function SendBatchToQuickbooks() As Boolean
On Error GoTo eh
    Dim s As String
    Dim r As String
    Dim i As Long
    Dim Count As Long
    Dim rs As Recordset
    Dim LastPO As String
    Dim ThisPO As String
    Dim IsUSVersion As Boolean
    Dim qty As Double
    Dim Amount As Double
    
    IsUSVersion = HFApp.Options.ValueByName("AccountingVersion") <> "CA"
    
    
'set posting date = null, then set each po's date as it is posted
    Call HFApp.SqlExec("update pomaster set postingdate=null where PostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
    
    Call FProgress.Progress("Posting purchase orders...", "querying...", 1, 2)
    s = "SELECT COUNT(*) FROM POMaster WHERE PostingBatch=" & DbQuote(Num, mBatch)
    Count = Val("" & HFApp.SqlExec(s)(0))
    
    If HFApp.Options(PostSummarizedPOs) Then
        s = ""
        'budget amounts
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,min(isnull(c.Description,'')) ItemDesc" & vbCrLf
        s = s & "      ,j.ExternalJobID,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,t.externalid" & vbCrLf
        s = s & "      ,SUM(i.JCTax-i.variancejctax)  JCTax" & vbCrLf
        s = s & "      ,SUM(i.NJCTax-i.variancenjctax) NJCTax" & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        s = s & "      ,SUM(i.Pretax-i.variancepretax) Pretax" & vbCrLf
        s = s & "      ,'' PartNumber" & vbCrLf
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       LEFT OUTER JOIN tbljobs j ON(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID = pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN taxgroups t ON(i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.jccategory=sc.category)" & vbCrLf
        s = s & " WHERE len(p.ponumber)<12 and p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "GROUP BY pi.POIndex,j.ExternalJobID,p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex),p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB,p.RetainagePercent,i.Job,i.JCExtra" & vbCrLf
        s = s & "      ,i.JCCostCode" & vbCrLf
        s = s & "      ,i.JCCategory" & vbCrLf
        s = s & "      ,t.externalid,sc.externalid,c.externalid" & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        'variance amounts
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,min(isnull(c.Description,'')) ItemDesc" & vbCrLf
        s = s & "      ,j.ExternalJobID,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,t.externalid" & vbCrLf
        s = s & "      ,SUM(i.varianceJCTax)  JCTax" & vbCrLf
        s = s & "      ,SUM(i.varianceNJCTax) NJCTax" & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        s = s & "      ,SUM(i.variancePretax) Pretax" & vbCrLf
        s = s & "      ,'' PartNumber" & vbCrLf
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       LEFT OUTER JOIN tbljobs j ON(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID = pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN taxgroups t ON(i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.jccategory=sc.category)" & vbCrLf
        s = s & " WHERE len(p.ponumber)<12 and p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "GROUP BY pi.POIndex,j.ExternalJobID,p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex),p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB,p.RetainagePercent,i.Job,i.JCExtra" & vbCrLf
        s = s & "      ,i.JCCostCode" & vbCrLf
        s = s & "      ,i.JCCategory" & vbCrLf
        s = s & "      ,t.externalid,c.externalid,sc.externalid" & vbCrLf
        s = s & "HAVING SUM(i.variancePretax)<>0" & vbCrLf
        s = s & "ORDER BY 1,10,11,12,13" & vbCrLf
    Else
        s = ""
        'budget amounts
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,i.Description ItemDesc" & vbCrLf
        s = s & "      ,i.SortOrder,j.ExternalJobID,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,t.externalid" & vbCrLf
        s = s & "      ,i.JCTax-i.variancejctax jctax" & vbCrLf
        s = s & "      ,i.NJCTax-i.variancenjctax njctax" & vbCrLf
        's = s & "      ,i.OrderQty Qty" & vbCrLf
        's = s & "      ,i.OrderUOM UOM" & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        s = s & "      ,i.Pretax-i.variancepretax pretax" & vbCrLf
        s = s & "      ,i.PartNumber" & vbCrLf
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       LEFT OUTER JOIN tbljobs j ON(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID = pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN taxgroups t ON(i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.jccategory=sc.category)" & vbCrLf
        s = s & " WHERE len(p.ponumber)<12 and p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        'variance amounts
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,i.Description ItemDesc" & vbCrLf
        s = s & "      ,i.SortOrder,j.ExternalJobID,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,t.externalid" & vbCrLf
        s = s & "      ,i.varianceJCTax jctax" & vbCrLf
        s = s & "      ,i.varianceNJCTax njctax " & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        s = s & "      ,i.variancepretax Pretax" & vbCrLf
        s = s & "      ,'' PartNumber" & vbCrLf
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       LEFT OUTER JOIN tbljobs j ON(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID =pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN taxgroups t ON(i.DivisionID = t.DivisionID and i.taxgroup=t.taxgroup)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.variancejccategory=sc.category)" & vbCrLf
        s = s & " WHERE len(p.ponumber)<12" & vbCrLf
        s = s & "   AND p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "   AND i.variancepretax <>0" & vbCrLf
        s = s & "ORDER BY 1,10,11,12,13,14" & vbCrLf
    End If
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Function
    
    s = HFApp.XmlQBStart()
    i = 0
    While Not rs.EOF
        ThisPO = "" & rs("PONumber")
        If ThisPO <> LastPO Then
            If LastPO <> "" Then
                s = s & "</PurchaseOrderAdd></PurchaseOrderAddRq>" & vbCrLf
                s = s & HFApp.XmlQBEnd()
                Call HFApp.XmlQBSubmit(s)
                Call HFApp.SqlExec("update pomaster set postingdate=getdate() where PONumber=" & DbQuote(Str, LastPO), dbHomefront)
            End If
            
            i = i + 1
            Call FProgress.Progress(, "processing...", i, Count)
            s = HFApp.XmlQBStart()
            s = s & "<PurchaseOrderAddRq><PurchaseOrderAdd>" & vbCrLf
            s = s & "<VendorRef>" & HFApp.XmlQBAdd(c, 40, "ListID", "" & rs("Vendor")) & "</VendorRef>" & vbCrLf
            s = s & "<ShipToEntityRef>" & HFApp.XmlQBAdd(c, 40, "ListID", "" & rs("ExternalJobID")) & "</ShipToEntityRef>" & vbCrLf
            s = s & HFApp.XmlQBAdd(d, 0, "TxnDate", "" & rs("PODate"))
            s = s & HFApp.XmlQBAdd(c, 11, "RefNumber", ThisPO)
            s = s & HFApp.XmlQBAdd(c, 4095, "Memo", "" & rs("PODesc"))
            LastPO = ThisPO
        End If
        s = s & "<PurchaseOrderLineAdd>" & vbCrLf
        s = s & "<ItemRef>" & HFApp.XmlQBAdd(c, 40, "ListID", "" & rs("JCCostCode")) & "</ItemRef>" & vbCrLf
        s = s & HFApp.XmlQBAdd(c, 4095, "Desc", IIf("" & rs("ItemDesc") = "", "untitled", "" & rs("ItemDesc")))
        If IsUSVersion Then
            qty = Val("" & rs("Qty"))
            Amount = (Val("" & rs("Pretax")) + Val("" & rs("JCTax")))
            s = s & HFApp.XmlQBAdd(n, 7.5, "Quantity", qty)
            s = s & "<ClassRef>" & HFApp.XmlQBAdd(c, 40, "ListID", "" & rs("JCCategory")) & "</ClassRef>" & vbCrLf
            s = s & HFApp.XmlQBAdd(n, 8.2, "Amount", Amount)
        Else
            qty = Val("" & rs("Qty"))
            Amount = Val("" & rs("Pretax"))
            s = s & HFApp.XmlQBAdd(n, 7.5, "Quantity", qty)
            s = s & "<ClassRef>" & HFApp.XmlQBAdd(c, 40, "ListID", "" & rs("JCCategory")) & "</ClassRef>" & vbCrLf
            s = s & HFApp.XmlQBAdd(n, 8.2, "Amount", Amount)
            s = s & "<SalesTaxCodeRef>" & HFApp.XmlQBAdd(c, 40, "ListID", "" & rs("ExternalID")) & "</SalesTaxCodeRef>" & vbCrLf
        End If
        s = s & "</PurchaseOrderLineAdd>" & vbCrLf
        rs.MoveNext
    Wend
    s = s & "</PurchaseOrderAdd></PurchaseOrderAddRq>" & vbCrLf
    s = s & HFApp.XmlQBEnd()
    Call HFApp.XmlQBSubmit(s)
    
    Call HFApp.SqlExec("update pomaster set postingdate=getdate() where DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, LastPO), dbHomefront)
    
    SendBatchToQuickbooks = True
    
Exit Function
eh: Select Case Err.Source
    Case "SubmitQBXml":  MsgBox "Unable to post PO number """ & LastPO & """" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.ProductName
    Case Else:           Call errHandler(SRCFILE & "SendBatchToQuickbooks", s)
    End Select
    
    Call HFApp.SqlExec("update pomaster set postingbatch=0 where PostingDate IS NULL and PostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
End Function

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


Private Sub SendBatchToSage100()
On Error GoTo eh
    Dim s As String
    Dim X As String
    Dim rs As Recordset
    Dim LastPO As String
    Dim POType As String
    Dim ThisPO As String
    
    Dim qty As Double
    Dim rate As Double
    Dim Amount As Double
    
    Dim i As Long
    Dim Count As Long
    
    Dim DontUseParts As Boolean
    Dim UseJobAsSubAcct As Boolean
    Dim isCanadian As Boolean
    
    'we dont support canadian version anymore.
    isCanadian = False 'HFApp.Options.ValueByName("MasterBuilderEdition") = "CA"

    'NOTES: - Masterbuilder job phase is equivilent to jcextra
    '       - Tried hard to implement variance reporting in masterbuilder but since we can only post qty and rate
    '         we decided that we arent going to post variances.
    
        
    'set all posting dates to null, then post each po individually. set each po's date as it is posted
    'when a po fails an error is thrown and processing is stopped. the po's with a posting date are marked
    'as posted, the rest are removed from the batch.
    Call HFApp.SqlExec("update pomaster set postingdate=null where PostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
    
    'get options
    DontUseParts = HFApp.Options.ValueByName("DontUseMBParts") = "True"
    UseJobAsSubAcct = HFApp.Options.ValueByName("UseJobAsSubAcct") = "True"
    
    
    'select PO's
    Call FProgress.Progress("Posting purchase orders...", "querying...", 1, 2)
    s = "SELECT COUNT(*) FROM POMaster WHERE PostingBatch=" & DbQuote(Num, mBatch)
    Count = Val("" & HFApp.SqlExec(s)(0))
    If HFApp.Options(PostSummarizedPOs) Then
        s = ""
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,min(isnull(c.Description,'')) ItemDesc" & vbCrLf
        s = s & "      ,0,i.Job,i.JCExtra,c.ExternalID JCCostCode,cat.ExternalID JCCategory,i.TaxGroup,c.DebitAccount" & vbCrLf
        s = s & "      ,SUM(i.JCTax+i.NJCTax) Tax" & vbCrLf
        s = s & "      ,1 Qty" & vbCrLf
        s = s & "      ,0 Rate" & vbCrLf
        s = s & "      ,'' UOM" & vbCrLf
        
        'posted amount should be pretax only
        's = s & "      ,SUM(i.Pretax+i.JCTax+i.NJCTax) Amount" & vbCrLf
        s = s & "      ,SUM(i.Pretax) Amount" & vbCrLf
        
        s = s & "      ,'' PartNumber,pi.potype" & vbCrLf
        If isCanadian Then
            s = s & "      ,case when t.rate1=0 then 'N' else 'Y' end gst" & vbCrLf
            s = s & "      ,case when t.rate2=0 then 'N' else 'Y' end pst" & vbCrLf
            s = s & "      ,case when t.rate3=0 then 'N' else 'Y' end hst" & vbCrLf
        End If
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID = pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories cat ON(i.DivisionID = c.DivisionID and i.jccategory=cat.category)" & vbCrLf
        If isCanadian Then
            s = s & "       LEFT OUTER JOIN taxgroups t on (i.divisionid=t.divisionid and i.taxgroup=t.taxgroup)" & vbCrLf
        End If
        s = s & " WHERE p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "GROUP BY p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex),p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB,p.RetainagePercent,i.Job,i.JCExtra" & vbCrLf
        s = s & "      ,c.externalid" & vbCrLf
        s = s & "      ,cat.externalid" & vbCrLf
        s = s & "      ,i.TaxGroup,c.debitaccount,pi.potype"
        If isCanadian Then
            s = s & "      ,t.rate1,t.rate2,t.rate3" & vbCrLf
        End If
        s = s & "ORDER BY 1,10,11,12,13,14"
    Else
        s = ""
        s = s & "SELECT p.PONumber,p.PODate,isnull(nullif(p.Description,''),pi.POIndex) PODesc,p.Vendor,p.OrderedBy,p.Terms,p.ShipVia,p.FOB" & vbCrLf
        s = s & "      ,p.RetainagePercent,i.Description ItemDesc" & vbCrLf
        s = s & "      ,i.SortOrder,0,i.Job,i.JCExtra,c.externalid JCCostCode,sc.externalid JCCategory,i.TaxGroup,Case when isnull(c.debitaccount,'')<>'' then c.debitaccount else isnull(sc.debitaccount,'') end DebitAccount" & vbCrLf
        s = s & "      ,i.JCTax+i.NJCTax Tax" & vbCrLf
        
        s = s & "      ,i.OrderQty Qty" & vbCrLf
        s = s & "      ,i.Rate" & vbCrLf
        
        s = s & "      ,i.OrderUOM UOM" & vbCrLf
        s = s & "      ,i.Pretax Amount" & vbCrLf
        s = s & "      ,i.PartNumber,pi.potype" & vbCrLf
        If isCanadian Then
            s = s & "      ,case when t.rate1=0 then 'N' else 'Y' end gst" & vbCrLf
            s = s & "      ,case when t.rate2=0 then 'N' else 'Y' end pst" & vbCrLf
            s = s & "      ,case when t.rate3=0 then 'N' else 'Y' end hst" & vbCrLf
        End If
        s = s & "  FROM POMaster p" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblpoindex pi ON(p.DivisionID = pi.DivisionID and p.poindex=pi.poindex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN POItems i ON(p.PONumber=i.PONumber and p.DivisionID = i.DivisionID)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcostcodes c ON(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "       LEFT OUTER JOIN standardcategories sc ON(i.DivisionID = sc.DivisionID and i.JCCategory=sc.Category)" & vbCrLf
        If isCanadian Then
            s = s & "       LEFT OUTER JOIN taxgroups t on (i.divisionid=t.divisionid and i.taxgroup=t.taxgroup)" & vbCrLf
        End If
        s = s & " WHERE p.PostingBatch=" & DbQuote(Num, mBatch) & vbCrLf
        s = s & "ORDER BY 1,11,12,13,14,15,16"
    End If
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then
        i = 0
        LastPO = ""
        While Not rs.EOF
            ThisPO = "" & rs("PONumber")
            If ThisPO <> LastPO Then
            
                'add header row
                If LastPO <> "" Then
                    X = X & "</" & POType & "AddRq>" & vbCrLf
                    X = X & HFApp.XmlMBEnd()
                    On Error Resume Next
                    s = HFApp.XmlMbSubmit(X, HFApp.Options(MasterBuilderPWD))
                    If Err.Number = 0 Then
                        Call HFApp.SqlExec("update pomaster set postingdate=getdate() where PONumber=" & DbQuote(Str, LastPO), dbHomefront)
                    Else
                        MsgBox "Unable to post PO number """ & LastPO & """" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.ProductName
                    End If
                    On Error GoTo eh
                End If
                
                i = i + 1
                Call FProgress.Progress(, "processing...", i, Count)
                X = HFApp.XmlMbStart(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
                POType = IIf("" & rs("POType") = "subcontract", "Subcontract", "PurchaseOrder")
                LastPO = ThisPO
                X = X & "<" & POType & "AddRq requestID=""1"">" & vbCrLf
                If POType = "subcontract" Then
                    X = X & HFApp.XmlMBAdd(c, 15, "SubcontractNumber", "" & rs("PONumber"))
                    X = X & HFApp.XmlMBAdd(n, 10, "VendorRef", "" & rs("Vendor"))
                    X = X & HFApp.XmlMBAdd(n, 10, "JobRef", "" & rs("Job"))
                    X = X & HFApp.XmlMBAdd(n, 10, "PhaseRef", Val("" & HFApp.SqlExec("select case count(distinct jcextra) when 1 then max(jcextra) else '' end from poitems where ponumber=" & DbQuote(Str, ThisPO), dbHomefront)(0)))
                    X = X & HFApp.XmlMBAdd(c, 25, "Desc", IIf("" & rs("PODesc") = "", "untitled", "" & rs("PODesc")))
                    If Not isCanadian Then
                        X = X & HFApp.XmlMBAdd(n, 6, "SalesTaxDistrictRef", "" & rs("TaxGroup"))
                    End If
                    X = X & HFApp.XmlMBAdd(n, 2, "ContractStatus", "3")
                    X = X & HFApp.XmlMBAdd(d, 0, "ContractDate", "" & rs("PODate"))
                Else
                    X = X & HFApp.XmlMBAdd(c, 15, "RefNumber", ThisPO)
                    X = X & HFApp.XmlMBAdd(d, 0, "OrderDate", "" & rs("PODate"))
                    X = X & HFApp.XmlMBAdd(n, 10, "VendorRef", "" & rs("Vendor"))
                    X = X & HFApp.XmlMBAdd(n, 10, "JobRef", "" & rs("Job"))
                    X = X & HFApp.XmlMBAdd(n, 10, "PhaseRef", Val("" & HFApp.SqlExec("select case count(distinct jcextra) when 1 then max(jcextra) else '' end from poitems where ponumber=" & DbQuote(Str, ThisPO), dbHomefront)(0)))
                    X = X & HFApp.XmlMBAdd(c, 25, "Desc", IIf("" & rs("PODesc") = "", "untitled", "" & rs("PODesc")))
                    If Not isCanadian Then
                        X = X & HFApp.XmlMBAdd(n, 6, "SalesTaxDistrictRef", "" & rs("TaxGroup"))
                    End If
                    X = X & HFApp.XmlMBAdd(c, 20, "Via", "" & rs("ShipVia"))
                    X = X & HFApp.XmlMBAdd(n, 2, "PurchaseOrderStatus", "1")
                End If
            End If
            
            'add line row
            X = X & "<" & POType & "LineAdd>" & vbCrLf
            
            Amount = Val("" & rs("Amount"))
            qty = Val("" & rs("Qty"))
            rate = Val("" & rs("Rate"))
            If Not PostPOQtyToAccounting Then
                qty = 1
                rate = Amount
            End If
            If Amount < 0 Then
                qty = -1 * Abs(qty)
            Else
                qty = Abs(qty)
            End If
            rate = Abs(rate)


            If POType = "subcontract" Then
                X = X & HFApp.XmlMBAdd(c, 50, "Desc", IIf("" & rs("ItemDesc") = "", "untitled", "" & rs("ItemDesc")))
                X = X & HFApp.XmlMBAdd(n, 15.3, "CostCodeRef", "" & rs("JCCostCode"))
                X = X & HFApp.XmlMBAdd(n, 2, "CostTypeRef", "" & rs("JCCategory"))
                X = X & HFApp.XmlMBAdd(n, 12.2, "Amount", Amount)
                If isCanadian Then
                    X = X & HFApp.XmlMBAdd(c, 1, "SubjectToGST", "" & rs("GST"))
                    X = X & HFApp.XmlMBAdd(c, 1, "SubjectToPST", "" & rs("PST"))
                    X = X & HFApp.XmlMBAdd(c, 1, "SubjectToHST", "" & rs("HST"))
                End If
            Else
                X = X & HFApp.XmlMBAdd(n, 15, "PartRef", IIf(DontUseParts, "" & rs("PartNumber"), ""))
                X = X & HFApp.XmlMBAdd(c, 75, "Desc", IIf("" & rs("ItemDesc") = "", "untitled", "" & rs("ItemDesc")))
                X = X & HFApp.XmlMBAdd(c, 10, "Unit", "" & rs("UOM"))
                
                X = X & HFApp.XmlMBAdd(n, 15.6, "Quantity", qty)
                X = X & HFApp.XmlMBAdd(n, 15.6, "UnitPrice", rate)
                
                If isCanadian Then
                    X = X & HFApp.XmlMBAdd(c, 1, "SubjectToGST", "" & rs("GST"))
                    X = X & HFApp.XmlMBAdd(c, 1, "SubjectToPST", "" & rs("PST"))
                    X = X & HFApp.XmlMBAdd(c, 1, "SubjectToHST", "" & rs("HST"))
                End If
                X = X & HFApp.XmlMBAdd(n, 15.3, "CostCodeRef", "" & rs("JCCostCode"))
                X = X & HFApp.XmlMBAdd(n, 2, "CostTypeRef", "" & rs("JCCategory"))
                X = X & HFApp.XmlMBAdd(c, 10, "AccountRef", "" & rs("DebitAccount"))
                If UseJobAsSubAcct And "" & rs("DebitAccount") <> "" Then
                    X = X & HFApp.XmlMBAdd(c, 10, "SubaccountRef", "" & rs("Job"))
                End If
            End If
            X = X & "</" & POType & "LineAdd>" & vbCrLf
            rs.MoveNext
        Wend
        X = X & "</" & POType & "AddRq>" & vbCrLf
        X = X & HFApp.XmlMBEnd()
        
        
        On Error Resume Next
        s = HFApp.XmlMbSubmit(X, HFApp.Options(MasterBuilderPWD))
        If Err.Number = 0 Then
            Call HFApp.SqlExec("update pomaster set postingdate=getdate() where PONumber=" & DbQuote(Str, LastPO), dbHomefront)
        Else
            Call errHandler(SRCFILE & "SendBatchToSage100", X)
        End If
        On Error GoTo eh
        
    End If
    
    Call HFApp.SqlExec("update pomaster set postingbatch=0 where PostingDate IS NULL and PostingBatch=" & DbQuote(Num, mBatch), dbHomefront)
    
        
Exit Sub
eh: Select Case Err.Source
    Case "SubmitMBXml":  MsgBox "Unable to post PO number """ & LastPO & """" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.ProductName
    Case Else:           Call errHandler(SRCFILE & "SendBatchToSage100", X)
    End Select
End Sub


