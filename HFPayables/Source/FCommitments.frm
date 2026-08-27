VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FCommitments 
   Caption         =   "<CommitmentCaption> List"
   ClientHeight    =   8175
   ClientLeft      =   2325
   ClientTop       =   2280
   ClientWidth     =   12165
   ControlBox      =   0   'False
   Icon            =   "FCommitments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12165
   Begin VB.Frame FrameSearch 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Caption         =   "4"
      Height          =   1515
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   12645
      Begin VB.TextBox txtPONumber 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1500
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optCompleted 
         BackColor       =   &H80000018&
         Caption         =   "Show Incomplete PO's only"
         Height          =   285
         Index           =   2
         Left            =   6600
         TabIndex        =   9
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton optCompleted 
         BackColor       =   &H80000018&
         Caption         =   "Show Completed PO's only"
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   8
         Top             =   390
         Width           =   3255
      End
      Begin VB.OptionButton optCompleted 
         BackColor       =   &H80000018&
         Caption         =   "Show All PO's"
         Height          =   285
         Index           =   0
         Left            =   6600
         TabIndex        =   7
         Top             =   180
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.TextBox txtTitle4 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4440
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtTitle3 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4440
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtTitle2 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4440
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtTitle1 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4440
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1500
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtJob 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1500
         TabIndex        =   0
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   315
         Left            =   6600
         TabIndex        =   13
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblPONumber 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<PONumber>"
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
         Left            =   435
         TabIndex        =   29
         Top             =   1200
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "FCommitments.frx":000C
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Desc"
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
         Left            =   720
         TabIndex        =   15
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label4 
         Height          =   975
         Index           =   0
         Left            =   4425
         TabIndex        =   28
         Top             =   465
         Width           =   1725
      End
      Begin VB.Label lblTitle4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Title4>"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3765
         TabIndex        =   19
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label lblTitle3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Title3>"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3765
         TabIndex        =   18
         Top             =   960
         Width           =   570
      End
      Begin VB.Label lblTitle2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Title2>"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3765
         TabIndex        =   17
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblTitle1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Title1>"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3765
         TabIndex        =   16
         Top             =   480
         Width           =   570
      End
      Begin VB.Label lblJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " <Job>"
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
         Left            =   915
         TabIndex        =   14
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Select <CommitmentCaption>s and items from the list by marking the check box"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   660
         TabIndex        =   27
         Top             =   60
         Width           =   6645
      End
      Begin VB.Label Label4 
         Height          =   735
         Index           =   1
         Left            =   1485
         TabIndex        =   26
         Top             =   705
         Width           =   1725
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000F&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2.30016e6
      End
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   11
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   12
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame TotalsFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   435
      Left            =   120
      TabIndex        =   20
      Top             =   7620
      Width           =   2655
      Begin VB.Label lblSelected 
         Alignment       =   1  'Right Justify
         Caption         =   "$9,999,999.99"
         Height          =   195
         Left            =   1500
         TabIndex        =   24
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice Remaining"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1380
      End
      Begin VB.Label lblRemaining 
         Alignment       =   1  'Right Justify
         Caption         =   "$9,999,999.99"
         Height          =   195
         Left            =   1500
         TabIndex        =   22
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Amount Selected"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   23
         Top             =   210
         Width           =   1380
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   5595
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   12015
      _cx             =   1988055753
      _cy             =   1988044429
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FCommitments.frx":08D6
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   3
      PicturesOver    =   0   'False
      FillStyle       =   0
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
   Begin VB.Menu mnuGrid 
      Caption         =   "<mnuGrid>"
      Visible         =   0   'False
      Begin VB.Menu mnuGridSub 
         Caption         =   "Sort Ascending"
         Index           =   0
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "Sort Descending"
         Index           =   1
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuGridSub 
         Caption         =   "Remove this column"
         Index           =   3
      End
      Begin VB.Menu mnuColumns 
         Caption         =   "Insert a column"
         Begin VB.Menu mnuColumnsSub 
            Caption         =   "(none available)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "FCommitments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FCommitments::"
Private mDebugStr As String

Private mInvoice       As Double
Private mVendor        As String

Private mCancel        As Boolean
Private mWhereClause   As String
Private mCommitmentCSV As String

'grid menu
Private MouseGrid As VSFlexGrid
Private MouseCol  As Long
Private Const mcGRID_ASC = 0
Private Const mcGRID_DESC = 1
Private Const mcGRID_HIDE = 3


Public Property Get WhereClause() As String
    WhereClause = mWhereClause
End Property
Public Property Get CommitmentCSV() As String
    CommitmentCSV = mCommitmentCSV
End Property

Public Function Choose(Vendor As String, Job As String, InvoiceAmount As Double) As Boolean
On Error GoTo eh
    
    Screen.MousePointer = vbHourglass
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    optCompleted(Val(IniGet(AppIni, "Options", "CompletedFilter"))).Value = True
    Call SetCaptions
    
    mVendor = Vendor
    txtJob.Text = Job
    mInvoice = InvoiceAmount
    
    gData.ColHidden(gData.ColIndex("Vendor")) = Not App.Options(AllowCrossPayingPOs)
    
    Call RecalcTotals
    Call LoadData(-1)
    gData.Col = 3
    Screen.MousePointer = vbDefault
    
    mCancel = False
    ReleaseCapture
    Me.Show vbModal
    Choose = Not mCancel

Exit Function:
eh: Call ErrHandler(SRCFILE & "Choose")
End Function

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    Dim r As Long
    Dim s As String
    Select Case Index
        Case 0 'OK
            With gData
                r = 1
                mCommitmentCSV = ""
                While IsBetween(r, 1, .Rows - 1)
                    If .Cell(flexcpChecked, r, 0) = flexChecked Then
                        mCommitmentCSV = mCommitmentCSV & "," & DbQuote(Str, Trim(.TextMatrix(r, .ColIndex("Commitment"))))
                        If .RowOutlineLevel(r) = 0 Then
                            'this is a po
                            If AccountingDB = dbTimberlinePVdata Then
                                s = s & " or (isub=" & DbQuote(Str, FormatTSField(12, .TextMatrix(r, .ColIndex("Commitment")))) & ")" & vbCrLf
                            Else
                                s = s & " or (isub=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Commitment"))) & ")" & vbCrLf
                            End If
                            r = .GetNodeRow(r, flexNTNextSibling)
                        Else
                            'this is a po item
                            If AccountingDB = dbTimberlinePVdata Then
                                s = s & " or (isub=" & DbQuote(Str, FormatTSField(12, .TextMatrix(r, .ColIndex("Commitment")))) & " AND item=" & Val(.TextMatrix(r, .ColIndex("item"))) & ")" & vbCrLf
                            Else
                                s = s & " or (isub=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Commitment"))) & " AND item=" & Val(.TextMatrix(r, .ColIndex("item"))) & ")" & vbCrLf
                            End If
                            r = r + 1
                        End If
                    Else
                        r = r + 1
                    End If
                Wend
            End With
            mCommitmentCSV = Mid(mCommitmentCSV, 2)
            mWhereClause = Mid(s, 5)
            mCancel = mWhereClause = ""
            
        Case 1 'Cancel
            mCancel = True
            
    End Select
    Unload Me

Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Sub


Private Sub cmdSearch_Click()
    Call LoadData(-1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 And Shift = vbCtrlMask Then
        MsgBox mDebugStr, vbInformation, App.ProductName
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
    Select Case True
        Case optCompleted(0).Value: Call IniPut(AppIni, "Options", "CompletedFilter", 0)
        Case optCompleted(1).Value: Call IniPut(AppIni, "Options", "CompletedFilter", 1)
        Case optCompleted(2).Value: Call IniPut(AppIni, "Options", "CompletedFilter", 2)
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    
    FrameSearch.Width = Me.ScaleWidth
    
    gData.Move 0, FrameSearch.Height, Me.ScaleWidth, Me.ScaleHeight - cmdNav(0).Height - 2 * margin - FrameSearch.Height
    TotalsFrame.Move margin, gData.Top + gData.Height + margin
    cmdNav(0).Move Me.ScaleWidth - 2 * cmdNav(0).Width - 2 * margin, Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(1).Move Me.ScaleWidth - cmdNav(0).Width - margin, Me.ScaleHeight - cmdNav(0).Height - margin
End Sub


Private Sub LoadData(parentRow As Long)
On Error GoTo eh
    Dim t As Single
    Dim s  As String
    Dim rs As Recordset
    Dim c  As Long
    Dim r  As Long
    Dim j  As String
    
    
    mDebugStr = ""
    Screen.MousePointer = vbHourglass
    With gData
        .Redraw = flexRDNone
        .ExplorerBar = flexExSortAndMove
        If parentRow = -1 Then
            'load po's
            .Rows = 1
            
            If Not TimberlineAccounting Then
                'use hf as po source
                s = ""
                s = s & "select" & vbCrLf
                s = s & "  p.po              Sub" & vbCrLf
                s = s & " ,p.Vendor          sVendor" & vbCrLf
                s = s & " ,p.vendordesc      sDesc" & vbCrLf
                s = s & " ,p.pojob           sJob" & vbCrLf
                s = s & " ,p.pojobdesc       jDesc" & vbCrLf
                s = s & " ,p.postarted       sactsd" & vbCrLf
                s = s & " ,p.pocompleted     sactcd" & vbCrLf
                s = s & " ,sum(p.linepretax+p.linetax) OrigAmt" & vbCrLf
                s = s & " ,sum(p.linechangespretax+p.linechangestax) ChangeAmt" & vbCrLf
                s = s & " ,sum(p.lineinvoicedpretax+p.lineinvoicedtax) Inv" & vbCrLf
                s = s & " ,sum(p.lineremainingpretax+p.lineremainingtax) Rem" & vbCrLf
                s = s & "from jcpodetails p" & vbCrLf
                s = s & "left outer join tbljobs j on p.pojob=j.job_no and p.divisionid=j.divisionid" & vbCrLf
                s = s & "where p.divisionid =" & DbQuote(num, HFApp.DivisionID) & vbCrLf
                s = s & "and p.vendor=" & DbQuote(Str, mVendor) & vbCrLf
                If HFApp.Options(AccountingSystem) = asTimberline Then s = s & "and p.postingbatch<>0" & vbCrLf
                If txtJob.Text <> "" Then s = s & "and p.pojob like " & DbQuote(Str, "%" & Trim(txtJob.Text) & "%") & vbCrLf
                If optCompleted(1).Value Then s = s & "and p.pocompleted is not null" & vbCrLf
                If optCompleted(2).Value Then s = s & "and p.pocompleted is null" & vbCrLf
                If txtDescription.Text <> "" Then s = s & "and p.pojobdesc like " & DbQuote(Str, "%" & Trim(txtDescription.Text) & "%") & vbCrLf
                If txtPONumber.Text <> "" Then s = s & "and p.po like " & DbQuote(Str, "%" & Trim(txtPONumber.Text) & "%") & vbCrLf
                If txtTitle1.Text <> "" Then s = s & "and j.pm like " & DbQuote(Str, "%" & Trim(txtTitle1.Text) & "%") & vbCrLf
                If txtTitle2.Text <> "" Then s = s & "and j.purchaser like " & DbQuote(Str, "%" & Trim(txtTitle2.Text) & "%") & vbCrLf
                If txtTitle3.Text <> "" Then s = s & "and j.estimator like " & DbQuote(Str, "%" & Trim(txtTitle3.Text) & "%") & vbCrLf
                s = s & "group by po,Vendor,vendordesc,pojob,pojobdesc,postarted,pocompleted" & vbCrLf
                s = s & "having sum(lineremainingpretax+lineremainingtax) <> 0" & vbCrLf
                s = s & "order by 1"
            Else
                If txtJob.Text = "" Or App.Options(JobBasedPONumbers) Then
                    'use tl with job based po numbers
                    s = ""
                    s = s & "SELECT sub" & vbCrLf
                    s = s & "      ,svendor" & vbCrLf
                    s = s & "      ,sdesc" & vbCrLf
                    s = s & "      ,jdesc" & vbCrLf
                    s = s & "      ,sactsd" & vbCrLf
                    s = s & "      ,sactcd" & vbCrLf
                    s = s & "      ,samt OrigAmt" & vbCrLf
                    s = s & "      ,sapprco         ChangeAmt" & vbCrLf
                    s = s & "      ,samtinv              Inv" & vbCrLf
                    s = s & "      ,samt+sapprco-samtinv Rem" & vbCrLf
                    s = s & "      ,sjob" & vbCrLf
                    s = s & "      ,jstatus" & vbCrLf
                    s = s & "  FROM master_jcm_record_12" & vbCrLf
                    s = s & "       LEFT OUTER JOIN master_jcm_record_1_1 ON(job=sjob)" & vbCrLf
                    s = s & " WHERE sclosed=0" & vbCrLf
                    s = s & "   AND samt+sapprco-samtinv<>0" & vbCrLf
                    If AccountingDB = dbTimberlinePVdata Then
                        If Val(mVendor) = 0 Then
                            s = s & "   AND svendor=" & DbQuote(Str, mVendor) & vbCrLf
                        Else
                            s = s & "   AND svendor=" & DbQuote(Str, FormatTSField(10, mVendor)) & vbCrLf
                        End If
                        If txtJob.Text <> "" Then
                            s = s & "   AND sub like " & DbQuote(Str, "%" & Trim(txtJob.Text) & "%") & vbCrLf
                        End If
                    Else
                        s = s & "   AND svendor=" & DbQuote(Str, mVendor) & vbCrLf
                        If txtJob.Text <> "" Then
                            s = s & "   AND sub like " & DbQuote(Str, "%" & Trim(txtJob.Text) & "%") & vbCrLf
                        End If
                    End If
                    If optCompleted(1).Value Then
                        s = s & "   AND not sactcd is null" & vbCrLf
                    End If
                    If optCompleted(2).Value Then
                        s = s & "   AND sactcd is null" & vbCrLf
                    End If
                    If txtDescription.Text <> "" Then s = s & "   AND jdesc like " & DbQuote(Str, "%" & Trim(txtDescription.Text) & "%") & vbCrLf
                    If txtPONumber.Text <> "" Then s = s & "   and sub like " & DbQuote(Str, "%" & Trim(txtPONumber.Text) & "%") & vbCrLf
                    If txtTitle1.Text <> "" Then s = s & "   AND jtitl1 like " & DbQuote(Str, "%" & Trim(txtTitle1.Text) & "%") & vbCrLf
                    If txtTitle2.Text <> "" Then s = s & "   AND jtitl2 like " & DbQuote(Str, "%" & Trim(txtTitle2.Text) & "%") & vbCrLf
                    If txtTitle3.Text <> "" Then s = s & "   AND jtitl3 like " & DbQuote(Str, "%" & Trim(txtTitle3.Text) & "%") & vbCrLf
                    If txtTitle4.Text <> "" Then s = s & "   AND jtitl4 like " & DbQuote(Str, "%" & Trim(txtTitle4.Text) & "%") & vbCrLf
                    s = s & "ORDER BY 1" & vbCrLf
                Else
                    'use tl with job on po details
                    s = ""
                    s = s & "SELECT DISTINCT" & vbCrLf
                    s = s & "       sub" & vbCrLf
                    s = s & "      ,svendor" & vbCrLf
                    s = s & "      ,sdesc" & vbCrLf
                    s = s & "      ,jdesc" & vbCrLf
                    s = s & "      ,sactsd" & vbCrLf
                    s = s & "      ,sactcd" & vbCrLf
                    s = s & "      ,samt OrigAmt" & vbCrLf
                    s = s & "      ,sapprco         ChangeAmt" & vbCrLf
                    s = s & "      ,samtinv              Inv" & vbCrLf
                    s = s & "      ,samt+sapprco-samtinv Rem" & vbCrLf
                    s = s & "      ,sjob" & vbCrLf
                    s = s & "      ,jstatus" & vbCrLf
                    s = s & "  FROM master_jcm_record_12" & vbCrLf
                    s = s & "       LEFT OUTER JOIN master_jcm_record_1_1 ON(job=sjob)" & vbCrLf
                    s = s & "      ,master_jcm_record_13" & vbCrLf
                    s = s & " WHERE sub=isub" & vbCrLf
                    s = s & "   AND sclosed=0" & vbCrLf
                    s = s & "   AND samt+sapprco-samtinv<>0" & vbCrLf
                    If optCompleted(1).Value Then s = s & "   AND not sactcd is null" & vbCrLf
                    If optCompleted(2).Value Then s = s & "   AND sactcd is null" & vbCrLf
                    If txtDescription.Text <> "" Then s = s & "   AND jdesc like " & DbQuote(Str, "%" & Trim(txtDescription.Text) & "%") & vbCrLf
                    If txtPONumber.Text <> "" Then s = s & "   and sub like " & DbQuote(Str, "%" & Trim(txtPONumber.Text) & "%") & vbCrLf
                    If txtTitle1.Text <> "" Then s = s & "   AND jtitl1 like " & DbQuote(Str, "%" & Trim(txtTitle1.Text) & "%") & vbCrLf
                    If txtTitle2.Text <> "" Then s = s & "   AND jtitl2 like " & DbQuote(Str, "%" & Trim(txtTitle2.Text) & "%") & vbCrLf
                    If txtTitle3.Text <> "" Then s = s & "   AND jtitl3 like " & DbQuote(Str, "%" & Trim(txtTitle3.Text) & "%") & vbCrLf
                    If txtTitle4.Text <> "" Then s = s & "   AND jtitl4 like " & DbQuote(Str, "%" & Trim(txtTitle4.Text) & "%") & vbCrLf
                    If AccountingDB = dbTimberlinePVdata Then
                        If txtJob.Text <> "" Then s = s & "   AND ijob = " & DbQuote(Str, FormatTSField(10, Trim(txtJob.Text), True)) & vbCrLf
                        If Not App.Options(AllowCrossPayingPOs) Then s = s & "   AND svendor=" & DbQuote(Str, FormatTSField(10, mVendor)) & vbCrLf
                    Else
                        If txtJob.Text <> "" Then s = s & "   AND ijob like " & DbQuote(Str, "%" & Trim(txtJob.Text) & "%") & vbCrLf
                        If Not App.Options(AllowCrossPayingPOs) Then s = s & "   AND svendor=" & DbQuote(Str, mVendor) & vbCrLf
                    End If
                    
                    s = s & "ORDER BY 1" & vbCrLf
                End If
            End If
            t = Timer()
            Set rs = HFApp.SqlExec(s, AccountingDB)
            mDebugStr = mDebugStr & "resolve query: " & Format(Timer() - t, "0.00") & " seconds" & vbCrLf
            t = Timer()
            r = 0
            While Not rs.EOF
                r = r + 1
                .AddItem ""
                .TextMatrix(r, .ColIndex("Commitment")) = Trim("" & rs("sub"))
                .TextMatrix(r, .ColIndex("DisplayCommitment")) = Trim("" & rs("sub"))
                .TextMatrix(r, .ColIndex("Vendor")) = Trim("" & rs("svendor"))
                .TextMatrix(r, .ColIndex("PODesc")) = Trim("" & rs("sdesc"))
                .TextMatrix(r, .ColIndex("JobDesc")) = Trim("" & rs("jdesc"))
                .TextMatrix(r, .ColIndex("DateStarted")) = Trim("" & rs("sactsd"))
                .TextMatrix(r, .ColIndex("DateCompleted")) = Trim("" & rs("sactcd"))
                .TextMatrix(r, .ColIndex("Job")) = Trim("" & rs("sjob"))
                .TextMatrix(r, .ColIndex("OriginalAmt")) = Trim("" & rs("OrigAmt"))
                .TextMatrix(r, .ColIndex("ChangeAmt")) = Trim("" & rs("ChangeAmt"))
                .TextMatrix(r, .ColIndex("Amount")) = Val("" & rs("OrigAmt")) + Val("" & rs("ChangeAmt"))
                .TextMatrix(r, .ColIndex("Invoiced")) = Trim("" & rs("Inv"))
                .TextMatrix(r, .ColIndex("Remaining")) = Trim("" & rs("Rem"))
                .Cell(flexcpChecked, r, 0) = flexUnchecked
                .IsSubtotal(r) = True
                .RowOutlineLevel(r) = 0
                
                r = r + 1
                .AddItem vbTab & "(dummy)"
                .IsSubtotal(r) = True
                .RowOutlineLevel(r) = 1
                
                rs.MoveNext
            Wend
            mDebugStr = mDebugStr & "fetch data:      " & Format(Timer() - t, "0.00") & " seconds" & vbCrLf
            .Outline 0
            
        Else
            'load po items
            Call .RemoveItem(.GetNodeRow(parentRow, flexNTFirstChild))
            .RowData(parentRow) = "loaded"
            
            
            If TimberlineAccounting Then
                s = ""
                s = s & "SELECT isub" & vbCrLf
                s = s & "      ,item" & vbCrLf
                s = s & "      ,idesc" & vbCrLf
                s = s & "      ,iamt                 OrigAmt" & vbCrLf
                s = s & "      ,iappcoa              ChangeAmt" & vbCrLf
                s = s & "      ,iamtinv              Inv" & vbCrLf
                s = s & "      ,iamt+iappcoa-iamtinv Rem" & vbCrLf
                s = s & "      ,ijob" & vbCrLf
                s = s & "      ,iextra" & vbCrLf
                s = s & "      ,iphase" & vbCrLf
                s = s & "      ,icat" & vbCrLf
                s = s & "  FROM master_jcm_record_12" & vbCrLf
                s = s & "      ,master_jcm_record_13" & vbCrLf
                s = s & " WHERE sub=isub" & vbCrLf
                If AccountingDB = dbTimberlinePVdata Then
                
                    If Val(mVendor) = 0 Then
                        s = s & "   AND svendor=" & DbQuote(Str, .TextMatrix(parentRow, .ColIndex("vendor"))) & vbCrLf
                    Else
                        s = s & "   AND svendor=" & DbQuote(Str, FormatTSField(10, .TextMatrix(parentRow, .ColIndex("vendor")))) & vbCrLf
                    End If
                    
                    s = s & "   AND sub=" & DbQuote(Str, FormatTSField(12, .TextMatrix(parentRow, .ColIndex("Commitment")))) & vbCrLf
                Else
                    s = s & "   AND svendor=" & DbQuote(Str, .TextMatrix(parentRow, .ColIndex("vendor"))) & vbCrLf
                    s = s & "   AND sub=" & DbQuote(Str, .TextMatrix(parentRow, .ColIndex("Commitment"))) & vbCrLf
                End If
                s = s & "ORDER BY 2" & vbCrLf
            Else
                s = ""
                s = s & "SELECT po isub" & vbCrLf
                s = s & "      ,line item" & vbCrLf
                s = s & "      ,linedesc idesc" & vbCrLf
                s = s & "      ,linepretax+linetax OrigAmt" & vbCrLf
                s = s & "      ,linechangespretax+linechangestax ChangeAmt" & vbCrLf
                s = s & "      ,lineinvoicedpretax+lineinvoicedtax Inv" & vbCrLf
                s = s & "      ,lineremainingpretax+lineremainingtax Rem" & vbCrLf
                s = s & "      ,linejob ijob" & vbCrLf
                s = s & "      ,lineextra iextra" & vbCrLf
                s = s & "      ,linecostcode iphase" & vbCrLf
                s = s & "      ,linecategory icat" & vbCrLf
                s = s & "from jcpodetails" & vbCrLf
                s = s & "where vendor=" & DbQuote(Str, .TextMatrix(parentRow, .ColIndex("vendor"))) & vbCrLf
                s = s & "  AND po=" & DbQuote(Str, .TextMatrix(parentRow, .ColIndex("Commitment"))) & vbCrLf
                s = s & "ORDER BY 2" & vbCrLf
            End If
            t = Timer()
            Set rs = HFApp.SqlExec(s, AccountingDB)
            mDebugStr = mDebugStr & "resolve query: " & Format(Timer() - t, "0.00") & " seconds" & vbCrLf
            t = Timer()
            r = parentRow
            While Not rs.EOF
                r = r + 1
                .AddItem "", r
                .TextMatrix(r, .ColIndex("Commitment")) = Trim("" & rs("isub"))
                .TextMatrix(r, .ColIndex("Item")) = Trim("" & rs("item"))
                .TextMatrix(r, .ColIndex("DisplayItem")) = Trim("" & rs("Item"))
                .TextMatrix(r, .ColIndex("PODesc")) = Trim("" & rs("idesc"))
                .TextMatrix(r, .ColIndex("OriginalAmt")) = Trim("" & rs("OrigAmt"))
                .TextMatrix(r, .ColIndex("ChangeAmt")) = Trim("" & rs("ChangeAmt"))
                .TextMatrix(r, .ColIndex("Amount")) = Val("" & rs("OrigAmt")) + Val("" & rs("ChangeAmt"))
                .TextMatrix(r, .ColIndex("Invoiced")) = Trim("" & rs("inv"))
                .TextMatrix(r, .ColIndex("Remaining")) = Trim("" & rs("rem"))
                .TextMatrix(r, .ColIndex("Job")) = Trim("" & rs("ijob"))
                .TextMatrix(r, .ColIndex("Extra")) = Trim("" & rs("iextra"))
                .TextMatrix(r, .ColIndex("Phase")) = Trim("" & rs("iphase"))
                .TextMatrix(r, .ColIndex("Category")) = Trim("" & rs("icat"))

                If .ValueMatrix(r, .ColIndex("Remaining")) = 0 Then .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbGrayText


                .Cell(flexcpChecked, r, 0) = .Cell(flexcpChecked, parentRow, 0)
                .IsSubtotal(r) = True
                .RowOutlineLevel(r) = 1
                rs.MoveNext
            Wend
            mDebugStr = mDebugStr & "fetch data: " & Format(Timer() - t, "0.00") & " seconds" & vbCrLf
            
        End If
        
        .ColWidth(0) = 555
        Call .AutoSize(1, .Cols - 1)
        
        .Redraw = flexRDBuffered
        Screen.MousePointer = vbDefault
        
    End With
Exit Sub
eh: Call ErrHandler(SRCFILE & "LoadData")
End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim n As VSFlexNode
    
    Dim parentRow As Long
    Dim childRow As Long
    
    With gData
        .Redraw = flexRDNone
        If .RowOutlineLevel(Row) = 0 Then
            parentRow = Row
            If .Cell(flexcpChecked, parentRow, 0) = flexTSChecked Then .Cell(flexcpChecked, parentRow, 0) = flexChecked
            If .Cell(flexcpChecked, parentRow, 0) = flexTSUnchecked Then .Cell(flexcpChecked, parentRow, 0) = flexUnchecked
            childRow = .GetNodeRow(parentRow, flexNTFirstChild)
            While childRow > 0
                .Cell(flexcpChecked, childRow, 0) = .Cell(flexcpChecked, parentRow, 0)
                childRow = .GetNodeRow(childRow, flexNTNextSibling)
            Wend
        Else
            childRow = Row
            parentRow = .GetNodeRow(childRow, flexNTParent)
            .Cell(flexcpChecked, parentRow, 0) = .Cell(flexcpChecked, childRow, 0)
            childRow = .GetNodeRow(parentRow, flexNTFirstChild)
            While childRow > 0
                If .Cell(flexcpChecked, childRow, 0) <> .Cell(flexcpChecked, parentRow, 0) Then
                    .Cell(flexcpChecked, parentRow, 0) = flexTSGrayed
                    .Redraw = flexRDBuffered
                    Call RecalcTotals
                    Exit Sub
                End If
                childRow = .GetNodeRow(childRow, flexNTNextSibling)
            Wend
        End If
        .Redraw = flexRDBuffered
    End With
    Call RecalcTotals
End Sub

Private Sub gData_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
    If Row < 1 Then Exit Sub
    If State = flexOutlineCollapsed Then Exit Sub
    If gData.RowData(Row) = "loaded" Then Exit Sub
    
    gData.Row = Row
    Call LoadData(Row)
    
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> 0
End Sub

Private Sub gData_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = 0
End Sub

Private Sub gData_GotFocus()
    cmdNav(0).Default = True
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    With gData
        If .Row < 1 Then Exit Sub
        Select Case True
        
            Case KeyCode = vbKeyRight
                If .RowOutlineLevel(.Row) = 0 Then
                    If Not .GetNode(.Row).Expanded Then
                        .GetNode(.Row).Expanded = True
                        .GetNode(.GetNodeRow(.Row, flexNTFirstChild)).EnsureVisible
                        .Col = .Col - 1
                    End If
                End If

            Case KeyCode = vbKeyLeft
                If .RowOutlineLevel(.Row) = 0 Then
                    If .GetNode(.Row).Expanded Then
                        .GetNode(.Row).Expanded = False
                        .Col = .Col + 1
                    End If
                Else
                    .Row = .GetNodeRow(.Row, flexNTParent)
                    .GetNode(.Row).EnsureVisible
                End If
                
            Case KeyCode = vbKeySpace And .Col <> 0
                .Cell(flexcpChecked, .Row, 0) = IIf(.Cell(flexcpChecked, .Row, 0) = flexUnchecked, flexChecked, flexUnchecked)
                Call gData_AfterEdit(.Row, 0)
                
                
        End Select
    End With
End Sub

Private Sub gData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call ShowColumnMenu(gData, False)
End Sub

Private Sub gData_RowColChange()
'    gData.Col = 3
End Sub

Private Sub RecalcTotals()
    Dim r As Long
    Dim d As Double
    
    With gData
        r = 1
        While IsBetween(r, 1, .Rows - 1)
            If .Cell(flexcpChecked, r, 0) = flexChecked Then
                d = d + .ValueMatrix(r, .ColIndex("Amount"))
                If .RowOutlineLevel(r) = 0 Then
                    r = .GetNodeRow(r, flexNTNextSibling)
                Else
                    r = r + 1
                End If
            Else
                r = r + 1
            End If
        Wend
    End With
'    d = mInvoice - d
    
    lblRemaining.Caption = Format(mInvoice, "$#,##0.00")
    lblRemaining.ForeColor = IIf(mInvoice < 0, vbRed, vbButtonText)
    
    lblSelected.Caption = Format(d, "$#,##0.00")
    lblSelected.ForeColor = IIf(d < 0, vbRed, vbButtonText)
End Sub

Private Sub SetCaptions()
    Me.Caption = App.Options(Caption_Commitment) & " List"
    
    lblCaption.Caption = "Select " & App.Options(Caption_Commitment) & "s and items from the list by marking the check box" & vbCrLf & _
                         "Search by entering values into the fields below."
    
    lblJob.Caption = App.Options(Caption_Job)
    lblPONumber.Caption = App.Options(Caption_Commitment)
    
    If HFApp.Options(AccountingSystem) = asTimberline Then
        lblTitle1.Caption = App.Options(Caption_JobTitle1)
        lblTitle2.Caption = App.Options(Caption_JobTitle2)
        lblTitle3.Caption = App.Options(Caption_JobTitle3)
        lblTitle4.Caption = App.Options(Caption_JobTitle4)
    Else
        lblTitle1.Caption = "Project Manager"
        lblTitle2.Caption = "Purchaser"
        lblTitle3.Caption = "Estimator"
        lblTitle4.Visible = False
        txtTitle4.Visible = False
        Label4(0).Height = 495
    End If
    
    With gData
        .TextMatrix(0, .ColIndex("DisplayCommitment")) = App.Options(Caption_Commitment)
        .TextMatrix(0, .ColIndex("Vendor")) = App.Options(Caption_Vendor)
        .TextMatrix(0, .ColIndex("Job")) = App.Options(Caption_Job)
        .TextMatrix(0, .ColIndex("Extra")) = App.Options(Caption_Extra)
        .TextMatrix(0, .ColIndex("Phase")) = App.Options(Caption_Phase)
        .TextMatrix(0, .ColIndex("Category")) = App.Options(Caption_Category)
    End With
End Sub

Private Sub ShowColumnMenu(Grid As VSFlexGrid, Optional Sortable As Boolean = True)
On Error GoTo eh
    Dim i As Long
    Dim j As Long
    
    'save this stuff for menu click
    Set MouseGrid = Grid
    MouseCol = Grid.MouseCol
    
    'set these
    mnuGridSub(mcGRID_ASC).Enabled = Sortable
    mnuGridSub(mcGRID_DESC).Enabled = Sortable
    mnuGridSub(mcGRID_HIDE).Enabled = MouseCol >= 0
    
    'load Grid column names
    mnuColumnsSub(0).Visible = True
    For i = mnuColumnsSub.UBound To 1 Step -1
        Unload mnuColumnsSub(i)
    Next
    For i = 0 To MouseGrid.Cols - 1
        If MouseGrid.ColHidden(i) And MouseGrid.TextMatrix(0, i) <> "" Then
            j = j + 1
            Load mnuColumnsSub(j)
            mnuColumnsSub(j).tag = MouseGrid.ColKey(i)
            mnuColumnsSub(j).Caption = MouseGrid.TextMatrix(0, i)
            mnuColumnsSub(j).Visible = True
            mnuColumnsSub(j).Enabled = True
        End If
    Next
    If j = 0 Then mnuColumnsSub(0).Caption = "(none available)"
    mnuColumnsSub(0).Visible = j = 0
    mnuColumnsSub(0).Enabled = False
    
    'show menu
    PopupMenu mnuGrid

    Exit Sub
eh: Call ErrHandler(SRCFILE & "ShowColumnMenu")
End Sub

Private Sub lblDescription_Click()
On Error Resume Next
    If FPickList.Choose(HFApp.Databases(AccountingDB), App.Options(Caption_Vendor), SelectJob, txtJob.Text) Then
        txtDescription.Text = FPickList.SelectedItem(2)
    End If
End Sub

Private Sub mnuColumnsSub_Click(Index As Integer)
    MouseGrid.ColHidden(MouseGrid.ColIndex(mnuColumnsSub(Index).tag)) = False
    MouseGrid.ColPosition(MouseGrid.ColIndex(mnuColumnsSub(Index).tag)) = IIf(MouseCol < 0, MouseGrid.Cols - 1, MouseCol)
    
    gData.ColPosition(gData.ColIndex("Add")) = 0
    MouseGrid.ColHidden(0) = False
    
End Sub
Private Sub mnuGridSub_Click(Index As Integer)
    Dim i As Long
    Select Case Index
        Case mcGRID_ASC
            MouseGrid.Col = MouseCol
            MouseGrid.Sort = flexSortGenericAscending
            
        Case mcGRID_DESC
            MouseGrid.Col = MouseCol
            MouseGrid.Sort = flexSortGenericDescending
            
        Case mcGRID_HIDE
            For i = 0 To MouseGrid.Cols - 1
                If MouseGrid.ColHidden(i) = False Then
                    MouseGrid.ColHidden(MouseCol) = True
                    Exit Sub
                End If
            Next
    End Select
End Sub

Private Sub optCompleted_Click(Index As Integer)
    Call LoadData(-1)
End Sub

Private Sub txtJob_Change()
    Dim i As Long
    i = txtJob.SelStart
    txtJob.Text = UCase(txtJob.Text)
    txtJob.SelStart = i
End Sub

Private Sub txtJob_GotFocus()
    cmdSearch.Default = True
End Sub
Private Sub txtDescription_GotFocus()
    cmdSearch.Default = True
End Sub
Private Sub txtTitle1_GotFocus()
    cmdSearch.Default = True
End Sub
Private Sub txtTitle2_GotFocus()
    cmdSearch.Default = True
End Sub
Private Sub txtTitle3_GotFocus()
    cmdSearch.Default = True
End Sub
Private Sub txtTitle4_GotFocus()
    cmdSearch.Default = True
End Sub

Private Sub lblJob_Click()
On Error Resume Next
    If FPickList.Choose(HFApp.Databases(AccountingDB), App.Options(Caption_Vendor), SelectJob, txtJob.Text) Then
        txtJob.Text = FPickList.SelectedItem(1)
    End If
End Sub



