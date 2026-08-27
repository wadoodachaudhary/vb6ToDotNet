VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FExportEstimates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimate Export Wizard"
   ClientHeight    =   5220
   ClientLeft      =   6045
   ClientTop       =   2385
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "FExportEstimates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6495
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4635
      Width           =   6495
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4140
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5340
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   60
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
      Index           =   1
      Left            =   7800
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CheckBox chkShowBatchReport 
         Caption         =   "Show summary report when finished."
         Height          =   285
         Left            =   1140
         TabIndex        =   12
         Top             =   2700
         Width           =   4455
      End
      Begin VB.CheckBox chkCloseWhenFinished 
         Caption         =   "Close this dialog when all estimates are completed."
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Top             =   2940
         Width           =   4455
      End
      Begin VB.TextBox txtEstPath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   300
         Width           =   4215
      End
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         Picture         =   "FExportEstimates.frx":000C
         ForeColor       =   0
         BarPicture      =   "FExportEstimates.frx":0028
         ShowText        =   -1  'True
         TextAlignX      =   0
         Text            =   "  Preparing Estimate..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Segments        =   -1  'True
         XpStyle         =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   300
         Picture         =   "FExportEstimates.frx":0044
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Options:"
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   2460
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblSaving 
         Caption         =   "Saving:"
         Height          =   255
         Left            =   1140
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "The selected estimates will be created in:"
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   7
         Top             =   0
         Width           =   3915
      End
      Begin VB.Label lblAssemblyCount 
         AutoSize        =   -1  'True
         Caption         =   "Assembly 3 of 17"
         Height          =   195
         Left            =   1380
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblCustomerDesc 
         AutoSize        =   -1  'True
         Caption         =   "Bob and Shirley MacClaney"
         Height          =   195
         Left            =   1380
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label lblItemDesc 
         AutoSize        =   -1  'True
         Caption         =   "Regency C - Extended Garage"
         Height          =   195
         Left            =   1380
         TabIndex        =   4
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3360
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   3120
         Left            =   1140
         TabIndex        =   11
         Top             =   240
         Width           =   5175
         _cx             =   9128
         _cy             =   5503
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FExportEstimates.frx":090E
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
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
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
         BackColorFrozen =   -2147483624
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
         Picture         =   "FExportEstimates.frx":0AD1
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Select which estimates you would like to produce."
         Height          =   255
         Index           =   5
         Left            =   1140
         TabIndex        =   10
         Top             =   0
         Width           =   3915
      End
   End
   Begin HFEst.WizHead WizHead 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1588
      Caption         =   "Export Timberline Estimates"
      Description     =   "The export wizard will write customer estimates to your estimates folder"
      Icon            =   "FExportEstimates.frx":139B
   End
End
Attribute VB_Name = "FExportEstimates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Rewrite As Boolean

Const SRCFILE = "FExportEstimates::"



Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData, , , Rewrite)
    chkCloseWhenFinished.Value = IIf(HFApp.Options(CloseAssemblyExportDialog), vbChecked, vbUnchecked)
    chkShowBatchReport.Value = IIf(HFApp.Options(ShowAssemblyExportReport), vbChecked, vbUnchecked)
    txtEstPath.Text = HFApp.Options(CustomerPeeRootFolder)
    CurrentFrame = 0
End Sub



Private Function SaveData() As Boolean
    Dim i     As Long
    Dim j     As Long
    Dim Rows  As Long
    Dim Batch As Long
    Dim f     As New FRptViewer

    If HFApp.Options(CustomerPeeRootFolder) = "" Then
        If vbYes = MsgBox("Before using this feature you must specify" & vbCrLf & "the root folder and naming convention." & vbCrLf & vbCrLf & "Would you like to do that now?", vbYesNo + vbQuestion) Then
            FOptions.Show vbModal
        End If
    End If

    txtEstPath.Text = HFApp.Options(CustomerPeeRootFolder)
    If HFApp.Options(CustomerPeeRootFolder) = "" Then Exit Function

    With gData
    
        cmdNav(0).Enabled = False
        cmdNav(1).Enabled = False
        cmdNav(2).Enabled = False
        lblSaving.Visible = True
        ProgressBar.Visible = True
        lblAssemblyCount.Visible = True
        lblCustomerDesc.Visible = True
        lblItemDesc.Visible = True
        
        'get row count
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then Rows = Rows + 1
        Next

        If Not Rewrite Then
            Call HFApp.SqlExec("insert into tblCustomerEstBatches(UStmp,TStmp) VALUES(" & DbQuote(Str, HFApp.LoginID) & ",GETDATE())")
            Batch = HFApp.SqlIdentity("tblCustomerEstBatches")
        End If
        
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
                .Row = i
                j = j + 1
                lblAssemblyCount.Caption = "Customer " & j & " of " & Rows & "..."
                lblCustomerDesc.Caption = .TextMatrix(i, .ColIndex("CustomerDesc"))
                lblItemDesc.Caption = .TextMatrix(i, .ColIndex("ModelDesc"))
                
                If Rewrite Then
                    Call WriteCustomerPee(.TextMatrix(i, .ColIndex("Customer")), .ValueMatrix(i, .ColIndex("EstimateIndex")), ProgressBar)
                Else
                    Call WriteCustomerEstimate(Batch, .TextMatrix(i, .ColIndex("Customer")), ProgressBar)
                End If
            End If
        Next
        
        
        On Error Resume Next
        ProgressBar.Text = ""
        ProgressBar.Value = 100
        lblAssemblyCount.Caption = "Finished"
        lblCustomerDesc.Visible = False
        lblItemDesc.Visible = False

        
        cmdNav(2).Caption = "&Close"
        cmdNav(2).Enabled = True
        cmdNav(2).SetFocus
        If chkShowBatchReport.Value = vbChecked Then Call f.ShowReport(HFApp.SystemFolder & "System\Reports\Estimating\exportestbatch.rpt", True, False, True, "BatchNo", Batch)
        If chkCloseWhenFinished.Value = vbChecked Then Unload Me
        
    End With
    SaveData = True

End Function

Private Sub LoadData()
    Dim rs As Recordset
    Dim r As Long
    Dim c As Long
    Dim s As String

    With gData
        .Redraw = flexRDNone
        .Rows = 1
        r = 0

        .ColHidden(.ColIndex("EstimateIndex")) = Not Rewrite
        If Rewrite Then
            s = ""
            s = s & "SELECT 0 Selected" & vbCrLf
            s = s & "      ,e.EstimateIndex" & vbCrLf
            s = s & "      ,c.Community" & vbCrLf
            s = s & "      ,l.Description CommunityDesc" & vbCrLf
            s = s & "      ,Job" & vbCrLf
            s = s & "      ,j.Description JobDesc" & vbCrLf
            s = s & "      ,c.Customer_No Customer" & vbCrLf
            s = s & "      ,c.Description CustomerDesc" & vbCrLf
            s = s & "      ,c.Model" & vbCrLf
            s = s & "      ,m.Description ModelDesc " & vbCrLf
            s = s & "      ,c.Series" & vbCrLf
            s = s & "  FROM tblCustomers c " & vbCrLf
            s = s & "       JOIN tblCustomerEstimates e ON(isnull(c.Customer_No,'')=isnull(e.Customer_No,''))" & vbCrLf
            s = s & "       LEFT JOIN tblLocality l ON(isnull(c.Community,'')=isnull(l.Area,''))" & vbCrLf
            s = s & "       LEFT JOIN tblJobs j ON(isnull(c.Job_No,'')=isnull(j.Job_No,''))" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblModels m ON (isnull(c.Model,'')=isnull(m.Model,'') AND isnull(c.Community,'')=isnull(m.area,'') AND isnull(c.Series,'')=isnull(m.Series,'') AND isnull(c.Phase,'')=isnull(m.CommunityPhase,''))" & vbCrLf
        Else
            s = ""
            s = s & "SELECT 0 Selected" & vbCrLf
            s = s & "      ,0 EstimateIndex" & vbCrLf
            s = s & "      ,c.Community" & vbCrLf
            s = s & "      ,a.Description CommunityDesc" & vbCrLf
            s = s & "      ,c.Job_No Job" & vbCrLf
            s = s & "      ,j.Description JobDesc" & vbCrLf
            s = s & "      ,c.Customer_No Customer" & vbCrLf
            s = s & "      ,c.Description CustomerDesc" & vbCrLf
            s = s & "      ,c.Model" & vbCrLf
            s = s & "      ,m.Description ModelDesc" & vbCrLf
            s = s & "      ,c.Series" & vbCrLf
            s = s & "  FROM tblCustomers c " & vbCrLf
            s = s & "       LEFT OUTER JOIN tblJobs j ON (isnull(c.Job_No,'')=isnull(j.Job_No,''))" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblModels m ON (isnull(c.Model,'')=isnull(m.Model,'') AND isnull(c.Community,'')=isnull(m.area,'') AND isnull(c.Series,'')=isnull(m.Series,'') AND isnull(c.Phase,'')=isnull(m.CommunityPhase,''))" & vbCrLf
            s = s & "      ,system_setup s" & vbCrLf
            s = s & "      ,tblLocality a" & vbCrLf
            s = s & " WHERE isnull(c.community,'')=isnull(a.area,'')" & vbCrLf
            s = s & "  AND ((c.purchased=1  AND c.cancelled=0  AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1)))"
            s = s & "    OR (c.home_selection <>'PreSale' AND c.sold_to_Customer=''))"
            s = s & "   AND c.Estimateindex=0" & vbCrLf
            s = s & "UNION" & vbCrLf
            s = s & "SELECT 0 Selected" & vbCrLf
            s = s & "      ,0 EstimateIndex" & vbCrLf
            s = s & "      ,c.Community" & vbCrLf
            s = s & "      ,l.Description CommunityDesc" & vbCrLf
            s = s & "      ,c.Job_No Job" & vbCrLf
            s = s & "      ,j.Description JobDesc" & vbCrLf
            s = s & "      ,c.Customer_No Customer" & vbCrLf
            s = s & "      ,c.Description CustomerDesc" & vbCrLf
            s = s & "      ,c.Model" & vbCrLf
            s = s & "      ,m.Description ModelDesc" & vbCrLf
            s = s & "      ,c.Series" & vbCrLf
            s = s & "  FROM tblCustomers c " & vbCrLf
            s = s & "       LEFT OUTER JOIN tblJobs j ON (isnull(c.Job_No,'')=isnull(j.Job_No,''))" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblModels m ON (isnull(c.Model,'')=isnull(m.Model,'') AND isnull(c.Community,'')=isnull(m.area,'') AND isnull(c.Series,'')=isnull(m.Series,'') AND isnull(c.Phase,'')=isnull(m.CommunityPhase,''))" & vbCrLf
            s = s & "      ,tblLocality  l" & vbCrLf
            s = s & "      ,TBLSCHEDULEB d" & vbCrLf
            s = s & "      ,system_setup s" & vbCrLf
            s = s & "WHERE c.Customer_No=d.Customer_No" & vbCrLf
            s = s & "  AND c.Community=l.Area" & vbCrLf
            s = s & "  AND d.estimateindex=0" & vbCrLf
            s = s & "  AND ((c.purchased=1  AND c.cancelled=0  AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1)))"
            s = s & "    OR (c.home_selection <>'PreSale' AND c.sold_to_Customer=''))"
            s = s & "UNION" & vbCrLf
            s = s & "SELECT 0 Selected" & vbCrLf
            s = s & "      ,0 EstimateIndex" & vbCrLf
            s = s & "      ,c.Community" & vbCrLf
            s = s & "      ,l.Description CommunityDesc" & vbCrLf
            s = s & "      ,c.Job_No Job" & vbCrLf
            s = s & "      ,j.Description JobDesc" & vbCrLf
            s = s & "      ,c.Customer_No Customer" & vbCrLf
            s = s & "      ,c.Description CustomerDesc" & vbCrLf
            s = s & "      ,c.Model" & vbCrLf
            s = s & "      ,m.Description ModelDesc" & vbCrLf
            s = s & "      ,c.Series" & vbCrLf
            s = s & "  FROM tblCustomers c " & vbCrLf
            s = s & "       LEFT OUTER JOIN tblJobs j ON (isnull(c.Job_No,'')=isnull(j.Job_No,''))" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblModels m ON (isnull(c.Model,'')=isnull(m.Model,'') AND isnull(c.Community,'')=isnull(m.area,'') AND isnull(c.Series,'')=isnull(m.Series,'') AND isnull(c.Phase,'')=isnull(m.CommunityPhase,''))" & vbCrLf
            s = s & "      ,tblLocality  l" & vbCrLf
            s = s & "      ,ChangeOrderMaster co" & vbCrLf
            s = s & "      ,ChangeOrderDetails d" & vbCrLf
            s = s & "      ,system_setup s" & vbCrLf
            s = s & " WHERE c.Community=l.Area" & vbCrLf
            s = s & "   AND c.Customer_No=co.Customer_No" & vbCrLf
            s = s & "   AND co.Customer_No=d.Customer_No" & vbCrLf
            s = s & "   AND co.Change_Order_No=d.Change_Order_No" & vbCrLf
            s = s & "   AND d.estimateindex=0" & vbCrLf
            s = s & "   AND co.bldr_approved=1" & vbCrLf
            s = s & "  AND ((c.purchased=1  AND c.cancelled=0  AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1)))"
            s = s & "    OR (c.home_selection <>'PreSale' AND c.sold_to_Customer=''))"
            s = s & "UNION" & vbCrLf
            s = s & "SELECT 0 Selected" & vbCrLf
            s = s & "      ,0 EstimateIndex" & vbCrLf
            s = s & "      ,c.Community" & vbCrLf
            s = s & "      ,l.Description CommunityDesc" & vbCrLf
            s = s & "      ,c.Job_No Job" & vbCrLf
            s = s & "      ,j.Description JobDesc" & vbCrLf
            s = s & "      ,c.Customer_No Customer" & vbCrLf
            s = s & "      ,c.Description CustomerDesc" & vbCrLf
            s = s & "      ,c.Model" & vbCrLf
            s = s & "      ,m.Description ModelDesc" & vbCrLf
            s = s & "      ,c.Series" & vbCrLf
            s = s & "  FROM tblCustomers c " & vbCrLf
            s = s & "       LEFT OUTER JOIN tblJobs j ON (isnull(c.Job_No,'')=isnull(j.Job_No,''))" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblModels m ON (isnull(c.Model,'')=isnull(m.Model,'') AND isnull(c.Community,'')=isnull(m.area,'') AND isnull(c.Series,'')=isnull(m.Series,'') AND isnull(c.Phase,'')=isnull(m.CommunityPhase,''))" & vbCrLf
            s = s & "      ,tblLocality  l" & vbCrLf
            s = s & "      ,DesignCenterMaster dc" & vbCrLf
            s = s & "      ,DesignCenterDetails d" & vbCrLf
            s = s & "      ,system_setup s" & vbCrLf
            s = s & " WHERE c.Community=l.Area" & vbCrLf
            s = s & "   AND c.Customer_No=dc.customer_No" & vbCrLf
            s = s & "   AND dc.customer_No=d.Customer_No" & vbCrLf
            s = s & "   AND dc.Change_Order_No=d.Change_Order_No" & vbCrLf
            s = s & "   AND d.estimateindex=0" & vbCrLf
            s = s & "   AND dc.bldr_approved=1" & vbCrLf
            s = s & "  AND ((c.purchased=1  AND c.cancelled=0  AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1)))"
            s = s & "    OR (c.home_selection <>'PreSale' AND c.sold_to_Customer=''))"
            s = s & "UNION" & vbCrLf
            s = s & "SELECT 0 Selected" & vbCrLf
            s = s & "      ,0 EstimateIndex" & vbCrLf
            s = s & "      ,c.Community" & vbCrLf
            s = s & "      ,l.Description CommunityDesc" & vbCrLf
            s = s & "      ,c.Job_No Job" & vbCrLf
            s = s & "      ,j.Description JobDesc" & vbCrLf
            s = s & "      ,c.Customer_No Customer" & vbCrLf
            s = s & "      ,c.Description CustomerDesc" & vbCrLf
            s = s & "      ,c.Model" & vbCrLf
            s = s & "      ,m.Description ModelDesc" & vbCrLf
            s = s & "      ,c.Series" & vbCrLf
            s = s & "  FROM tblCustomers c " & vbCrLf
            s = s & "       LEFT OUTER JOIN tblJobs j ON (isnull(c.Job_No,'')=isnull(j.Job_No,''))" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblModels m ON (isnull(c.Model,'')=isnull(m.Model,'') AND isnull(c.Community,'')=isnull(m.area,'') AND isnull(c.Series,'')=isnull(m.Series,'') AND isnull(c.Phase,'')=isnull(m.CommunityPhase,''))" & vbCrLf
            s = s & "      ,tblLocality  l" & vbCrLf
            s = s & "      ,DeletedItemsMaster di" & vbCrLf
            s = s & "      ,tblDeletedItems    d " & vbCrLf
            s = s & "      ,system_setup s" & vbCrLf
            s = s & " WHERE c.Community = l.Area" & vbCrLf
            s = s & "   AND c.Customer_No = di.Customer_No" & vbCrLf
            s = s & "   AND di.Customer_No = d.Customer_No" & vbCrLf
            s = s & "   AND di.Change_Order_No = d.Change_Order_No" & vbCrLf
            s = s & "   AND d.EstimateIndex = 0" & vbCrLf
            s = s & "   AND di.Bldr_Approved = 1" & vbCrLf
            s = s & "  AND ((c.purchased=1  AND c.cancelled=0  AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1)))"
            s = s & "    OR (c.home_selection <>'PreSale' AND c.sold_to_Customer=''))"
        End If
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            For c = 0 To .Cols - 1
                .TextMatrix(r, c) = "" & rs(.ColKey(c))
            Next
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    HFApp.Options.Value(CloseAssemblyExportDialog) = chkCloseWhenFinished.Value = vbChecked
    HFApp.Options.Value(ShowAssemblyExportReport) = chkShowBatchReport.Value = vbChecked
    Call IniPutGrid(Me, gData, , Rewrite)
    Call IniPutForm(Me)
End Sub


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = gData.ColKey(Col) <> "Selected"
End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton And gData.MouseRow = 0 Then
        Cancel = True
        Call FMain.ShowColumnMenu(gData, False)
    End If
End Sub




Private Sub cmdNav_Click(Index As Integer)
'On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3:
            If SaveData Then
                If chkCloseWhenFinished Then
                    Unload Me
                Else
                    cmdNav(0).Enabled = True
                    cmdNav(0).Caption = "Close"
                    cmdNav(1).Enabled = False
                    cmdNav(2).Enabled = False
                    cmdNav(3).Enabled = False
                End If
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
        WizFrame(i).Move 120, 960
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
    
    
    
    Select Case RHS
        Case 0: Call LoadData
        Case 1:
    End Select
    
End Property


