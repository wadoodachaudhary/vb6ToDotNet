VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FInboxTBDAssignments 
   Caption         =   "TBD Purchase Orders"
   ClientHeight    =   6135
   ClientLeft      =   1350
   ClientTop       =   2070
   ClientWidth     =   8400
   Icon            =   "FInboxTBDAssignments.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8400
   Begin HFEst.WizHead WizHead 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   1588
      Caption         =   "Confirm Vendor TBD Assignments"
      Description     =   "Refresh costs on TBD POs for assigned Vendors"
      Icon            =   "FInboxTBDAssignments.frx":000C
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Index           =   1
      Left            =   7245
      TabIndex        =   2
      Top             =   5655
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Approve"
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FInboxTBDAssignments.frx":08E6
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
   End
End
Attribute VB_Name = "FInboxTBDAssignments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const SRCFILE = "FInboxTBDAssignments::"
Private mDirty   As Boolean

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'approve
            Call SaveData
            
        Case 1 'cancel
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    Call LoadData
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    gData.Move margin, WizHead.Height + margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - WizHead.Height - margin * 3 - cmdNav(0).Height
    cmdNav(0).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
    cmdNav(1).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
End Sub



Private Function SaveData() As String
On Error GoTo eh
    Dim r As Long
    Dim POs As String
    Dim s As String
    
    With gData
        For r = 1 To .Rows - 1
            If .Cell(flexcpChecked, r, .ColIndex("PONumber")) = flexChecked Then
                POs = POs & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("PONumber")))
            End If
        Next
        POs = Mid(POs, 2)
        If POs = "" Then Exit Function
    
        s = ""
        s = s & "update e set" & vbCrLf
        s = s & " PORate = isnull(nullif(dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Rate)" & vbCrLf
        s = s & ",POPretax = round(e.POQty * isnull(nullif(dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Rate),2)" & vbCrLf
        s = s & ",POTaxGroup = t.TaxGroup" & vbCrLf
        s = s & ",POJCTaxRate = t.JCRate" & vbCrLf
        s = s & ",PONJCTaxRate = t.NJCRate" & vbCrLf
        s = s & ",POJCTax  = t.JCRate/100  * round(e.POQty * isnull(nullif(dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Rate),2)" & vbCrLf
        s = s & ",PONJCTax = t.NJCRate/100 * round(e.POQty * isnull(nullif(dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Rate),2)" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "join poitems i on p.ponumber=i.ponumber and p.divisionid=i.divisionid" & vbCrLf
        s = s & "join estimateitems e on i.estitemid=e.estitemid" & vbCrLf
        s = s & "join estimateassemblies a on a.estassemblyid=e.estassemblyid" & vbCrLf
        s = s & "left join tblvendors tv on p.vendor=tv.vendor_id and p.divisionid=tv.divisionid" & vbCrLf
        s = s & "left join tblvendors av on p.BuildPro_AssignedVendor=av.vendor_id and p.divisionid=av.divisionid" & vbCrLf
        s = s & "left join tbljobs j on p.job=j.job_no and p.divisionid=j.divisionid" & vbCrLf
        s = s & "left join taxgroups t on t.taxgroup=dbo.Purch_GetDefaultTaxGroup(j.Job_No,j.Community,j.CommunityPhase,e.Model,e.Assembly,e.Phase,e.Item,p.buildpro_assignedvendor,e.JCCategory,e.DivisionID) " & vbCrLf
        s = s & "left join tblpoindex x on p.poindex=x.poindex and p.divisionid=x.divisionid" & vbCrLf
        s = s & "where p.divisionid =" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and p.ponumber in(" & POs & ")" & vbCrLf
        Call HFApp.SqlExec(s)
    
        s = ""
        s = s & "update i set" & vbCrLf
        s = s & " Rate = isnull(nullif(dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Rate)" & vbCrLf
        s = s & ",Pretax = round(e.POQty * isnull(nullif(dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Rate),2)" & vbCrLf
        s = s & ",TaxGroup = t.TaxGroup" & vbCrLf
        s = s & ",JCTaxRate = t.JCRate" & vbCrLf
        s = s & ",NJCTaxRate = t.NJCRate" & vbCrLf
        s = s & ",JCTax  = t.JCRate/100  * round(e.POQty * isnull(nullif(dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Rate),2)" & vbCrLf
        s = s & ",NJCTax = t.NJCRate/100 * round(e.POQty * isnull(nullif(dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Rate),2)" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "join poitems i on p.ponumber=i.ponumber and p.divisionid=i.divisionid" & vbCrLf
        s = s & "join estimateitems e on i.estitemid=e.estitemid" & vbCrLf
        s = s & "join estimateassemblies a on a.estassemblyid=e.estassemblyid" & vbCrLf
        s = s & "left join tblvendors tv on p.vendor=tv.vendor_id and p.divisionid=tv.divisionid" & vbCrLf
        s = s & "left join tblvendors av on p.BuildPro_AssignedVendor=av.vendor_id and p.divisionid=av.divisionid" & vbCrLf
        s = s & "left join tbljobs j on p.job=j.job_no and p.divisionid=j.divisionid" & vbCrLf
        s = s & "left join taxgroups t on t.taxgroup=dbo.Purch_GetDefaultTaxGroup(j.Job_No,j.Community,j.CommunityPhase,e.Model,e.Assembly,e.Phase,e.Item,p.buildpro_assignedvendor,e.JCCategory,e.DivisionID) " & vbCrLf
        s = s & "left join tblpoindex x on p.poindex=x.poindex and p.divisionid=x.divisionid" & vbCrLf
        s = s & "where p.divisionid =" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and p.ponumber in(" & POs & ")" & vbCrLf
        Call HFApp.SqlExec(s)
    
        s = ""
        s = s & "update p set" & vbCrLf
        s = s & " Vendor = BuildPro_AssignedVendor" & vbCrLf
        s = s & ",BuildPro_AssignedVendor = ''" & vbCrLf
        s = s & ",DateSentToBuildPro = null" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "where p.divisionid =" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and p.ponumber in(" & POs & ")" & vbCrLf
        Call HFApp.SqlExec(s)
        
        Call SendBuildProPOs("", POs)
        
        For r = .Rows - 1 To 1 Step -1
            If .Cell(flexcpChecked, r, .ColIndex("PONumber")) = flexChecked Then
                .RemoveItem r
            End If
        Next
    
    End With
    SaveData = ""

Exit Function
eh: Call errHandler(SRCFILE & "SaveData")
End Function

Private Sub LoadData()

Dim rs As Recordset
Dim r As Long
Dim s As String
    
    
    With gData
        .Redraw = flexRDNone
        .Rows = 1
        
        r = 0
        
        
        s = ""
        s = s & "select" & vbCrLf
        s = s & " p.PONumber" & vbCrLf
        s = s & ",p.Description PODesc" & vbCrLf
        s = s & ",av.Vendor_ID AssignedVendor" & vbCrLf
        s = s & ",av.Vendor_Name AssignedVendorName" & vbCrLf
        s = s & ",tv.Vendor_ID TBDVendor" & vbCrLf
        s = s & ",tv.Vendor_Name TBDVendorName" & vbCrLf
        s = s & ",p.Job" & vbCrLf
        s = s & ",j.Description JobDesc" & vbCrLf
        s = s & ",p.POIndex" & vbCrLf
        s = s & ",x.description POIndexDesc" & vbCrLf
        s = s & ",sum(i.Pretax) BudgetAmt" & vbCrLf
        s = s & ",sum(round(isnull(nullif(e.POQty * dbo.Purch_GetItemRate(0,0,j.Community,j.CommunityPhase,e.Assembly,e.model,a.optionid,e.phase,e.item,0,p.buildpro_assignedvendor,getdate(),a.divisionid),0),i.Pretax),2)) VendorAmt" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "left join poitems i on p.ponumber=i.ponumber and p.divisionid=i.divisionid" & vbCrLf
        s = s & "left join estimateitems e on i.estitemid=e.estitemid" & vbCrLf
        s = s & "left join estimateassemblies a on a.estassemblyid=e.estassemblyid" & vbCrLf
        s = s & "left join tblvendors tv on p.vendor=tv.vendor_id and p.divisionid=tv.divisionid" & vbCrLf
        s = s & "left join tblvendors av on p.BuildPro_AssignedVendor=av.vendor_id and p.divisionid=av.divisionid" & vbCrLf
        s = s & "left join tbljobs j on p.job=j.job_no and p.divisionid=j.divisionid" & vbCrLf
        s = s & "left join tblpoindex x on p.poindex=x.poindex and p.divisionid=x.divisionid" & vbCrLf
        s = s & "where p.divisionid = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and isnull(BuildPro_AssignedVendor,'')<>''" & vbCrLf
        s = s & "group by p.PONumber,p.Description ,av.Vendor_ID ,av.Vendor_Name ,tv.Vendor_ID ,tv.Vendor_Name ,p.Job,j.Description ,p.POIndex,x.description " & vbCrLf
        s = s & "order by 1,4,7 --> PONumber,AssignedVendorName,Job" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        
        While Not rs.EOF
            
            r = r + 1
            .AddItem ""
            .TextMatrix(r, .ColIndex("PONumber")) = "" & rs("PONumber")
            
            
            .TextMatrix(r, .ColIndex("PODesc")) = "" & rs("PODesc")
            .TextMatrix(r, .ColIndex("AssignedVendor")) = "" & rs("AssignedVendor")
            .TextMatrix(r, .ColIndex("AssignedVendorName")) = "" & rs("AssignedVendorName")
            .TextMatrix(r, .ColIndex("TBDVendor")) = "" & rs("TBDVendor")
            .TextMatrix(r, .ColIndex("TBDVendorName")) = "" & rs("TBDVendorName")
            .TextMatrix(r, .ColIndex("Job")) = "" & rs("Job")
            .TextMatrix(r, .ColIndex("JobDesc")) = "" & rs("JobDesc")
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(r, .ColIndex("POIndexDesc")) = "" & rs("POIndexDesc")
            .TextMatrix(r, .ColIndex("BudgetAmt")) = "" & rs("BudgetAmt")
            .TextMatrix(r, .ColIndex("VendorAmt")) = "" & rs("VendorAmt")
            
            If Val("" & rs("VendorAmt")) <> Val("" & rs("BudgetAmt")) Then
                .Cell(flexcpForeColor, r, .ColIndex("VendorAmt")) = vbRed
            End If
            
            .Cell(flexcpChecked, r, .ColIndex("PONumber")) = flexUnchecked
            
            rs.MoveNext
        Wend
        
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With gData
        .Cell(flexcpChecked, .Row, .ColIndex("PONumber"), .RowSel, .ColIndex("PONumber")) = .Cell(flexcpChecked, Row, .ColIndex("PONumber"))
    End With
End Sub

Private Sub gData_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    If gData.Rows < 2 Then Exit Sub
    Call gData.GetSelection(r1, c1, r2, c2)
    c1 = gData.ColIndex("PONumber")
    Call gData.Select(r1, c1, r2, c2)
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        Cancel = .ColKey(Col) <> "PONumber"
    End With
End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If gData.MouseRow = 0 And Button = vbRightButton Then
        Cancel = True
        Call FMain.ShowColumnMenu(gData, , , , , False)
    End If
End Sub

