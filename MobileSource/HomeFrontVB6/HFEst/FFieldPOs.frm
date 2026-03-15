VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FFieldPOs 
   Caption         =   "Field PO Requests"
   ClientHeight    =   8895
   ClientLeft      =   1245
   ClientTop       =   2130
   ClientWidth     =   9015
   Icon            =   "FFieldPOs.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   9015
   Begin HFEst.Slider Slider 
      Height          =   45
      Left            =   90
      Top             =   3900
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   79
      Orientation     =   1
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1058
      ButtonWidth     =   1376
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Approve"
            Key             =   "Approve"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Decline"
            Key             =   "Decline"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid gRequests 
      Height          =   3135
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   8535
      _cx             =   1999125199
      _cy             =   1999115674
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FFieldPOs.frx":000C
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
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
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
   Begin VSFlex8Ctl.VSFlexGrid gDetails 
      Height          =   3135
      Left            =   90
      TabIndex        =   2
      Top             =   4050
      Width           =   8535
      _cx             =   1999125199
      _cy             =   1999115674
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
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FFieldPOs.frx":01F3
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
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
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
End
Attribute VB_Name = "FFieldPOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FFieldPOs"
Private mview As Long

Private Sub Form_Load()
On Error Resume Next
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gRequests)
    Call IniGetGrid(Me, gDetails)
    Call LoadRequests
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Slider.Move 0, Slider.Top, Me.ScaleWidth
    gRequests.Move 0, Toolbar.Height - Screen.TwipsPerPixelY, Me.ScaleWidth, Slider.Top - Toolbar.Height + -Screen.TwipsPerPixelY
    gDetails.Move 0, Slider.Top + Slider.Height, Me.ScaleWidth, Me.ScaleHeight - Slider.Top - Slider.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gRequests)
    Call IniPutGrid(Me, gDetails)
End Sub

Private Sub LoadRequests()
On Error GoTo eh
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "SELECT" & vbCrLf
    s = s & "  fpr.RequestID" & vbCrLf
    s = s & " ,fpu.USER_NAME RequestedBy" & vbCrLf
    s = s & " ,fpr.RequestDate RequestedDate" & vbCrLf
    s = s & " ,j.Job_No Job" & vbCrLf
    s = s & " ,j.Description JobDesc" & vbCrLf
    s = s & " ,fpr.POIndex" & vbCrLf
    s = s & " ,fpr.VendorID Vendor" & vbCrLf
    s = s & " ,v.Vendor_Name VendorDesc" & vbCrLf
    s = s & " ,fpr.Description RequestDesc" & vbCrLf
    s = s & " ,fpr.BackChargeVendorID BCVendor" & vbCrLf
    s = s & " ,bc.Vendor_Name BCVendorDesc" & vbCrLf
    s = s & " ,fpr.BackChargeComments Comments" & vbCrLf
    s = s & "FROM FieldPORequests fpr " & vbCrLf
    s = s & "LEFT OUTER JOIN tblVendors v ON fpr.VendorID = v.Vendor_ID and fpr.DivisionID = v.DivisionID" & vbCrLf
    s = s & "LEFT OUTER JOIN tblVendors bc ON fpr.BackChargeVendorID = bc.Vendor_ID and fpr.DivisionID = v.DivisionID" & vbCrLf
    s = s & "LEFT OUTER JOIN FieldPOUsers fpu ON fpu.User_ID = fpr.UserID" & vbCrLf
    s = s & "LEFT OUTER JOIN tblJobs j ON fpr.Job=j.Job_No and fpr.DivisionID = j.DivisionID" & vbCrLf
    s = s & "WHERE fpr.Status=0" & vbCrLf
    s = s & "and fpr.DivisionID = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "order by fpr.RequestID"
    Set rs = HFApp.SqlExec(s, dbHomefront)

    With gRequests
        .Rows = 1
        While Not rs.EOF
            .AddItem ""
            For c = 0 To .Cols - 1
                .TextMatrix(.Rows - 1, c) = "" & rs(.ColKey(c))
            Next
            Call rs.MoveNext
        Wend
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadRequests", s)
End Sub


Private Sub LoadDetails(RequestID As Long)
On Error GoTo eh
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "SELECT" & vbCrLf
    s = s & "  d.sequence" & vbCrLf
    s = s & " ,d.phase" & vbCrLf
    s = s & " ,d.item" & vbCrLf
    s = s & " ,d.description" & vbCrLf
    s = s & " ,d.job" & vbCrLf
    s = s & " ,j.Description JobDesc" & vbCrLf
    s = s & " ,d.jcextra" & vbCrLf
    s = s & " ,d.jccostcode" & vbCrLf
    s = s & " ,scc.description jccostcodedesc" & vbCrLf
    s = s & " ,d.jccategory" & vbCrLf
    s = s & " ,scat.description jccategorydesc" & vbCrLf
    s = s & " ,d.taxgroup" & vbCrLf
    s = s & " ,round(d.qty,4) Qty" & vbCrLf
    s = s & " ,round(d.rate,4) Rate" & vbCrLf
    s = s & " ,round(d.amount,2) Amount" & vbCrLf
    s = s & " ,round(d.jctax,2) jctax" & vbCrLf
    s = s & " ,round(d.njctax,2) njctax" & vbCrLf
    s = s & " ,d.orderuom uom" & vbCrLf
    s = s & " ,round(d.totalamount,2) totalamount" & vbCrLf
    s = s & "FROM FieldPORequests fpr " & vbCrLf
    s = s & "LEFT OUTER JOIN tblVendors v ON fpr.VendorID = v.Vendor_ID and fpr.DivisionID = v.DivisionID" & vbCrLf
    s = s & "LEFT OUTER JOIN tblVendors bc ON fpr.BackChargeVendorID = bc.Vendor_ID and fpr.DivisionID = v.DivisionID" & vbCrLf
    s = s & "LEFT OUTER JOIN FieldPOUsers fpu ON fpu.User_ID = fpr.UserID" & vbCrLf
    s = s & "left outer join fieldporequestdetails d on fpr.requestid=d.requestid" & vbCrLf
    s = s & "LEFT OUTER JOIN tblJobs j ON d.Job=j.Job_No and fpr.DivisionID = j.DivisionID" & vbCrLf
    s = s & "left outer join standardcostcodes scc on d.jccostcode=scc.costcode and fpr.DivisionID = scc.DivisionID" & vbCrLf
    s = s & "left outer join standardcategories scat on d.jccategory=scat.category and fpr.DivisionID = scat.DivisionID" & vbCrLf
    s = s & "WHERE fpr.RequestID=" & DbQuote(Num, RequestID) & vbCrLf
    s = s & "and fpr.DivisionID = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "order by d.sequence"
    Set rs = HFApp.SqlExec(s, dbHomefront)

    With gDetails
        .Rows = 1
        While Not rs.EOF
            .AddItem ""
            For c = 0 To .Cols - 1
                .TextMatrix(.Rows - 1, c) = "" & rs(.ColKey(c))
            Next
            Call rs.MoveNext
        Wend
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadDetails", s)
End Sub

Private Sub gRequests_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error Resume Next
    If NewRow <> OldRow And NewRow > 0 Then
        Call LoadDetails(gRequests.TextMatrix(NewRow, gRequests.ColIndex("RequestID")))
    End If
End Sub

Private Sub gRequests_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = vbRightButton And gRequests.MouseRow = 0 Then
        Call FMain.ShowColumnMenu(gRequests)
    End If
End Sub

Private Sub Slider_Move()
On Error Resume Next
    Call Form_Resize
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim r As Long
    Dim s As String
    
    With gRequests
    r = .RowSel - .Row + 1
    If r > 1 Then
        Select Case Button.Key
            Case "Approve": s = "Approve these " & r & " requests?"
            Case "Decline": s = "Decline these " & r & " requests?"
        End Select
        If MsgBox(s, vbYesNo + vbQuestion, "Confirm") = vbNo Then Exit Sub
    End If
    
    For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
        Select Case Button.Key
            Case "Approve": s = "exec FieldPO_ApprovePORequest " & DbQuote(Num, .TextMatrix(r, .ColIndex("RequestID"))) & ", " & DbQuote(Date, Now()) & "," & DbQuote(Str, HFApp.LoginID)
            Case "Decline": s = "exec FieldPO_DeclinePORequest " & DbQuote(Num, .TextMatrix(r, .ColIndex("RequestID"))) & ", " & DbQuote(Date, Now())
        End Select
        Call HFApp.SqlExec(s)
        Call .RemoveItem(r)
    Next
    .SetFocus
    End With
    
Exit Sub
eh: 'Call errHandler(SRCFILE & "Toolbar_ButtonClick",s)
s = Err.Description
s = Parse(s, Parse(s, , "]"), "]")
MsgBox s, vbExclamation, App.ProductName

End Sub


