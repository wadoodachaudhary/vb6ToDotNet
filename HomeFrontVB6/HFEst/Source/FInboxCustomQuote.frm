VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FInboxCustomQuote 
   Caption         =   "Custom Requests"
   ClientHeight    =   6135
   ClientLeft      =   4320
   ClientTop       =   1875
   ClientWidth     =   8400
   Icon            =   "FInboxCustomQuote.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8400
   Begin VB.Frame frmView 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   240
      Left            =   5700
      TabIndex        =   4
      Top             =   600
      Width           =   1905
      Begin HFEst.VBCombo cboView 
         Height          =   240
         Left            =   540
         TabIndex        =   6
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   15
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   315
      Index           =   0
      Left            =   6180
      TabIndex        =   2
      Top             =   4620
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Index           =   1
      Left            =   7260
      TabIndex        =   1
      Top             =   4620
      Width           =   1035
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   8175
      _cx             =   14420
      _cy             =   4154
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
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   38
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FInboxCustomQuote.frx":000C
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
   Begin HFEst.WizHead WizHead 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   1588
      Caption         =   "Prepare quotes for custom requests"
      Description     =   "Double click on a request to build the quote. Requests can also be declined."
      Icon            =   "FInboxCustomQuote.frx":05F2
   End
End
Attribute VB_Name = "FInboxCustomQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FInboxCustomQuote::"

Private mReadOnly As Boolean
Private mDirty   As Boolean
Private mTimerTask As String 'stupid menus

Private Sub cboView_Click()
    Call LoadData
    gData.Editable = flexEDKbdMouse ' IIf(cboView.ListIndex = 0, flexEDKbdMouse, flexEDNone)
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If SaveData(Index = 1) Then Unload Me
End Sub
Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    With cboView
        .Clear
        .AddItem "New requests"
        .AddItem "180 days old"
        .AddItem "1 year old"
        .AddItem "All old requests"
        .ListIndex = 0
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If SaveData(True) Then
        Call IniPutForm(Me)
        Call IniPutGrid(Me, gData)
    Else
        Cancel = True
    End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    frmView.Left = Me.ScaleWidth - frmView.Width - 835
    gData.Move margin, WizHead.Height + margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - WizHead.Height - margin * 3 - cmdNav(1).Height
    cmdNav(0).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
    cmdNav(1).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
End Sub




Private Sub LoadData()
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Dim s As String
    Dim maxrows As String
    
    Dim r As Long
    Dim c As Long
    Dim ra As Long
    With gData
        mReadOnly = True
        .Redraw = flexRDNone
        .Rows = 2
        
        maxrows = HFApp.Options.ValueByName("MaxUnQuotedOptions")
        If maxrows = "" Then maxrows = "5500"
        If Val(maxrows) = 0 Then
            s = ""
            s = s & "select CommunityDesc, Customer_No, Customer, creationdate, Model, ModelDesc, Job_No, TaxRate, Address, Location, Sales_Person_Name, ChangeOrderNo, Description, Comments, Category,major_group, CategoryDesc, EstimatorNotes, Location, seq, Qty, UOM, OptionType, Cost, Price, o.Community, AttachmentID ,saleslocation,color,style,finish,other,bldr_declined_reason"
        Else
            s = ""
            s = s & "select top " & maxrows & " CommunityDesc, Customer_No, Customer, creationdate, Model, ModelDesc, Job_No, TaxRate, Address, Location, Sales_Person_Name, ChangeOrderNo, Description, Comments, Category,major_group, CategoryDesc, EstimatorNotes, Location, seq, Qty, UOM, OptionType, Cost, Price, o.Community, AttachmentID ,saleslocation,color,style,finish,other,bldr_declined_reason "
        End If
        
        Select Case cboView.ListIndex
        Case 0
            mReadOnly = False
            s = s & "from UnQuotedOptions o "
            s = s & " where o.divisionid = " & HFApp.DivisionID & vbCrLf
        Case 1
            s = s & ",status "
            s = s & "from QuotedOptions o "
            s = s & " where o.divisionid = " & HFApp.DivisionID & " and datediff(d,o.quote_entered_date,getdate())<=180 " & vbCrLf
           
        Case 2
            s = s & ",status "
            s = s & " from QuotedOptions o "
            s = s & " where o.divisionid = " & HFApp.DivisionID & " and datediff(yy,quote_entered_date,getdate())<=1" & vbCrLf
            
        Case 3
            s = s & ",status "
            s = s & " from QuotedOptions o "
            s = s & " where o.divisionid = " & HFApp.DivisionID & vbCrLf
        End Select
        
        s = s & "order by o.creationdate desc,o.communitydesc,o.customer_no,o.changeorderno,o.seq"
        
        'disable save button
        cmdNav(0).Enabled = Not mReadOnly
        
        Set cn = New ADODB.Connection
        cn.Open (HFApp.ConnectionString(dbHomefront))
    
        Set rs = cn.Execute(s, ra)
        'HFApp.SqlExec(s, dbHomeFront)
    
        
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            On Error Resume Next
            
            
            For c = 0 To .Cols - 1
                Select Case .ColKey(c)
                    Case "CreationDate":     .Cell(flexcpText, r, c) = format("" & rs(.ColKey(c)), "Long Date")
                    Case "Job_no":           .Cell(flexcpText, r, c) = HFApp.FormatJob("" & rs(.ColKey(c)))
                    Case "Location":         .Cell(flexcpText, r, c) = "" & rs("Location")
                    Case Else:               .Cell(flexcpText, r, c) = "" & rs(.ColKey(c))
                End Select
            Next
            On Error GoTo 0
            rs.MoveNext
        Wend
        .Cell(flexcpBackColor, 1, 0, 1, .Cols - 1) = vbInfoBackground
        .FrozenRows = 1
        .Redraw = flexRDBuffered
    End With
    mDirty = False
End Sub

Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    Dim qty As Double
    Dim taxrate As Double
    Dim Amount  As Double
    Dim cost    As Double
    Dim Customer As String
    
    If mReadOnly Or Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    
'select top 50
'  Qty
' ,rate
' ,tax            -- rate * taxrate
' ,totalamount    -- rate * (1+taxrate)
' ,override_price -- qty * rate
' ,nettax         -- qty * rate * taxrate
' ,grandtotal     -- qty * rate * (1+taxrate)
'From tblscheduleb Or changeorderdetails (for designcenterdetails override_price is named amount)
'Where Qty > 1
    
    Screen.MousePointer = vbHourglass
    With gData
    For i = 2 To .Rows - 1
        If .RowData(i) = "Dirty" Then
        
            Customer = .TextMatrix(i, .ColIndex("Customer_No"))
            qty = Val(.TextMatrix(i, .ColIndex("Qty")))
            taxrate = Val(.TextMatrix(i, .ColIndex("TaxRate"))) / 100
            
            If HFApp.Options.value(includetax) <> "True" Then taxrate = 0
            
            Amount = Val(.TextMatrix(i, .ColIndex("Price")))
            cost = Val(.TextMatrix(i, .ColIndex("Cost")))
            s = ""
            Select Case .TextMatrix(i, .ColIndex("Location"))
                Case "ADDENDUM":       s = s & "update tblscheduleb" & vbCrLf
                Case "CHANGEORDER":    s = s & "update changeorderdetails" & vbCrLf
                Case "DESIGNCENTER":   s = s & "update designcenterdetails" & vbCrLf
                Case Else
                    Err.Raise 5, , "invalid option location: " & .TextMatrix(i, .ColIndex("Location"))
            End Select
            s = s & "   set quote_entered_date=" & DbQuote(Date, IIf(Trim(.TextMatrix(i, .ColIndex("Status"))) = "", "null", Now())) & vbCrLf
            s = s & "      ,Require_Quote=" & DbQuote(Bit, Trim(.TextMatrix(i, .ColIndex("Status"))) = "") & vbCrLf
            s = s & "      ,Bldr_Declined=" & DbQuote(Bit, .TextMatrix(i, .ColIndex("Status")) = "Declined") & vbCrLf
            s = s & "      ,Bldr_Declined_Reason=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Bldr_Declined_Reason"))) & vbCrLf
            s = s & "      ,EstimatorNotes=" & DbQuote(Str, .TextMatrix(i, .ColIndex("EstimatorNotes"))) & vbCrLf
            s = s & "      ,Description=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Description"))) & vbCrLf
            s = s & "      ,category=" & DbQuote(Str, .TextMatrix(i, .ColIndex("category"))) & vbCrLf
            s = s & "      ,major_group=" & DbQuote(Str, .TextMatrix(i, .ColIndex("major_group"))) & vbCrLf
            s = s & "      ,Comments=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Comments"))) & vbCrLf
            s = s & "      ,UOM=" & DbQuote(Str, .TextMatrix(i, .ColIndex("UOM"))) & vbCrLf
            s = s & "      ,Qty   =" & DbQuote(Num, qty) & vbCrLf
            s = s & "      ,Rate  =" & DbQuote(Num, Amount) & vbCrLf
            s = s & "      ,Cost_amount  =" & DbQuote(Num, cost) & vbCrLf
            s = s & "      ,tax   =" & DbQuote(Num, Round(Amount * taxrate, 2)) & vbCrLf
            s = s & "      ,totalamount=" & DbQuote(Num, Round(Amount * (1 + taxrate), 2)) & vbCrLf
            Select Case .TextMatrix(i, .ColIndex("Location"))
                Case "ADDENDUM", "CHANGEORDER": s = s & "      ,override_price=" & DbQuote(Num, qty * Amount) & vbCrLf
                Case "DESIGNCENTER":            s = s & "      ,amount=" & DbQuote(Num, qty * Amount) & vbCrLf
            End Select
            s = s & "      ,nettax    =" & DbQuote(Num, Round(qty * Amount * taxrate, 2)) & vbCrLf
            s = s & "      ,grandtotal=" & DbQuote(Num, Round(qty * Amount * (1 + taxrate), 2)) & vbCrLf
            s = s & "where seq=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Seq"))) & vbCrLf
            Call HFApp.SqlExec(s)
            
        
            Select Case .TextMatrix(i, .ColIndex("Location"))
                Case "ADDENDUM"
                    s = "Exec dbo.CalcAddendumTotals " & DbQuote(Str, Customer)
                    Call HFApp.SqlExec(s)
                Case "CHANGEORDER"
                    s = "Exec dbo.SalesCalcCustCOMaster " & DbQuote(Str, Customer)
                    Call HFApp.SqlExec(s)
                    s = "Exec dbo.CalcCOMasterTotals " & DbQuote(Str, Customer)
                    Call HFApp.SqlExec(s)
                Case "DESIGNCENTER"
                    s = "Exec dbo.SalesCalcCustDCMaster " & DbQuote(Str, Customer)
                    Call HFApp.SqlExec(s)
                    s = "Exec dbo.CalcDCCOMasterTotals " & DbQuote(Str, Customer)
                    Call HFApp.SqlExec(s)
            End Select
            s = "Exec dbo.CustomerCalcTotals " & DbQuote(Str, Customer)
            Call HFApp.SqlExec(s)
            
            s = "error in call to SendSaleRepEmail()"
            Call SendSaleRepEmail(.TextMatrix(i, .ColIndex("Location")), .TextMatrix(i, .ColIndex("Seq")))
        
        
            .RowData(i) = ""
        End If
        
    Next
    mDirty = False
    End With
    SaveData = True
    
    Screen.MousePointer = vbDefault
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function

Private Sub SendSaleRepEmail(Location, seq)
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select" & vbCrLf
    s = s & "  isnull(nullif(s.email,''),u.email) Email" & vbCrLf
    s = s & " ,c.customer_no,c.Description CustDesc,c.Job_no,c.municipal_address" & vbCrLf
    s = s & " ,o.changeorderno" & vbCrLf
    s = s & " ,case when o.bldr_declined=1 then 'declined' else 'completed' end status" & vbCrLf
    s = s & " ,o.description" & vbCrLf
    s = s & " ,o.bldr_declined_reason reason" & vbCrLf
    s = s & "from alloptions o" & vbCrLf
    s = s & " join tblcustomers c on o.customer_no=c.customer_no" & vbCrLf
    s = s & " left join tblsales_persons s on o.sales_person_id = s.sales_person_id" & vbCrLf
    s = s & " left join user_manager u on o.sales_person_id = u.user_id" & vbCrLf
    s = s & "where o.require_quote=0" & vbCrLf
    s = s & "  and isnull(isnull(nullif(s.email,''),u.email),'') <>''" & vbCrLf
    s = s & "  and o.location=" & DbQuote(Str, Location) & vbCrLf
    s = s & "  and o.seq=" & DbQuote(Num, seq) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    If rs.EOF Then Exit Sub
    
    s = "Your price request has been " & rs("Status") & vbCrLf & vbCrLf
    s = s & "Job: " & rs("Job_No") & "  " & rs("municipal_address") & vbCrLf
    s = s & "Customer: " & rs("Customer_no") & "  " & rs("CustDesc") & vbCrLf
    s = s & "ChangeOrder: " & rs("changeorderno") & vbCrLf
    s = s & "Description: " & rs("description") & vbCrLf & vbCrLf
    If "" & rs("Status") = "declined" Then
        s = s & "Reason: " & rs("reason")
    End If
    Call HFApp.SendMail(False, rs("email"), "", "Custom Option Price Request " & rs("Status"), s, "")
    
Exit Sub
eh: Call errHandler(SRCFILE & "SendSaleRepEmail", s)
End Sub


Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim r As Long, c As Long
With gData
        If .Row = 1 Then
            .Redraw = flexRDNone
            For r = 2 To .Rows - 1
                .RowHidden(r) = False
                If .EditText <> "" Then
                    For c = 0 To .Cols - 1
                    If .ColKey(c) <> "Status" Then
                        If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                            .RowHidden(r) = True
                            Exit For
                        End If
                    End If
                    Next
                End If
            Next
            .Redraw = flexRDBuffered
        End If
End With
End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    With gData
        Select Case True
        
            Case Button = vbRightButton
                Cancel = True
                Call FMain.ShowColumnMenu(gData)
                
                
            Case Button = vbRightButton
                
        End Select
    End With

End Sub

Private Sub gData_BeforeSort(ByVal Col As Long, Order As Integer)
    gData.FixedRows = 2
End Sub
Private Sub gData_AfterSort(ByVal Col As Long, Order As Integer)
    gData.FixedRows = 1
End Sub
Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        .AutoSearch = flexSearchNone
        .ComboList = ""

        Select Case True
            Case Row = 1
            
            Case IsIn(.ColKey(Col), "Price", "Cost", "Status")
                    If cboView.ListIndex <> 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
            Case .ColKey(Col) = "Bldr_Declined_Reason"
                    If cboView.ListIndex <> 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                
                .EditMaxLength = 500
                .ComboList = "|..."
            
            Case .ColKey(Col) = "CategoryDesc"
                If cboView.ListIndex <> 0 Then
                    Cancel = True
                    Exit Sub
                End If
                .ComboList = "..."
            
            Case IsIn(.ColKey(Col), "EstimatorNotes", "Comments")
                If cboView.ListIndex <> 0 Then
                    .ComboList = "..."
                Else
                    .EditMaxLength = 4000
                    .ComboList = "|..."
                End If
            
            Case IsIn(.ColKey(Col), "Description")
                    If cboView.ListIndex <> 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                
                .EditMaxLength = 200
                .ComboList = "|..."
                
            Case Else
                Cancel = True
                .AutoSearch = flexSearchFromCursor
        End Select
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gData
        Select Case .ColKey(Col)
        Case "CategoryDesc"
            s = "select Description,category,group_code from tblcategories order by 1"
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, , , , , "category,group_code") Then Exit Sub
            .TextMatrix(Row, .ColIndex("CategoryDesc")) = FPickList.SelectedItem("description")
            .TextMatrix(Row, .ColIndex("Category")) = FPickList.SelectedItem("Category")
            .TextMatrix(Row, .ColIndex("major_group")) = FPickList.SelectedItem("major_group")
            .RowData(Row) = "Dirty"
            mDirty = True
            
        
        Case "Description", "Comments", "Bldr_Declined_Reason", "EstimatorNotes"
            s = .Text
            If FComments.Edit(s, gData, cboView.ListIndex = 0, .TextMatrix(0, Col), IIf(Col = .ColIndex("Description"), 200, 4000)) Then
                If cboView.ListIndex <> 0 Then Exit Sub
                
                .Text = s
                .RowData(Row) = "Dirty"
                
                
                If .ColKey(Col) = "Bldr_Declined_Reason" Then
                    .TextMatrix(Row, .ColIndex("Status")) = "Declined"
                End If
                
                mDirty = True
            End If
        End Select
    End With
End Sub

Private Sub gData_ChangeEdit()
    Dim b As Boolean
    Dim c As Long
    Dim r As Long
    
    Dim t As String
    
    With gData
        If .Row = 1 Then Exit Sub
'        For r = 2 To .Rows - 1
'            b = True
'            For c = 0 To .cols - 1
'
'                t = IIf(c = .Col, .EditText, .TextMatrix(1, c))
'
'                b = b And (t = "" Or (UCase(.TextMatrix(r, c)) Like "*" & UCase(t) & "*"))
'
'            Next
'            .RowHidden(r) = Not b
'
'        Next
    End With
End Sub


Private Sub gData_DblClick()
    Call gData_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo eh
    Dim cost As Double
    Dim price As Double
    Dim qty As Double
    Dim UOM As String
    
    With gData
    Select Case KeyCode
    
        Case vbKeyReturn
            If .Row < 2 Then Exit Sub
            
        
            UOM = .TextMatrix(.Row, .ColIndex("UOM"))
            cost = .ValueMatrix(.Row, .ColIndex("Cost"))
            price = .ValueMatrix(.Row, .ColIndex("Price"))
            qty = .ValueMatrix(.Row, .ColIndex("Qty"))
            If qty = 0 Then qty = 1
            
            If FCustomQuote.BuildQuote(cboView.ListIndex > 0 _
                                       , .TextMatrix(.Row, .ColIndex("AttachmentID")) _
                                       , .TextMatrix(.Row, .ColIndex("Location")) _
                                       , .TextMatrix(.Row, .ColIndex("Model")) _
                                       , .TextMatrix(.Row, .ColIndex("OptionID")) _
                                       , .TextMatrix(.Row, .ColIndex("Job_No")) _
                                       , .TextMatrix(.Row, .ColIndex("Customer_No")) _
                                       , .TextMatrix(.Row, .ColIndex("Community")) _
                                       , .TextMatrix(.Row, .ColIndex("CommunityPhase")) _
                                       , .TextMatrix(.Row, .ColIndex("Extra")) _
                                       , .TextMatrix(.Row, .ColIndex("Description")) _
                                       , .TextMatrix(.Row, .ColIndex("Seq")) _
                                       , .TextMatrix(.Row, .ColIndex("Category")) _
                                       , cost, price, qty, UOM) Then
                                       
                .RowData(.Row) = "Dirty"
                .TextMatrix(.Row, .ColIndex("Cost")) = cost * qty
                .TextMatrix(.Row, .ColIndex("Price")) = price
                .TextMatrix(.Row, .ColIndex("Qty")) = qty
                .TextMatrix(.Row, .ColIndex("UOM")) = UOM
                mDirty = True
                
            End If

    
        Case vbKeyDelete
            If .Row <> 1 Then Exit Sub
            .Text = ""
            Call gData_AfterEdit(.Row, .Col)
            
            
    End Select
    End With
    Exit Sub
eh:
    If Err.Number <> 0 Then
         MsgBox "Inbox custom quote " & vbCrLf & Err.Description
         
    End If
End Sub


Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim s As String

If Row > 1 Then
    With gData
        
        Select Case .ColKey(Col)
        
            Case "Status"
                If .EditText = "Declined" And .TextMatrix(Row, .ColIndex("Bldr_Declined_Reason")) = "" And Row <> 1 Then
                    s = InputBox("Reason:", App.ProductName, .TextMatrix(Row, .ColIndex("Bldr_Declined_Reason")))
                    If s = "" Then
                        Cancel = True
                    Else
                        .TextMatrix(Row, .ColIndex("Status")) = "Declined"
                        .TextMatrix(Row, .ColIndex("Bldr_Declined_Reason")) = s
                        .TextMatrix(Row, .ColIndex("Price")) = ".00"
                    End If
                End If
                
                
            Case "Bldr_Declined_Reason"
                .TextMatrix(Row, .ColIndex("status")) = "Declined"
                
            Case "Price"
                .EditText = Val(.EditText)
                If Val(.EditText) <> 0 Then .TextMatrix(Row, .ColIndex("Status")) = "Approved"
                
            Case "Cost"
                .EditText = Val(.EditText)
                
            Case "EstimatorNotes"
        
        End Select
        If .TextMatrix(Row, .ColIndex("Status")) <> "Declined" Then
            .TextMatrix(Row, .ColIndex("Bldr_Declined_Reason")) = ""
        End If
        .RowData(Row) = "Dirty"
        mDirty = True
    End With
End If
End Sub











