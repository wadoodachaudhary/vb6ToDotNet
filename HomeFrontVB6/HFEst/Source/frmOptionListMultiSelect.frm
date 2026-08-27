VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{C9460280-3EED-11D0-A647-00A0C91EF7B9}#1.0#0"; "ImageViewer2.OCX"
Begin VB.Form frmOptionListMultiSelect 
   Caption         =   "HSMC - Option List"
   ClientHeight    =   10050
   ClientLeft      =   3930
   ClientTop       =   3450
   ClientWidth     =   16080
   Icon            =   "frmOptionListMultiSelect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   16080
   StartUpPosition =   2  'CenterScreen
   Begin SCRIBBLELib.ImageViewer ImageViewer1 
      Height          =   3435
      Left            =   10170
      TabIndex        =   11
      Top             =   5490
      Width           =   4635
      _Version        =   65536
      _ExtentX        =   8176
      _ExtentY        =   6059
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   12000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Save to Customer"
      Height          =   450
      Left            =   10080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1620
   End
   Begin TDBText6Ctl.TDBText txtLocation 
      Height          =   375
      Left            =   11280
      TabIndex        =   8
      Top             =   960
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "frmOptionListMultiSelect.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOptionListMultiSelect.frx":0078
      Key             =   "frmOptionListMultiSelect.frx":0096
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   "TDBText1"
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid gitems 
      Height          =   8985
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   9735
      _cx             =   17171
      _cy             =   15849
      Appearance      =   1
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
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
   Begin VB.Frame Frame4 
      Caption         =   "Option Type"
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13830
      Begin VB.OptionButton Option5 
         Caption         =   "Change Requests"
         Height          =   420
         Left            =   10920
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Customer Shopping Cart"
         Height          =   420
         Left            =   8640
         TabIndex        =   4
         Top             =   120
         Width           =   2340
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Personal Choice Options"
         Height          =   420
         Left            =   5760
         TabIndex        =   3
         Top             =   120
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Model Specific Options"
         Height          =   420
         Left            =   3120
         TabIndex        =   2
         Top             =   120
         Width           =   3360
      End
      Begin VB.OptionButton Option3 
         Caption         =   "All Options Combined"
         Height          =   420
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         Width           =   2265
      End
   End
   Begin MSAdodcLib.Adodc curPick 
      Height          =   330
      Left            =   1440
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   16
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "HSMC"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ADMIN_STDOPT"
      Caption         =   "curPick"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc curPick1 
      Height          =   330
      Left            =   6360
      Top             =   7200
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   16
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "curPick1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLocation 
      Caption         =   "Location"
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptionListMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vExtra_Code_No As String
Dim vindex As Integer
Dim vESC As Integer
Dim curTemp As ADODB.Recordset
Dim UseCORequests As Boolean
Dim itemschecked As Boolean
Dim oldgcall As Integer
Dim CustomerNo  As String
Dim sJob As String
Dim rs As Recordset
Option Compare Text


Private Sub cmdCancel_Click()
Dim r As Long

If itemschecked Then
    If MsgBox("Close Form and return to Job?", vbYesNo) = vbYes Then
        gCancel = True
        Unload Me
        Set frmOptionListMultiSelect = Nothing
    Else
        itemschecked = False
        If Option1.value = True Then
            Call Option1_Click
        ElseIf Option2.value = True Then
            Call Option2_Click
        ElseIf Option3.value = True Then
            Call Option3_Click
        ElseIf Option4.value = True Then
            Call Option4_Click
        ElseIf Option5.value = True Then
            Call Option5_Click
        End If
        
    End If
Else
    gCancel = True
    Unload Me
    Set frmOptionListMultiSelect = Nothing
End If
End Sub

Private Sub cmdCancel_GotFocus()
Call cmdCancel_Click
End Sub
Private Sub cmdSelect_Click()


    Dim r As Long
    Dim i As Long
    Dim vAdded As Boolean
    Dim NewRow As Long
    Dim vSeq As Long
    Dim tax As Double
    Dim r2 As Long
    Dim ra As Long
   
    Dim s As String
    Dim firstJob As String
    Dim lastCustomer  As String
    Dim EstimateIndex As Long

    If Option1.value = True Then
       gOptionPick = 1
    ElseIf Option2.value = True Then
       gOptionPick = 2
    ElseIf Option3.value = True Then
        gOptionPick = 3
    Else
        gOptionPick = 4
    End If
    
On Error GoTo eh
'used to return boolean
'now returns the first job number processed

    
    firstJob = ""
    With gItems
                    EstimateIndex = HFApp.SqlExec("exec Purch_CreateCustomerEstimate " & DbQuote(str, rs!Customer_no) & "," & DbQuote(str, HFApp.LoginID))(0)
                        
        For r = 0 To .Rows - 1
            If gItems.Cell(flexcpChecked, r, gItems.ColIndex("Pick")) = flexChecked And gItems.IsSubtotal(r) = False Then
                s = "Insert into tblScheduleB(Customer_No,Item_Seq,Category,UserID"
                s = s & ",Cost_Amount,Qty,Option_Type,UOM,OPT,Assembly,Location,Color,Style,Finish,Other,Base_Price,Rate,Override_Price"
                s = s & ",Description,Comments,EstimatorNotes,TL_Extra"
                s = s & ",Major_Group,CreationDate,ColorListID,StyleListID,FinishListID,OtherListID) Values("
                s = s & DbQuote(str, rs!Customer_no) & "," & vbCrLf
                s = s & DbQuote(Num, 1) & "," & vbCrLf
                s = s & DbQuote(str, .TextMatrix(r, .ColIndex("Category"))) & "," & vbCrLf
                s = s & DbQuote(str, HFApp.LoginID) & "," & vbCrLf
                s = s & DbQuote(Num, Val("" & .ValueMatrix(r, .ColIndex("Cost_Amount")))) & "," & vbCrLf
                s = s & DbQuote(Num, IIf(Not IsNull(.ValueMatrix(r, .ColIndex("Qty"))), .ValueMatrix(r, .ColIndex("Qty")), 1)) & "," & vbCrLf
                
                If Option1.value = True Then
                   s = s & "1," & vbCrLf
                ElseIf Option2.value = True Then
                   s = s & "2," & vbCrLf
                Else
                   s = s & .ValueMatrix(r, .ColIndex("Option_Type")) & "," & vbCrLf
                End If
                s = s & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("UOM")), "S")) & "," & vbCrLf
                s = s & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("OPT")), "S")) & "," & vbCrLf
                s = s & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Assembly")), "S")) & "," & vbCrLf
                s = s & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Location")), "S")) & "," & vbCrLf
                s = s & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Color")), "S")) & "," & vbCrLf
                s = s & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Style")), "S")) & "," & vbCrLf
                s = s & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Finish")), "S")) & "," & vbCrLf
                s = s & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Other")), "S")) & "," & vbCrLf
                    
                s = s & DbQuote(Num, NullCheck(.ValueMatrix(r, .ColIndex("Price")))) & vbCrLf
                s = s & "," & DbQuote(Num, NullCheck(.ValueMatrix(r, .ColIndex("Price")))) & vbCrLf
                s = s & "," & DbQuote(Num, NullCheck(.ValueMatrix(r, .ColIndex("Price"))) * IIf(Not IsNull(.ValueMatrix(r, .ColIndex("Qty"))), .ValueMatrix(r, .ColIndex("Qty")), 1)) & vbCrLf
                s = s & "," & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Description")), "S")) & vbCrLf
                s = s & "," & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Comments")), "S")) & vbCrLf
                s = s & "," & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("EstimatorNotes")), "S")) & vbCrLf
                s = s & "," & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("TL_Extra")), "S")) & vbCrLf
                s = s & "," & DbQuote(str, NullCheck(.TextMatrix(r, .ColIndex("Major_Group")), "S")) & vbCrLf
                s = s & ",getdate()" & vbCrLf
                s = s & "," & DbQuote(Num, NullCheck(.ValueMatrix(r, .ColIndex("ColorListID")))) & vbCrLf
                s = s & "," & DbQuote(Num, NullCheck(.ValueMatrix(r, .ColIndex("StyleListID")))) & vbCrLf
                s = s & "," & DbQuote(Num, NullCheck(.ValueMatrix(r, .ColIndex("FinishListID")))) & vbCrLf
                s = s & "," & DbQuote(Num, NullCheck(.ValueMatrix(r, .ColIndex("OtherListID")))) & vbCrLf
                s = s & ")"
                HFApp.SqlExec s
                vSeq = HFApp.SqlIdentity("tblScheduleB", dbHomefront)
                      
                   
                
                    
                    s = "exec dbo.Purch_PutEstimateAssembly " & DbQuote(str, rs!Customer_no) & _
                                                          "," & DbQuote(Num, EstimateIndex) & _
                                                          "," & DbQuote(Num, 1) & _
                                                          "," & DbQuote(str, "") & _
                                                          "," & DbQuote(Num, vSeq) & _
                                                          "," & DbQuote(str, "", True) & _
                                                          "," & DbQuote(str, HFApp.LoginID) & _
                                                          "," & DbQuote(Num, HFApp.DivisionID)
                                                          
                    HFApp.SqlExec s

            End If
        Next
    End With
    
Unload Me
Exit Sub
eh: Call errHandler("frmOptionListMultiSelect " & "SaveData")
End Sub




Public Sub ShowForm(Job As String)
sJob = Job
Me.Show vbModal
End Sub


Private Sub Form_Load()
On Error Resume Next
Dim s As String
gCancel = False
Dim CustShoppingCart As Integer, OPTCount As Integer, GlobalOptCount As Integer, CORequests As Integer
txtLocation.Text = ""
txtLocation.DropWndHeight = 1
Call IniGetForm(Me)
Me.Caption = App.ProductName & " - " & "Option List"
oldgcall = gcall
UseCORequests = UCase("" & HFApp.SqlExec("Select OptionValue from AppOptions where DivisionID = " & HFApp.DivisionID & " and OptionName = 'ChangeRequests'")(0)) = "TRUE"
'Set curTemp = New ADODB.Recordset
curPick.ConnectionString = HFApp.ConnectionString(dbHomefront)
curPick.CommandType = adCmdText
CustShoppingCart = 0
Set rs = HFApp.SqlExec("select top 1 customer_no,isnull(Community,'') Community,gst_rate,Item_seq,isnull(Elevation,'') Elevation,isnull(Phase,'') Phase,Model,isnull(Series,'') Series from tblcustomers where job_no = " & DbQuote(str, sJob) & " and DivisionID = " & HFApp.DivisionID)
CustShoppingCart = Val("" & HFApp.SqlExec("Select count(customer_no) from CustomerShoppingCart where isnull(addtoco,0) = 0 and customer_no = " & DbQuote(str, rs!Customer_no))(0))


    Option2.value = 1
    If CustShoppingCart < 1 Then
        Option4.Visible = False
    End If
    DataGrid1.EvenRowStyle.BackColor = gGridBackColor
    If Option1.value = True Then
       'Call Option1_Click
    ElseIf Option2.value = True Then
       'Call Option2_Click
    End If
    If CustShoppingCart = 0 And UseCORequests Then
        Option5.Left = Option4.Left
    End If

Exit Sub


herror:
If Err.Number <> 0 Then
    Call errHandler("frmOptionListMultiSelect " & "FormLoad")
End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If Option1.value = True Then
   gOptionPick = 1
ElseIf Option2.value = True Then
   gOptionPick = 2
ElseIf Option3.value = True Then
    gOptionPick = 3
ElseIf Option4.value = True Then
    gOptionPick = 4
Else
    gOptionPick = 5
End If
Call IniPutGrid(Me, gItems)
Call IniPutForm(Me)

Exit Sub

herror:

MsgBox Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
On Error Resume Next
Const margin = 60
    gItems.Move margin * 2, Frame4.Top + Frame4.Height + margin, Me.ScaleWidth - (4 * margin) - ImageViewer1.Width, Me.ScaleHeight - Frame4.Top - Frame4.Height - (4 * margin)
    cmdSelect.Move gItems.Left + gItems.Width + 225, Me.ScaleHeight - 480
    cmdCancel.Move cmdSelect.Left + cmdSelect.Width + 480, cmdSelect.Top
    ImageViewer1.Move gItems.Left + gItems.Width + (2 * margin), cmdSelect.Top - 225 - ImageViewer1.Height
    lblLocation.Move ImageViewer1.Left
    txtLocation.Move ImageViewer1.Left + lblLocation.Width - (2 * margin)
   ' Frame1.Move gitems.left, Frame4.Height + gitems.Height + (3 * margin), Me.ScaleWidth - (3 * margin)
    'Me.Command6.Move Frame1.left + margin
    Me.cmdSelect.Move gItems.Left + gItems.Width + 225, Me.ScaleHeight - 480
    Me.cmdCancel.Move cmdSelect.Left + cmdSelect.Width + 480, cmdSelect.Top

End Sub

Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    Dim c As Long
    Dim Count As Long
    
    With gItems

        If Row <> 1 Then
            For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                If .ColIndex("Pick") = Col And .Cell(flexcpChecked, r, .ColIndex("Pick")) = flexChecked Then
                    If txtLocation.Text <> "" And .TextMatrix(r, .ColIndex("Location")) = "" Then
                        .TextMatrix(r, .ColIndex("Location")) = txtLocation.Text
                    End If
                End If
                If .ColIndex("Pick") = Col And .ValueMatrix(r, .ColIndex("Qty")) = 0 And .Cell(flexcpChecked, r, .ColIndex("Pick")) = flexChecked Then
                    .TextMatrix(r, .ColIndex("Qty")) = 1
                End If
                If .ColIndex("Qty") = Col And .Cell(flexcpChecked, r, .ColIndex("Pick")) = flexUnchecked And .ValueMatrix(r, .ColIndex("Qty")) <> 0 Then
                    .Cell(flexcpChecked, r, .ColIndex("Pick")) = flexChecked
                    If txtLocation.Text <> "" And .TextMatrix(r, .ColIndex("Location")) = "" Then
                        .TextMatrix(r, .ColIndex("Location")) = txtLocation.Text
                    End If
                End If
                If Col = .ColIndex("Location") Or Col = .ColIndex("Color") Or Col = .ColIndex("Style") Or Col = .ColIndex("Finish") Or Col = .ColIndex("Other") Or Col = .ColIndex("Comments") Or Col = .ColIndex("Price") Or Col = .ColIndex("CO_Price") Or Col = .ColIndex("GrandTotal") Or Col = .ColIndex("GrandCOTotal") Then
                    If .Cell(flexcpChecked, r, .ColIndex("Pick")) = flexUnchecked Then
                        .Cell(flexcpChecked, r, .ColIndex("Pick")) = flexChecked
                    End If
                    If .ValueMatrix(r, .ColIndex("Qty")) = 0 And .Cell(flexcpChecked, r, .ColIndex("Pick")) = flexChecked Then
                        .TextMatrix(r, .ColIndex("Qty")) = 1
                    End If
                    If Col <> .ColIndex("Location") And txtLocation.Text <> "" And .TextMatrix(r, .ColIndex("Location")) = "" Then
                        .TextMatrix(r, .ColIndex("Location")) = txtLocation.Text
                    End If
                    Call ColorizeItems(0)
                End If
            Next
            Exit Sub
        Else
            .Redraw = flexRDNone
            For r = 2 To .Rows - 1
                .RowHidden(r) = False
                If HFApp.Options(UseMajorGroup) Then
                    
                    If .TextMatrix(r, .ColIndex("MajDesc")) = "" And .TextMatrix(r, .ColIndex("catDesc")) = "" And .TextMatrix(r, .ColIndex("Description")) = "" Then .RowHidden(r) = True
                Else
                    If .TextMatrix(r, .ColIndex("catDesc")) = "" And .TextMatrix(r, .ColIndex("Description")) = "" Then .RowHidden(r) = True
                End If
                If .EditText <> "" Then
                    For c = 0 To .cols - 1
                        If .Cell(flexcpTextDisplay, 1, c) <> "" And .Cell(flexcpTextDisplay, 1, c) <> "$.00" Then
                        If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                        'If Not (UCase(.TextMatrix(r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
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

Private Sub gItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error Resume Next
With gItems
If NewRow <> 0 Then
    If Len(.TextMatrix(NewRow, .ColIndex("Comments"))) > 100 Then
        .RowHeight(NewRow) = 240 * 3
        '.RowHeight(newrow) = 240 * Len(.TextMatrix(newrow, .ColIndex("Comments"))) / 100
    
    Else
        .RowHeight(NewRow) = 240
    End If
    If NewRow <> OldRow Then .RowHeight(OldRow) = 240
    If .ColWidth(.ColIndex("Description")) / 75 < Len(.TextMatrix(NewRow, .ColIndex("Description"))) Then 'Len(.TextMatrix(.Row, .ColIndex("Description"))) > 50 Then
        .RowHeight(NewRow) = 240 * 3
        
    End If
    
    If OldRow > 1 Then
    .Cell(flexcpBackColor, OldRow, 1, OldRow, .cols - 1) = vbWindowBackground
    End If
        .Cell(flexcpBackColor, NewRow, 1, NewRow, .cols - 1) = gGridBackColor 'vbbuttonface
        Call ColorizeItems(OldRow)
        Call ColorizeItems(NewRow)
End If
End With



End Sub

Private Sub gItems_AfterSort(ByVal Col As Long, Order As Integer)
gItems.FixedRows = 1
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


    With gItems
        .ComboList = ""
        If Row = 1 Then
            Cancel = False
            .AutoSearch = flexSearchNone
        ElseIf Col = .ColIndex("Pick") Then
            Cancel = False
            .AutoSearch = flexSearchNone
        ElseIf Col = .ColIndex("NoCharge") Then
            Cancel = False
            .AutoSearch = flexSearchNone
        ElseIf Col = .ColIndex("Qty") Then
            Cancel = False
            .AutoSearch = flexSearchNone
        ElseIf Col = .ColIndex("Comments") Then
            Cancel = False
            .AutoSearch = flexSearchNone
            .ComboList = "|..."
        ElseIf Col = .ColIndex("Color") Then
            Cancel = False
            .AutoSearch = flexSearchNone
            If .ValueMatrix(Row, .ColIndex("ColorListID")) <> 0 Then
                .ComboList = "|..."
            Else
                .ComboList = ""
            End If
        ElseIf Col = .ColIndex("Location") Then
            .ComboList = "|..."
        ElseIf Col = .ColIndex("Style") Then
            Cancel = False
            .AutoSearch = flexSearchNone
            If .ValueMatrix(Row, .ColIndex("StyleListID")) <> 0 Then
                .ComboList = "|..."
            Else
                .ComboList = ""
            End If
        ElseIf Col = .ColIndex("Finish") Then
            Cancel = False
            .AutoSearch = flexSearchNone
            If .ValueMatrix(Row, .ColIndex("FinishListID")) <> 0 Then
                .ComboList = "|..."
            Else
                .ComboList = ""
            End If
        ElseIf Col = .ColIndex("GrandTotal") And HFApp.UserPermission("SalesOverride") Then
            Cancel = False
            .AutoSearch = flexSearchNone
        ElseIf Col = .ColIndex("GrandCOTotal") And HFApp.UserPermission("SalesOverride") Then
            Cancel = False
            .AutoSearch = flexSearchNone
        ElseIf Col = .ColIndex("Price") And HFApp.UserPermission("SalesOverride") Then
            Cancel = False
            .AutoSearch = flexSearchNone
        ElseIf Col = .ColIndex("CO_Price") And HFApp.UserPermission("SalesOverride") Then
            Cancel = False
            .AutoSearch = flexSearchNone
        ElseIf Col = .ColIndex("Other") Then
            Cancel = False
            .AutoSearch = flexSearchNone
            If .ValueMatrix(Row, .ColIndex("OtherListID")) <> 0 Then
                .ComboList = "|..."
            Else
                .ComboList = ""
            End If
        Else
            .ComboList = ""
            Cancel = True
            .AutoSearch = flexSearchFromCursor
        End If
    End With
End Sub

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
With gItems
        Select Case True
        
            Case .MouseRow = 0 And Button = vbRightButton
                Cancel = True

                
        
                
        End Select
End With
End Sub

Private Sub gItems_BeforeSort(ByVal Col As Long, Order As Integer)
gItems.FixedRows = 2
End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim s As String
    With gItems
        Select Case .ColKey(Col)
            Case "Comments"
                s = .Cell(flexcpText, Row, gItems.ColIndex("Comments"))
                If FComments.Edit(s, gItems, True, , 4000) Then
                    .Cell(flexcpText, Row, Col, .RowSel, Col) = s
                End If

            Case "Color"
                s = "select Value from AttributeListValues where ListID = " & .ValueMatrix(Row, .ColIndex("ColorListID"))
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Color Selection", s) Then
                    .Cell(flexcpText, Row, .ColIndex("Color"), .RowSel, .ColIndex("Color")) = FPickList.SelectedItem("Value")
                    .Cell(flexcpChecked, Row, .ColIndex("Pick"), .RowSel, .ColIndex("Pick")) = flexChecked
                End If
            Case "Location"
                s = "select a.Description,r.Model, r.Room from Rooms r" & vbCrLf
                s = s & "left outer join tblareas a on r.Room = a.Area" & vbCrLf
                s = s & " where Model =" & DbQuote(str, rs!Model) & vbCrLf
                s = s & "Union" & vbCrLf
                s = s & "select tblAreas.Description,DistinctModels.Model,tblAreas.Area from tblareas,DistinctModels" & vbCrLf
                s = s & " where Model =" & DbQuote(str, rs!Model)
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Location Selection", s, .TextMatrix(Row, Col), True, False, , "Model,Room") Then
                    .Cell(flexcpChecked, Row, .ColIndex("Pick"), .RowSel, .ColIndex("Pick")) = flexChecked
                    .Cell(flexcpText, Row, .ColIndex("Location"), .RowSel, .ColIndex("Location")) = FPickList.SelectedItem("Description")
                End If
            Case "Style"
                s = "select Value from AttributeListValues where ListID = " & .ValueMatrix(Row, .ColIndex("StyleListID"))
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Style Selection", s, .TextMatrix(Row, Col)) Then
                    .Cell(flexcpText, Row, .ColIndex("Style"), .RowSel, .ColIndex("Style")) = FPickList.SelectedItem("Value")
                    .Cell(flexcpChecked, Row, .ColIndex("Pick"), .RowSel, .ColIndex("Pick")) = flexChecked
                    If txtLocation.Text <> "" And .TextMatrix(Row, .ColIndex("Location")) = "" Then
                        .Cell(flexcpText, Row, .ColIndex("Location"), .RowSel, .ColIndex("Location")) = txtLocation.Text
                    End If
                End If
            Case "Finish"
               s = "select Value from AttributeListValues where ListID = " & .ValueMatrix(Row, .ColIndex("FinishListID"))
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Finish Selection", s, .TextMatrix(Row, Col)) Then
                    .Cell(flexcpText, Row, .ColIndex("Finish"), .RowSel, .ColIndex("Finish")) = FPickList.SelectedItem("Value")
                    .Cell(flexcpChecked, Row, .ColIndex("Pick"), .RowSel, .ColIndex("Pick")) = flexChecked
                    If txtLocation.Text <> "" And .TextMatrix(Row, .ColIndex("Location")) = "" Then
                        .Cell(flexcpText, Row, .ColIndex("Location"), .RowSel, .ColIndex("Location")) = txtLocation.Text
                    End If
                End If
            Case "Other"
                s = "select Value from AttributeListValues where ListID = " & .ValueMatrix(Row, .ColIndex("OtherListID"))
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Other Selection", s, .TextMatrix(Row, Col)) Then
                    .Cell(flexcpText, Row, .ColIndex("Other"), .RowSel, .ColIndex("Other")) = FPickList.SelectedItem("Value")
                    .Cell(flexcpChecked, Row, .ColIndex("Pick"), .RowSel, .ColIndex("Pick")) = flexChecked
                    If txtLocation.Text <> "" And .TextMatrix(Row, .ColIndex("Location")) = "" Then
                        .Cell(flexcpText, Row, .ColIndex("Location"), .RowSel, .ColIndex("Location")) = txtLocation.Text
                    End If
                End If
            Case Else
            Cancel = True
                        
        End Select
        
        
        Call gItems_AfterEdit(Row, Col)
        
    End With

End Sub

Public Sub GroupGrid()
On Error GoTo eh
    
    Dim i As Long
    Dim Row As Long
    Dim Col As Long
    
    
    
    With gItems
        Row = .Row
        Col = .Col
        

        
        GroupedColumns = 0
        For i = 0 To .cols - 1
            If HFApp.Options(UseMajorGroup) Then
                If i = gItems.ColIndex("MajDesc") Then
                    .ColPosition(i) = GroupedColumns + 2
                    GroupedColumns = GroupedColumns + 1
                End If
            End If
            If i = gItems.ColIndex("CatDesc") Then
                   
                    .ColPosition(i) = GroupedColumns + 2
                    GroupedColumns = GroupedColumns + 1
            End If
        Next
        
        If GroupedColumns = 0 Then
            Call .SubTotal(flexSTClear)
        Else
            .SubTotal flexSTClear
            'mFunnyFlag = True
            .Col = 2
            .ColSel = 1 + GroupedColumns
            .Sort = flexSortGenericAscending
            
            
            .OutlineCol = 2
            .SubtotalPosition = flexSTAbove
            For i = 2 To GroupedColumns + 1
                .SubTotal flexSTCount, i, .ColIndex("OPT"), "$(#,###.00)", , vbHighlight, True, "%s", 0
                
            Next
        End If
    
        On Error Resume Next
        .Row = Row
        .Col = Col
    End With
    'Call ColorizeItems(0)
    If gItems.Cell(flexcpTextDisplay, 2, gItems.ColIndex("Description")) = "" Then
        gItems.RowHidden(2) = True
        If HFApp.Options(UseMajorGroup) Then
            gItems.RowHidden(3) = True
        End If
    End If
Exit Sub
eh:
MsgBox Err.Description

 
Call errHandler(Err.Number & " GroupGrid" & " " & Err.Description & "frmOptionListMultiSelect")

End Sub

Private Sub gitems_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, done As Boolean)
    With gItems
    If Row > 1 Then
        If HFApp.Options(UseMajorGroup) Then
            
            If Col = .ColIndex("MajDesc") And Row > .FixedRows And Not .IsSubtotal(Row) Then
                done = True
            End If
            If Col = .ColIndex("CatDesc") And Row > .FixedRows And Not .IsSubtotal(Row) Then
                done = True
            End If
        Else
            If Col = .ColIndex("CatDesc") And Row > .FixedRows And Not .IsSubtotal(Row) Then
                done = True
            End If
        End If
        If HFApp.Options(UseMajorGroup) Then
            If .IsSubtotal(Row) And Col = .ColIndex("MajDesc") And .RowOutlineLevel(Row) > 2 And Col < .RowOutlineLevel(Row) Then
                done = True
            End If
        End If
    End If
    End With
End Sub

Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gItems
        If .Row = 1 And .RowSel = 1 And KeyCode = vbKeyDelete Then
            .Text = ""
            For i = 2 To .Rows - 1
                .RowHidden(i) = False
            Next
        End If
    End With
End Sub

Private Sub gitems_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
 With gItems
        If Col = .ColIndex("Color") Then
            If .Cell(flexcpChecked, Row, .ColIndex("ColorStrict")) = flexChecked Then
                MsgBox "Selection must be made from the list", vbCritical + vbOKOnly, "Color Selection"
                Cancel = True
            End If
        ElseIf Col = .ColIndex("Style") Then
            If .Cell(flexcpChecked, Row, .ColIndex("StyleStrict")) = flexChecked Then
                MsgBox "Selection must be made from the list", vbCritical + vbOKOnly, "Style Selection"
                Cancel = True
            End If
        ElseIf Col = .ColIndex("Finish") Then
            If .Cell(flexcpChecked, Row, .ColIndex("FinishStrict")) = flexChecked Then
                MsgBox "Selection must be made from the list", vbCritical + vbOKOnly, "Finish Selection"
                Cancel = True
            End If
        ElseIf Col = .ColIndex("Other") Then
            If .Cell(flexcpChecked, Row, .ColIndex("OtherStrict")) = flexChecked Then
                MsgBox "Selection must be made from the list", vbCritical + vbOKOnly, "Other Selection"
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub gItems_RowColChange()
On Error Resume Next

'    Command6.Enabled = .TextMatrix(.Row, .ColIndex("Graphic_Path")) <> ""
With gItems
    If .TextMatrix(.Row, .ColIndex("Graphic_Path")) <> "" Then
       If FileExists(.TextMatrix(.Row, .ColIndex("Graphic_Path"))) Then
          If LCase(Right(.TextMatrix(.Row, .ColIndex("Graphic_Path")), 3)) = "pdf" Then
            ImageViewer1.FileName = App.Path & "\pdf.jpg"
            'ImageViewer1.filename = (.TextMatrix(.Row, .ColIndex("Graphic_Path")))
            ImageViewer1.view = 9
            
                'StartDoc .TextMatrix(.Row, .ColIndex("Graphic_Path"))
          Else
            ImageViewer1.FileName = (.TextMatrix(.Row, .ColIndex("Graphic_Path")))
            ImageViewer1.FileLeft = 0
            ImageViewer1.view = 9
            
          End If
          ImageViewer1.Visible = True
       End If
       gItems.SetFocus
    Else
        ImageViewer1.Visible = False
    End If

End With

End Sub

Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Count As Long
With gItems
Select Case .ColKey(Col)
Case "Color"
    If .Cell(flexcpChecked, r, .ColIndex("ColorRequired")) = flexChecked Then
        Count = HFApp.SqlExec("Select count(*) from AttributeListValues where ListID = " & DbQuote(Num, .ValueMatrix(r, .ColIndex("ColorListID"))) & " and Value = " & DbQuote(str, .TextMatrix(Row, .ColIndex("Color"))))(0)
        If Count = 0 Then
           MsgBox "Invalid Selection. Please choose again."
           Cancel = True
        End If
    End If
Case "Style"
Case "Finish"
Case "Other"
End Select
End With
End Sub

Private Sub ImageViewer1_Click()
On Error Resume Next
If ImageViewer1.FileName = App.Path & "\pdf.jpg" Then
    Call ShellFile(Me.hwnd, (gItems.TextMatrix(gItems.Row, gItems.ColIndex("Graphic_path"))))
End If
End Sub

Private Sub ImageViewer1_DblClick()
On Error Resume Next
'If ImageViewer1.filename = App.Path & "\pdf.jpg" Then
    Call ShellFile(Me.hwnd, (gItems.TextMatrix(gItems.Row, gItems.ColIndex("Graphic_path"))))
'End If
End Sub

Private Sub Option1_Click()
On Error GoTo herror
Dim s As String
Screen.MousePointer = vbHourglass

curPick.RecordSource = ""
curPick.CursorType = adOpenStatic

    'Changed by Fazlul August29, 2005
    
       If HFApp.Options(UseMajorGroup) = True Then
          s = ""
          s = s & "SELECT T1.Area ,T1.Model ,T1.Series, T1.Option_Type " & vbCrLf
          s = s & ",T1.Option_Type_Desc ,T1.OPT , t1.Description, t1.Construction_Cut_Off" & vbCrLf
          s = s & ",t1.D_Structural_Change,T1.Cost_Amount ,T1.Price ,T1.TL_Extra" & vbCrLf
          s = s & ",case when isnull(t1.co_price,0) < t1.price then t1.price else t1.co_price end CO_Price ,T1.Category ,T1.Comments , t1.UOM, t1.Major_Group" & vbCrLf
          s = s & ",T1.Price as BasePrice,case when isnull(t1.co_price,0) < t1.price then t1.price else t1.co_price end BaseCO_Price" & vbCrLf
          s = s & ",t1.NoCommission , t1.French_Desc, t1.French_Comments , t1.graphic_path" & vbCrLf
          s = s & ",t1.Qty,T1.Pick ,T1.WebUpdated ,T1.NoDataChanged , t1.Web_View, t1.EnerGuideRating" & vbCrLf
          s = s & ",t1.GreenHouseRating, t1.FuelCost , t1.UseTax, t1.NetTax, t1.TotalAmount " & vbCrLf
          s = s & ",t1.NetCOTax, t1.TotalCOAmount ,T2.Description as CatDesc ,T3.Description as MajDesc" & vbCrLf
          s = s & ",isnull(T1.Price,0)+isnull(T1.NetTAX,0) as GrandTotal,isnull(T1.CO_Price,0)+isnull(T1.NetCOTAX,0) as GrandCOTotal" & vbCrLf
          s = s & ",T1.SalesWorkSheet,T1.Assembly,T1.ApplyIncentive,T1.EstimatorNotes,t1.Pick as NoCharge,t4.Description Location,t1.Color,t1.ColorListID,t1.StyleListID,t1.FinishListID,t1.OtherListID,t1.Style,t1.Finish,t1.Other,l1.Required ColorRequired,l2.Required StyleRequired,l3.Required FinishRequired,l4.Required OtherRequired,l1.StrictList ColorStrict,l2.StrictList StyleStrict,l3.StrictList FinishStrict,l4.StrictList OtherStrict  FROM tblOptions t1 LEFT OUTER JOIN tblcategories T2 ON T1.Category = T2.Category " & vbCrLf
          s = s & "LEFT OUTER JOIN attributelists l1 on T1.ColorListid=l1.Listid" & vbCrLf
          s = s & "LEFT OUTER JOIN attributelists l2 on T1.styleListid=l2.Listid" & vbCrLf
          s = s & "LEFT OUTER JOIN attributelists l3 on T1.FinishListid=l3.Listid" & vbCrLf
          s = s & "LEFT OUTER JOIN attributelists l4 on T1.OtherListid=l4.Listid" & vbCrLf
          s = s & "LEFT OUTER JOIN TBLAREAS T4 ON T1.Location = T4.Area " & vbCrLf
          s = s & "LEFT OUTER JOIN tblMajorGroups T3 ON T2.Group_Code = T3.Major_Group where isnull(t1.inactive,0) = 0 and T1.model='" & rs!Model & "'" & vbCrLf
          s = s & " and T1.Series='" & rs!series & "'" & vbCrLf
          s = s & " and (T1.Elevation = " & DbQuote(str, rs!Elevation) & " or t1.elevation ='' or t1.elevation is null)" & vbCrLf
          If HFApp.Options(OptionByArea) Or HFApp.Options(OptionByAreaPhase) Then
                s = s & " and (T1.Area = dbo.Sales_GetOptionCommunity(Model,OPT,'" & rs!community & "','tblOptions') or T1.Area = '')"
                
                If HFApp.Options(OptionByAreaPhase) Then
                    s = s & " and (t1.communityphase ='" & rs!Phase & "' or t1.communityphase ='' or t1.communityphase is null)" & vbCrLf
                Else
                    s = s & " and (t1.communityphase ='' or t1.communityphase is null)" & vbCrLf
                End If
          Else
             s = s & " and (T1.Area='' or t1.Area is null or t1.Area = 'ALL AREA' or t1.Area = 'ALL AREAS')" & vbCrLf
          End If
          s = s & " and t1.DivisionID = " & HFApp.DivisionID & vbCrLf
          s = s & " and (T1.Construction_Cut_Off>0" & vbCrLf
          s = s & " or T1.Construction_Cut_Off=0) Order by T3.Description,T2.Description,T1.Description"
          If HFApp.Options(OptionByAreaPhase) Then
                s = s & ",T1.Communityphase"
          End If
          curPick.RecordSource = s
       Else
          s = ""
          s = s & "SELECT T1.Area ,T1.Model ,T1.Series, T1.Option_Type " & vbCrLf
          s = s & ",T1.Option_Type_Desc ,T1.OPT , t1.Description, t1.Construction_Cut_Off" & vbCrLf

          s = s & ",t1.D_Structural_Change,T1.Cost_Amount ,T1.Price ,T1.TL_Extra" & vbCrLf
          s = s & ",case when isnull(t1.co_price,0) < t1.price then t1.price else t1.co_price end CO_Price ,T1.Category ,T1.Comments , t1.UOM, t1.Major_Group" & vbCrLf
          s = s & ",T1.Price as BasePrice,case when isnull(t1.co_price,0) < t1.price then t1.price else t1.co_price end BaseCO_Price" & vbCrLf
          s = s & ",t1.NoCommission , t1.French_Desc, t1.French_Comments , t1.graphic_path" & vbCrLf
          s = s & ",t1.Qty,T1.Pick ,T1.WebUpdated ,T1.NoDataChanged , t1.Web_View, t1.EnerGuideRating" & vbCrLf
          s = s & ",t1.GreenHouseRating, t1.FuelCost , t1.UseTax, t1.NetTax, t1.TotalAmount" & vbCrLf
          s = s & ",t1.NetCOTax, t1.TotalCOAmount ,T2.Description as CatDesc " & vbCrLf
          s = s & ",isnull(T1.Price,0)+isnull(T1.NetTAX,0) as GrandTotal,isnull(T1.CO_Price,0)+isnull(T1.NetCOTAX,0) as GrandCOTotal" & vbCrLf
          s = s & ",T1.SalesWorkSheet,T1.Assembly,T1.ApplyIncentive,T1.EstimatorNotes,t1.Pick as NoCharge,t4.Description Location,t1.Color,t1.ColorListID,t1.StyleListID,t1.FinishListID,t1.OtherListID,t1.Style,t1.Finish,t1.Other,l1.Required ColorRequired,l2.Required StyleRequired,l3.Required FinishRequired,l4.Required OtherRequired,l1.StrictList ColorStrict,l2.StrictList StyleStrict,l3.StrictList FinishStrict,l4.StrictList OtherStrict  FROM tblOptions t1 LEFT OUTER JOIN tblcategories T2 ON T1.Category = T2.Category " & vbCrLf
          s = s & "LEFT OUTER JOIN attributelists l1 on T1.ColorListid=l1.Listid" & vbCrLf
          s = s & "LEFT OUTER JOIN attributelists l2 on T1.styleListid=l2.Listid" & vbCrLf
          s = s & "LEFT OUTER JOIN attributelists l3 on T1.FinishListid=l3.Listid" & vbCrLf
          s = s & "LEFT OUTER JOIN attributelists l4 on T1.OtherListid=l4.Listid" & vbCrLf
          s = s & "LEFT OUTER JOIN TBLAREAS T4 ON T1.Location = T4.Area" & vbCrLf
          s = s & " where isnull(t1.inactive,0) = 0 and T1.model='" & gModel & "' " & vbCrLf
          s = s & "and T1.Series='" & gSeries & "' " & vbCrLf
          s = s & "and (T1.Elevation = " & DbQuote(str, "" & rs!Elevation) & " or t1.elevation ='' or t1.elevation is null) " & vbCrLf
          If HFApp.Options(OptionByArea) Or HFApp.Options(OptionByAreaPhase) Then
            
             s = s & " and (T1.Area = dbo.Sales_GetOptionCommunity(Model,OPT,'" & "" & rs!community & "','tblOptions'))"
                If HFApp.Options(OptionByAreaPhase) Then
                    s = s & " and (t1.communityphase ='" & rs!Phase & "' or isnull(t1.communityphase,'') ='')" & vbCrLf
                Else
                    s = s & " and (t1.communityphase ='' or t1.communityphase is null)" & vbCrLf
                End If
          Else
             s = s & " and (T1.Area='' or T1.Area is null or t1.Area = 'ALL AREA' or t1.Area = 'ALL AREAS')" & vbCrLf
          End If
          s = s & " and t1.DivisionID = " & HFApp.DivisionID & vbCrLf
          s = s & "and (T1.Construction_Cut_Off>0" & vbCrLf
          s = s & " or T1.Construction_Cut_Off=0) Order by T2.Description,T1.Description"
          If HFApp.Options(OptionByAreaPhase) Then
                s = s & ",T1.Communityphase"
          End If
          'Order by T3.Description,T2.Description,T1.Description
          curPick.RecordSource = s
       End If
    
curPick.Refresh

'Set curTemp = Clone(curPick.Recordset)

Call LoadGrid
Exit Sub

herror:
If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        Call errHandler(Err.Number & " Option 1 build recordset " & Err.Description & Me.Caption)
        'Resume
        Call cmdCancel_Click
        
End If
End Sub

Private Sub LoadGrid()
On Error GoTo herror
Dim X As Integer

gItems.Editable = flexEDKbdMouse
gItems.Rows = 1
gItems.cols = 1
Set gItems.DataSource = curPick
gItems.ColWidth(gItems.ColIndex("Comments")) = 5000
gItems.ColWidth(gItems.ColIndex("Description")) = 10000
Call IniGetGrid(Me, gItems)
Call gItems.AddItem("", 1)
gItems.Cell(flexcpBackColor, 1, 0, 1, gItems.cols - 1) = vbInfoBackground
Me.ImageViewer1.Visible = False
On Error Resume Next
'gItems.ColData(gItems.ColIndex("MajDesc")) = "GROUPED"
'gItems.ColData(gItems.ColIndex("CatDesc")) = "GROUPED"
On Error GoTo herror
For X = 0 To curPick.Recordset.fields.Count - 1
    With gItems

            If curPick.Recordset.fields(X).Name <> "Color" And curPick.Recordset.fields(X).Name <> "Location" And curPick.Recordset.fields(X).Name <> "Style" And curPick.Recordset.fields(X).Name <> "Finish" And curPick.Recordset.fields(X).Name <> "Other" And curPick.Recordset.fields(X).Name <> "UOM" And curPick.Recordset.fields(X).Name <> "Description" And curPick.Recordset.fields(X).Name <> "CatDesc" And curPick.Recordset.fields(X).Name <> "MajDesc" And curPick.Recordset.fields(X).Name <> "Comments" And curPick.Recordset.fields(X).Name <> "Qty" And curPick.Recordset.fields(X).Name <> "Pick" And curPick.Recordset.fields(X).Name <> "EstimatorNotes" Then
                If curPick.Recordset.fields(X).Name = "Price" Or curPick.Recordset.fields(X).Name = "CO_Price" Or curPick.Recordset.fields(X).Name = "GrandCOTotal" Or curPick.Recordset.fields(X).Name = "GrandTotal" Then

                    .ColFormat(.ColIndex(curPick.Recordset.fields(X).Name)) = "#,###,###.##"
                    
                        If curPick.Recordset.fields(X).Name <> "Price" Then
                           .ColHidden(.ColIndex(curPick.Recordset.fields(X).Name)) = curPick.Recordset.fields(X).Name <> "Price"
                        Else
                           .ColHidden(.ColIndex(curPick.Recordset.fields(X).Name)) = False
                           
                        End If
                Else
                    .ColHidden(.ColIndex(curPick.Recordset.fields(X).Name)) = True
                    .TextMatrix(0, .ColIndex(curPick.Recordset.fields(X).Name)) = ""
                End If
            End If
    End With
Next
gItems.ColHidden(gItems.ColIndex("EstimatorNotes")) = False
gItems.FrozenRows = 1
X = 1
gItems.ColPosition(gItems.ColIndex("Pick")) = X
X = 2
X = X + 1
gItems.ColPosition(gItems.ColIndex("Qty")) = X
X = X + 1
gItems.ColPosition(gItems.ColIndex("UOM")) = X
    X = X + 1
    gItems.ColPosition(gItems.ColIndex("Price")) = X
    X = X + 1


If HFApp.Options(UseMajorGroup) Then
    gItems.ColPosition(gItems.ColIndex("MajDesc")) = X
    gItems.TextMatrix(0, gItems.ColIndex("MajDesc")) = "Group"
    X = X + 1
End If
gItems.ColPosition(gItems.ColIndex("CatDesc")) = X
X = X + 1
gItems.TextMatrix(0, gItems.ColIndex("CatDesc")) = "Category"
gItems.ColPosition(gItems.ColIndex("Description")) = X



'gItems.AutoSizeMode = flexAutoSizeRowHeight
gItems.RowHeightMax = 2
'If HFApp.Options(Show_CDN_GST) = False Then
    gItems.FrozenCols = 6
'Else
'    gItems.FrozenCols = 5
'End If

'gItems.AutoResize = True

gItems.WordWrap = True
gItems.Visible = True
Screen.MousePointer = vbDefault
Call GroupGrid
Call ColorizeItems(0)
Exit Sub

herror:
If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        Call errHandler(Err.Number & " Load Grid " & Err.Description)
End If

End Sub

Private Sub Option2_Click()
On Error Resume Next
Dim s As String
Screen.MousePointer = vbHourglass
curPick.CursorType = adOpenDynamic
curPick.RecordSource = ""
'Changed by Fazlul August29, 2005

   If HFApp.Options(UseMajorGroup) = True Then
      s = ""
      s = s & "SELECT T1.Option_Type ,T1.Option_Type_Desc , t1.OPT, t1.Description" & vbCrLf
      s = s & ",T1.Construction_Cut_Off,T1.D_Structural_Change,t1.Cost_Amount,t1.Price" & vbCrLf
      s = s & ",t1.TL_Extra,t1.Category,case when isnull(t1.co_price,0) < t1.price then t1.price else t1.co_price end CO_Price,t1.Comments,t1.UOM,t1.Major_Group,t1.NoCommission" & vbCrLf
      s = s & ",t1.French_Desc,t1.French_Comments,t1.graphic_path,t1.Qty,t1.Pick,t1.WebUpdated" & vbCrLf
      s = s & ",t1.NoDataChanged,t1.Web_View,t1.EnerGuideRating,t1.GreenHouseRating,t1.FuelCost" & vbCrLf
      s = s & ",t1.UseTax,t1.NetTAX,t1.TotalAmount,t1.NetCOTax,t1.TotalCOAmount,t2.Description as CatDesc" & vbCrLf
      s = s & ",T3.Description as MajDesc,isnull(T1.Price,0)+isnull(T1.NetTAX,0) as GrandTotal" & vbCrLf
      s = s & ",isnull(T1.CO_Price,0)+isnull(T1.NetCOTAX,0) as GrandCOTotal,T1.SalesWorkSheet" & vbCrLf
      s = s & ",T1.Price as BasePrice,case when isnull(t1.co_price,0) < t1.price then t1.price else t1.co_price end BaseCO_Price" & vbCrLf
      s = s & ",T1.Assembly,T1.ApplyIncentive,T1.EstimatorNotes,t1.Pick as NoCharge,t4.Description Location,t1.Color,t1.ColorListID,t1.StyleListID,t1.FinishListID,t1.OtherListID,t1.Style,t1.Finish,t1.Other,l1.Required ColorRequired,l2.Required StyleRequired,l3.Required FinishRequired,l4.Required OtherRequired,l1.StrictList ColorStrict,l2.StrictList StyleStrict,l3.StrictList FinishStrict,l4.StrictList OtherStrict  FROM tblGlobalOptions T1 LEFT Join tblcategories T2 " & vbCrLf
      s = s & "ON T1.Category = T2.Category " & vbCrLf
      s = s & "LEFT OUTER JOIN attributelists l1 on T1.ColorListid=l1.Listid" & vbCrLf
      s = s & "LEFT OUTER JOIN attributelists l2 on T1.styleListid=l2.Listid" & vbCrLf
      s = s & "LEFT OUTER JOIN attributelists l3 on T1.FinishListid=l3.Listid" & vbCrLf
      s = s & "LEFT OUTER JOIN attributelists l4 on T1.OtherListid=l4.Listid" & vbCrLf
      s = s & "LEFT OUTER JOIN TBLAREAS T4 ON T1.Location = T4.Area " & vbCrLf
      s = s & "LEFT JOIN tblMajorGroups T3 ON T2.Group_Code = T3.Major_Group " & vbCrLf
      s = s & "where isnull(t1.inactive,0) = 0 and (T1.Construction_Cut_Off>0"
      s = s & " or T1.Construction_Cut_Off=0) " & vbCrLf
      If HFApp.Options(GlobalOptionByArea) Or HFApp.Options(GlobalOptionByAreaPhase) Then
            
            s = s & " and (t1.community = dbo.Sales_GetOptionCommunity('',OPT,'" & rs!community & "','tblGlobalOptions'))"
            If HFApp.Options(GlobalOptionByAreaPhase) Then
                s = s & " and (t1.communityphase ='" & rs!Phase & "' or isnull(t1.communityphase,'') ='')" & vbCrLf
            Else
                s = s & " and (t1.communityphase ='' or t1.communityphase is null)" & vbCrLf
            End If
      Else
            s = s & " and (T1.Community='' or t1.Community is null or t1.Community = 'ALL AREA')" & vbCrLf
      End If
      s = s & " and t1.DivisionID = " & HFApp.DivisionID & vbCrLf
      s = s & "Order by T3.Description,T2.Description,T1.Description"
      If HFApp.Options(OptionByAreaPhase) Then
            s = s & ",T1.Communityphase"
      End If
      curPick.RecordSource = s
   Else
      s = ""
      s = s & "SELECT T1.Option_Type ,T1.Option_Type_Desc , t1.OPT, t1.Description" & vbCrLf
      s = s & ",T1.Construction_Cut_Off,T1.D_Structural_Change,t1.Cost_Amount,t1.Price" & vbCrLf
      s = s & ",t1.TL_Extra,t1.Category,case when isnull(t1.co_price,0) < t1.price then t1.price else t1.co_price end CO_Price,t1.Comments,t1.UOM,t1.Major_Group,t1.NoCommission" & vbCrLf
      s = s & ",t1.French_Desc,t1.French_Comments,t1.graphic_path,t1.Qty,t1.Pick,t1.WebUpdated" & vbCrLf
      s = s & ",t1.NoDataChanged,t1.Web_View,t1.EnerGuideRating,t1.GreenHouseRating,t1.FuelCost" & vbCrLf
      s = s & ",t1.UseTax,t1.NetTAX,t1.TotalAmount,t1.NetCOTax,t1.TotalCOAmount,t2.Description as CatDesc" & vbCrLf
      s = s & ",isnull(T1.Price,0)+isnull(T1.NetTAX,0) as GrandTotal" & vbCrLf
      s = s & ",T1.Price as BasePrice,case when isnull(t1.co_price,0) < t1.price then t1.price else t1.co_price end BaseCO_Price" & vbCrLf
      s = s & ",isnull(T1.CO_Price,0)+isnull(T1.NetCOTAX,0) as GrandCOTotal,T1.SalesWorkSheet" & vbCrLf
      s = s & ",T1.Assembly,T1.ApplyIncentive,T1.EstimatorNotes,t1.Pick as NoCharge,t4.Description Location,t1.Color,t1.ColorListID,t1.StyleListID,t1.FinishListID,t1.OtherListID,t1.Style,t1.Finish,t1.Other,l1.Required ColorRequired,l2.Required StyleRequired,l3.Required FinishRequired,l4.Required OtherRequired,l1.StrictList ColorStrict,l2.StrictList StyleStrict,l3.StrictList FinishStrict,l4.StrictList OtherStrict  FROM tblGlobalOptions T1 "
      s = s & "LEFT Join tblcategories T2 ON T1.Category = T2.Category " & vbCrLf
      s = s & "LEFT OUTER JOIN attributelists l1 on T1.ColorListid=l1.Listid" & vbCrLf
      s = s & "LEFT OUTER JOIN attributelists l2 on T1.styleListid=l2.Listid" & vbCrLf
      s = s & "LEFT OUTER JOIN attributelists l3 on T1.FinishListid=l3.Listid" & vbCrLf
      s = s & "LEFT OUTER JOIN attributelists l4 on T1.OtherListid=l4.Listid" & vbCrLf

      s = s & "LEFT OUTER JOIN TBLAREAS T4 ON T1.Location = T4.Area " & vbCrLf
      s = s & "where isnull(t1.inactive,0) = 0 and (T1.Construction_Cut_Off>0"
      s = s & " or T1.Construction_Cut_Off=0) " & vbCrLf
      If HFApp.Options(GlobalOptionByArea) Or HFApp.Options(GlobalOptionByAreaPhase) Then

             s = s & " and (t1.community = dbo.Sales_GetOptionCommunity('',OPT,'" & rs!community & "','tblGlobalOptions'))"
            If HFApp.Options(GlobalOptionByAreaPhase) Then
                s = s & " and (t1.communityphase ='" & rs!Phase & "' or isnull(t1.communityphase,'') ='')" & vbCrLf
            Else
                s = s & " and (t1.communityphase ='' or t1.communityphase is null)" & vbCrLf
            End If
      Else
            s = s & " and (T1.Community='' or t1.community is null or t1.Community = 'ALL AREA')" & vbCrLf
      End If
      s = s & " and t1.DivisionID = " & HFApp.DivisionID & vbCrLf
      s = s & "Order by T2.Description,T1.Description"
      If HFApp.Options(OptionByAreaPhase) Then
            s = s & ",T1.Communityphase"
      End If
      curPick.RecordSource = s
    End If

curPick.Refresh

'Set curTemp = Clone(curPick.Recordset)


Screen.MousePointer = vbDefault

If curPick.Recordset.RecordCount = 0 Then
   Command1.Enabled = False
   Command2.Enabled = False
   cmdSelect.Enabled = False
Else
   Command1.Enabled = True
   Command2.Enabled = True
   cmdSelect.Enabled = True
End If

Call LoadGrid
Exit Sub

herror:
If Err.Number <> 0 Then
   Screen.MousePointer = vbDefault
   Call errHandler(Err.Number & " " & Err.Description)
End If
End Sub

Private Sub Option3_Click()
On Error Resume Next
Dim s As String
Screen.MousePointer = vbHourglass
curPick.CursorType = adOpenStatic
curPick.RecordSource = ""
Dim doUnion As Boolean

If HFApp.Options(UseMajorGroup) = True Then
        's = "select * from CombinedOptionList "
        s = "select Community, Model, Series, Option_Type, Option_Type_Desc, OPT, CombinedOptionList.Description, Construction_Cut_Off, D_Structural_Change, Cost_Amount, Price, Price as BasePrice, TL_Extra, CO_Price,CO_Price as BaseCO_Price, Category, Comments, UOM, Major_Group, NoCommission, French_Desc, French_Comments, graphic_path, Qty, Pick, WebUpdated, NoDataChanged, Web_View, EnerGuideRating, GreenHouseRating, FuelCost, UseTax, NetTax, TotalAmount, NetCOTax, TotalCOAmount, CatDesc, MajDesc, GrandTotal, GrandCOTotal, SalesWorkSheet, Assembly, ApplyIncentive, EstimatorNotes, NoCharge, t4.Description Location, Color, IncludedOption, inactive, communityphase, Elevation,ColorListID,StyleListID,FinishListID,OtherListID,Style,Finish,Other,l1.Required ColorRequired,l2.Required StyleRequired,l3.Required FinishRequired,l4.Required OtherRequired,l1.StrictList ColorStrict,l2.StrictList StyleStrict,l3.StrictList FinishStrict,l4.StrictList OtherStrict  from CombinedOptionList " & vbCrLf
        s = s & "LEFT OUTER JOIN attributelists l1 on combinedoptionlist.ColorListid=l1.Listid" & vbCrLf
        s = s & "LEFT OUTER JOIN attributelists l2 on combinedoptionlist.styleListid=l2.Listid" & vbCrLf
        s = s & "LEFT OUTER JOIN attributelists l3 on combinedoptionlist.FinishListid=l3.Listid" & vbCrLf
        s = s & "LEFT OUTER JOIN attributelists l4 on combinedoptionlist.OtherListid=l4.Listid" & vbCrLf
        s = s & "LEFT OUTER JOIN TBLAREAS T4 ON CombinedOptionList.Location = T4.Area " & vbCrLf
        s = s & "where isnull(inactive,0) = 0 and (isnull(model,'')='" & gModel & "' or isnull(Model,'')='')" & vbCrLf
        s = s & " and (isnull(Series,'')='" & gSeries & "' or isnull(Series,'') = '')" & vbCrLf
        s = s & " and (Elevation = " & DbQuote(str, rs!Elevation) & " or isnull(elevation,'') ='')" & vbCrLf
        If HFApp.Options(OptionByArea) Or HFApp.Options(OptionByAreaPhase) Then
            s = s & " and (Community = dbo.Sales_GetOptionCommunity(Model,OPT,'" & rs!community & "','CombinedOptionlist'))"
           
              If HFApp.Options(OptionByAreaPhase) Then
                  s = s & " and (communityphase ='" & rs!Phase & "' or isnull(communityphase,'') ='')" & vbCrLf
              Else
                  s = s & " and (communityphase ='' or communityphase is null)" & vbCrLf
              End If
        Else
           s = s & " and (Community='' or Community is null or Community = 'ALL AREA' or Community = 'ALL AREAS')" & vbCrLf
        End If
        s = s & " and (Construction_Cut_Off>0" & vbCrLf
        s = s & " or Construction_Cut_Off=0)"
        s = s & " and CombinedOptionList.DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & " Order by CatDesc,CombinedOptionList.Description"
        If HFApp.Options(OptionByAreaPhase) Then
              s = s & ",Communityphase"
        End If
Else
    s = ""
    s = s & "SELECT Community ,Model ,Series, Option_Type " & vbCrLf
    s = s & ",Option_Type_Desc ,OPT , CombinedOptionList.Description, Construction_Cut_Off" & vbCrLf
    s = s & ",D_Structural_Change,Cost_Amount ,Price ,TL_Extra" & vbCrLf
    s = s & ",CO_Price ,Category ,Comments , UOM, Major_Group" & vbCrLf
    s = s & ",NoCommission , French_Desc, French_Comments , graphic_path" & vbCrLf
    s = s & ",Qty,Pick ,WebUpdated ,NoDataChanged , Web_View, EnerGuideRating" & vbCrLf
    s = s & ",GreenHouseRating, FuelCost , UseTax, NetTax, TotalAmount" & vbCrLf
    s = s & ",NetCOTax, TotalCOAmount ,CatDesc " & vbCrLf
    s = s & ",GrandTotal,GrandCOTotal" & vbCrLf
    s = s & ",Price as BasePrice,Co_price as BaseCO_Price" & vbCrLf
    s = s & ",SalesWorkSheet,Assembly,ApplyIncentive,EstimatorNotes,NoCharge,t4.Description Location,Color,ColorListID,StyleListID,FinishListID,OtherListID,Style,Finish,Other,l1.Required ColorRequired,l2.Required StyleRequired,l3.Required FinishRequired,l4.Required OtherRequired,l1.StrictList ColorStrict,l2.StrictList StyleStrict,l3.StrictList FinishStrict,l4.StrictList OtherStrict  FROM CombinedOptionList" & vbCrLf
    s = s & "LEFT OUTER JOIN attributelists l1 on combinedoptionlist.ColorListid=l1.Listid" & vbCrLf
    s = s & "LEFT OUTER JOIN attributelists l2 on combinedoptionlist.styleListid=l2.Listid" & vbCrLf
    s = s & "LEFT OUTER JOIN attributelists l3 on combinedoptionlist.FinishListid=l3.Listid" & vbCrLf
    s = s & "LEFT OUTER JOIN attributelists l4 on combinedoptionlist.OtherListid=l4.Listid" & vbCrLf
   s = s & "LEFT OUTER JOIN TBLAREAS T4 ON CombinedOptionList.Location = T4.Area " & vbCrLf
   s = s & "where isnull(inactive,0) = 0 and (isnull(model,'')='" & gModel & "' or isnull(Model,'')='')" & vbCrLf
   s = s & " and (isnull(Series,'')='" & gSeries & "' or isnull(Series,'') = '')" & vbCrLf
   s = s & " and (Elevation = " & DbQuote(str, rs!Elevation) & " or isnull(elevation,'') ='')" & vbCrLf
   If HFApp.Options(OptionByArea) Or HFApp.Options(OptionByAreaPhase) Then
       
         s = s & " and (Community = dbo.Sales_GetOptionCommunity(Model,OPT,'" & rs!community & "','CombinedOptionlist'))"
         If HFApp.Options(OptionByAreaPhase) Then
             s = s & " and (communityphase ='" & rs!Phase & "' or isnull(communityphase,'') ='')" & vbCrLf
         Else
             s = s & " and (communityphase ='' or communityphase is null)" & vbCrLf
         End If
   Else
      s = s & " and (Community='' or Community is null or Community = 'ALL AREA' or Community = 'ALL AREAS')" & vbCrLf
   End If
   s = s & "and (Construction_Cut_Off>0" & vbCrLf
   s = s & " or Construction_Cut_Off=0)"
   s = s & " and CombinedOptionList.DivisionID = " & HFApp.DivisionID & vbCrLf
   s = s & " Order by CatDesc,CombinedOptionList.Description"
   If HFApp.Options(OptionByAreaPhase) Then
         s = s & ",Communityphase"
   End If
End If



curPick.RecordSource = s
curPick.Refresh

'Set curTemp = Clone(curPick.Recordset)


'MsgBox "Cursortype = " & curTemp.CursorType
'MsgBox "Cursorlocation = " & curTemp.CursorLocation

Screen.MousePointer = vbDefault

If curPick.Recordset.RecordCount = 0 Then
   Command1.Enabled = False
   Command2.Enabled = False
   cmdSelect.Enabled = False
Else
'   Command1.Enabled = True
'   Command2.Enabled = True
   cmdSelect.Enabled = True
End If
Call LoadGrid
Exit Sub

herror:
If Err.Number <> 0 Then
   Screen.MousePointer = vbDefault
   Call errHandler(Err.Number & " " & Err.Description)
End If

End Sub

Private Sub Option4_Click()
On Error Resume Next
Dim s As String
Screen.MousePointer = vbHourglass
curPick.CursorType = adOpenDynamic
curPick.RecordSource = ""
s = "select c.Customer_No,c.Location as Option_type,l.Category,l.Major_group,l.MajDesc,l.CatDesc, c.OptionID as OPT,c.Description, c.Qty, c.Rate as Price,c.Rate as BasePrice, l.GrandTotal" & vbCrLf
s = s & ",l.GrandCoTotal, l.Graphic_path , c.Comments, c.OptionListingSeq,l.Pick,l.French_desc,l.French_Comments,0 as NoCharge" & vbCrLf
s = s & ",l.ApplyIncentive, l.Cost_Amount,l.Assembly,l.NetTax,l.UseTax,l.CO_Price,l.CO_Price as BaseCO_Price,l.D_Structural_Change,l.TL_Extra,l.NoCommission,l.EstimatorNotes,l.UOM,l.SalesWorksheet,c.Seq" & vbCrLf
s = s & ",l.EnerGuideRating,l.GreenHouseRating,l.FuelCost,T4.Description Location,l.Color,l.ColorListID,l.StyleListID,l.FinishListID,l.OtherListID,l.Style,l.Finish,l.Other,l1.Required ColorRequired,l2.Required StyleRequired,l3.Required FinishRequired,l4.Required OtherRequired,l1.StrictList ColorStrict,l2.StrictList StyleStrict,l3.StrictList FinishStrict,l4.StrictList OtherStrict from CustomerShoppingCart c" & vbCrLf
s = s & "join combinedoptionlist l on (c.location = l.option_type and c.OptionListingSeq = l.Seq)" & vbCrLf
s = s & "LEFT OUTER JOIN attributelists l1 on l.ColorListid=l1.Listid" & vbCrLf
s = s & "LEFT OUTER JOIN attributelists l2 on l.styleListid=l2.Listid" & vbCrLf
s = s & "LEFT OUTER JOIN attributelists l3 on l.FinishListid=l3.Listid" & vbCrLf
s = s & "LEFT OUTER JOIN attributelists l4 on l.OtherListid=l4.Listid" & vbCrLf
s = s & "LEFT OUTER JOIN TBLAREAS T4 ON l.Location = T4.Area " & vbCrLf
s = s & "where isnull(c.AddToCO,0) = 0 and c.Customer_No =" & DbQuote(str, rs!Customer_no) & vbCrLf
s = s & "order by l.CatDesc,c.Description"
curPick.RecordSource = s
curPick.Refresh

'Set curTemp = Clone(curPick.Recordset)


Screen.MousePointer = vbDefault

If curPick.Recordset.RecordCount = 0 Then
   Command1.Enabled = False
   Command2.Enabled = False
   cmdSelect.Enabled = False
Else
   Command1.Enabled = True
   Command2.Enabled = True
   cmdSelect.Enabled = True
End If


Call LoadGrid
Exit Sub

herror:
If Err.Number <> 0 Then
   Screen.MousePointer = vbDefault
   Call errHandler(Err.Number & " " & Err.Description)
End If

End Sub

Private Sub Option5_Click()
On Error Resume Next
Dim s As String
Screen.MousePointer = vbHourglass
curPick.CursorType = adOpenDynamic
curPick.RecordSource = ""
s = "select c.Customer_No,c.Option_type,c.Category,c.Major_group,g.Description as MajDesc,l.description as CatDesc, c.OPT,c.Description, c.Qty, c.Rate as Price, c.Rate as BasePrice, c.Override_price as GrandTotal" & vbCrLf
s = s & ", c.Comments,c.ExportedtoBMT as Pick,'' as French_desc,'' as French_Comments,0 as NoCharge,c.EstimatorNotes,c.NoCommission,EnerGuideRating,GreenHouseRating,FuelCost" & vbCrLf
s = s & ",0 as usetax,0 as D_Structural_Change, c.Rate as CO_Price,c.Rate as BaseCO_Price,0 as ApplyIncentive,c.Cost_Amount,c.Assembly,c.Override_price as GrandCOTotal,0 as NetCOTax,c.TL_Extra,c.UOM,c.SalesWorksheet,c.Seq" & vbCrLf
s = s & ",c.Custom, c.Require_Quote,c.EstimateIndex,T4.Description Location,c.Color,c.ColorListID,c.StyleListID,c.FinishListID,c.OtherListID,c.Style,c.Finish,c.Other,l1.Required ColorRequired,l2.Required StyleRequired,l3.Required FinishRequired,l4.Required OtherRequired,l1.StrictList ColorStrict,l2.StrictList StyleStrict,l3.StrictList FinishStrict,l4.StrictList OtherStrict from tblscheduleb c" & vbCrLf
s = s & "LEFT OUTER JOIN attributelists l1 on c.ColorListid=l1.Listid" & vbCrLf
s = s & "LEFT OUTER JOIN attributelists l2 on c.styleListid=l2.Listid" & vbCrLf
s = s & "LEFT OUTER JOIN attributelists l3 on c.FinishListid=l3.Listid" & vbCrLf
s = s & "LEFT OUTER JOIN attributelists l4 on c.OtherListid=l4.Listid" & vbCrLf
s = s & "left outer join tblcategories l on (c.category = l.category)" & vbCrLf
s = s & "left outer join tblmajorgroups g on (l.Group_Code = g.Major_Group)" & vbCrLf
s = s & "LEFT OUTER JOIN TBLAREAS T4 ON c.Location = T4.Area " & vbCrLf
s = s & "where isnull(c.CONumber,'') = '' and c.Customer_No =" & DbQuote(str, rs!Customer_no) & vbCrLf
s = s & "order by l.Description,c.Description"
curPick.RecordSource = s
curPick.Refresh

'Set curTemp = Clone(curPick.Recordset)

Screen.MousePointer = vbDefault

If curPick.Recordset.RecordCount = 0 Then
   Command1.Enabled = False
   Command2.Enabled = False
   cmdSelect.Enabled = False
Else
   Command1.Enabled = True
   Command2.Enabled = True
   cmdSelect.Enabled = True
End If


Call LoadGrid
Exit Sub

herror:
If Err.Number <> 0 Then
   Screen.MousePointer = vbDefault
   Call errHandler(Err.Number & " " & Err.Description)
End If
End Sub

Public Sub ColorizeItems(Row As Long)
    Const RateZeroColor = 16646111 'light cyan &HC0C0FF 'pink
    Dim r As Long
    Dim startRow    As Long
    Dim EndRow      As Long
        
        
    With gItems
        .Redraw = flexRDNone
 
        For r = IIf(Row > 1, Row, 2) To IIf(Row > 1, Row, .Rows - 1)
            If .ValueMatrix(r, .ColIndex("ColorListID")) <> 0 And .TextMatrix(r, .ColIndex("Color")) = "" Then
                If .Cell(flexcpChecked, r, .ColIndex("ColorRequired")) = flexChecked Then
                    .Cell(flexcpForeColor, r, .ColIndex("Color")) = HFApp.Options(Format_InvalidData_ForeColor)
                    .Cell(flexcpBackColor, r, .ColIndex("Color")) = HFApp.Options(Format_InvalidData_BackColor)
                Else
                    .Cell(flexcpForeColor, r, .ColIndex("Color")) = HFApp.Options(Format_QtyRateEQZero_ForeColor)
                    .Cell(flexcpBackColor, r, .ColIndex("Color")) = HFApp.Options(Format_QtyRateEQZero_BackColor)
                End If
            Else
                .Cell(flexcpForeColor, r, .ColIndex("Color")) = vbWindowText
                .Cell(flexcpBackColor, r, .ColIndex("Color")) = vbWindowBackground
            End If
            If .ValueMatrix(r, .ColIndex("StyleListID")) <> 0 And .TextMatrix(r, .ColIndex("Style")) = "" Then
                If .Cell(flexcpChecked, r, .ColIndex("StyleRequired")) = flexChecked Then
                    .Cell(flexcpForeColor, r, .ColIndex("Style")) = HFApp.Options(Format_InvalidData_ForeColor)
                    .Cell(flexcpBackColor, r, .ColIndex("Style")) = HFApp.Options(Format_InvalidData_BackColor)
                Else
                    .Cell(flexcpForeColor, r, .ColIndex("Style")) = HFApp.Options(Format_QtyRateEQZero_ForeColor)
                    .Cell(flexcpBackColor, r, .ColIndex("Style")) = HFApp.Options(Format_QtyRateEQZero_BackColor)
                End If
            Else
                .Cell(flexcpForeColor, r, .ColIndex("Style")) = vbWindowText
                .Cell(flexcpBackColor, r, .ColIndex("Style")) = vbWindowBackground
            End If
            If .ValueMatrix(r, .ColIndex("FinishListID")) <> 0 And .TextMatrix(r, .ColIndex("Finish")) = "" Then
                If .Cell(flexcpChecked, r, .ColIndex("FinishRequired")) = flexChecked Then
                    .Cell(flexcpForeColor, r, .ColIndex("Finish")) = HFApp.Options(Format_InvalidData_ForeColor)
                    .Cell(flexcpBackColor, r, .ColIndex("Finish")) = HFApp.Options(Format_InvalidData_BackColor)
                Else
                    .Cell(flexcpForeColor, r, .ColIndex("Finish")) = HFApp.Options(Format_QtyRateEQZero_ForeColor)
                    .Cell(flexcpBackColor, r, .ColIndex("Finish")) = HFApp.Options(Format_QtyRateEQZero_BackColor)
                End If
            Else
                .Cell(flexcpForeColor, r, .ColIndex("Finish")) = vbWindowText
                .Cell(flexcpBackColor, r, .ColIndex("Finish")) = vbWindowBackground
            End If
            If .ValueMatrix(r, .ColIndex("OtherListID")) <> 0 And .TextMatrix(r, .ColIndex("Other")) = "" Then
                If .Cell(flexcpChecked, r, .ColIndex("OtherRequired")) = flexChecked Then
                    .Cell(flexcpForeColor, r, .ColIndex("Other")) = HFApp.Options(Format_InvalidData_ForeColor)
                    .Cell(flexcpBackColor, r, .ColIndex("Other")) = HFApp.Options(Format_InvalidData_BackColor)
                Else
                    .Cell(flexcpForeColor, r, .ColIndex("Other")) = HFApp.Options(Format_QtyRateEQZero_ForeColor)
                    .Cell(flexcpBackColor, r, .ColIndex("Other")) = HFApp.Options(Format_QtyRateEQZero_BackColor)
                End If
            Else
                .Cell(flexcpForeColor, r, .ColIndex("Other")) = vbWindowText
                .Cell(flexcpBackColor, r, .ColIndex("Other")) = vbWindowBackground
            End If
                
        Next
        
        
        
        .Redraw = flexRDBuffered
    End With


End Sub




Private Sub txtLocation_DropOpen(NoDefault As Boolean)
On Error Resume Next
Dim s As String
                s = "select a.Description,r.Model, r.Room from Rooms r" & vbCrLf
                s = s & "left outer join tblareas a on r.Room = a.Area" & vbCrLf
                s = s & " where Model =" & DbQuote(str, rs!Model) & vbCrLf
                s = s & "Union" & vbCrLf
                s = s & "select tblAreas.Description,DistinctModels.Model,tblAreas.Area from tblareas,DistinctModels" & vbCrLf
                s = s & " where Model =" & DbQuote(str, rs!Model)
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Description", s, "", True, , , "Model,Room") Then
                     txtLocation.Text = FPickList.SelectedItem("Description")
                End If
End Sub


Private Function Clone(ByVal oRs As ADODB.Recordset, _
    Optional ByVal LockType As ADODB.LockTypeEnum = adLockUnspecified) As ADODB.Recordset
    
    Dim oStream As ADODB.stream
    Dim oRsClone As ADODB.Recordset
    
    'save the recordset to the stream object
    Set oStream = New ADODB.stream
    oRs.Save oStream
    
    'and now open the stream object into a new recordset
    Set oRsClone = New ADODB.Recordset
    oRsClone.Open oStream, , , LockType
    
    'return the cloned recordset
    Set Clone = oRsClone
    
    'release the reference
    Set oRsClone = Nothing
End Function

Private Function NullCheck(ByVal VAmount, Optional vType As String)
On Error Resume Next
If UCase$(vType) = "S" Then
    NullCheck = "" & VAmount
Else
    NullCheck = IIf(IsNull(VAmount) Or VAmount = "", 0, VAmount)
End If
End Function
