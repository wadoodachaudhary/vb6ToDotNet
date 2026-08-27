VERSION 5.00
Begin VB.Form FChangeOrder 
   Caption         =   "Change Order"
   ClientHeight    =   4260
   ClientLeft      =   6225
   ClientTop       =   1395
   ClientWidth     =   11295
   Icon            =   "FChangeOrder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   11295
   Begin VB.TextBox txtApprovedDate 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   4890
      TabIndex        =   3
      Top             =   450
      Width           =   1740
   End
   Begin VB.TextBox txtApprovedBy 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   4890
      MaxLength       =   50
      TabIndex        =   2
      Top             =   210
      Width           =   2760
   End
   Begin VB.TextBox txtChangeDate 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1500
      TabIndex        =   1
      Top             =   450
      Width           =   1950
   End
   Begin VB.TextBox txtChangeNumber 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   630
   End
   Begin VB.TextBox txtComments 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3135
      Left            =   1485
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   930
      Width           =   9630
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   1
      Left            =   3465
      Picture         =   "FChangeOrder.frx":000C
      Top             =   450
      Width           =   240
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   0
      Left            =   6645
      Picture         =   "FChangeOrder.frx":0156
      Top             =   450
      Width           =   240
   End
   Begin VB.Label Label133 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Index           =   5
      Left            =   4470
      TabIndex        =   9
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label133 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Approved By"
      Height          =   195
      Index           =   4
      Left            =   3900
      TabIndex        =   8
      Top             =   210
      Width           =   915
   End
   Begin VB.Label Label133 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Change Date"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   450
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Change Number"
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   210
      Width           =   1155
   End
   Begin VB.Label Label111 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Comments"
      Height          =   195
      Index           =   10
      Left            =   675
      TabIndex        =   5
      Top             =   930
      Width           =   735
   End
End
Attribute VB_Name = "FChangeOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public EventTraps As Collection
Const SRCFILE = "FChangeOrder::"

Private mDirty    As Boolean
Private mCustomer    As String
Private mChangeOrder As String





Public Sub Edit(Customer As String, ChangeOrder As String)
    
    mCustomer = Customer
    mChangeOrder = ChangeOrder
    Me.Show vbModal
    
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadData
End Sub



Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
        
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "  from changeordermaster" & vbCrLf
    s = s & " where customer_no=" & DbQuote(Str, mCustomer)
    s = s & "   and change_order_no=" & DbQuote(Str, mChangeOrder)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        txtChangeNumber.Text = ""
        txtChangeDate.Text = format(Now(), "medium date")
        txtApprovedBy.Text = ""
        txtApprovedDate.Text = ""
        txtComments.Text = ""
    Else
        txtChangeNumber.Text = "" & rs("change_order_no")
        txtApprovedDate.Text = format("" & rs("change_date"), "medium date")
        txtApprovedBy.Text = "" & rs("bldr_approved_by")
        txtApprovedDate.Text = format("" & rs("bldr_approved_date"), "medium date")
        txtComments.Text = "" & rs("comments")
    End If
    
    Dirty = False
    
End Sub



Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    
'    Dim GLPrefix As String
'    Dim s As String
'    Dim Community As String
'    Dim CommunityPhase As String
'    Dim Model As String
'    Dim r As Long
'
'    If Not Dirty Then
'        SaveData = True
'        Exit Function
'    End If
'    If prompt Then
'        Select Case MsgBox(Me.Caption & " has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
'            Case vbNo
'                SaveData = True
'                Exit Function
'            Case vbCancel
'                SaveData = False
'                Exit Function
'        End Select
'    End If
'
'    txtJob.Text = Trim(txtJob.Text)
'    If Trim(txtJob.Text) = "" Then
'        MsgBox "A job number is required.", vbExclamation, App.ProductName
'        Call SetCtrlFocus(txtJob)
'        Exit Function
'    End If
'
'
'
'
'
'    Screen.MousePointer = vbHourglass
'
'
'
'    Community = GetComboBoxListKey(cboCommunity)
'    CommunityPhase = cboPhase.Text
'    GLPrefix = IIf(cboGLPrefix.Visible, GetComboBoxListKey(cboGLPrefix), txtGLPrefix.Text)
'    Model = GetComboBoxListKey(cboModel)
'
'    'save job
'    With gProperties
'        If mJob = Chr(1) Then
'            If Not ValidateJobNumber() Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Incorrect job format." & vbCrLf & "Correct format is " & Replace(HFApp.Options(Job_Mask), "&", "x") & vbCrLf & "Reenter job.", vbExclamation, App.ProductName
'                Call SetCtrlFocus(txtJob)
'                Exit Function
'            End If
'
'            s = ""
'            s = s & "INSERT INTO tblJobs(isquote,Job_No,ExternalJobID,Description,Notes,GL_Prefix,Community,CommunityPhase,municipal_address,city,province,zip,sitephone,sitefax,Inactive" & vbCrLf
'            s = s & "                   ,Lot,Block,LotPlan,Model" & vbCrLf
'            s = s & "                   ,LabourTaxGroup,MaterialTaxGroup,SubContractTaxGroup,EquipmentTaxGroup,OverheadTaxGroup,OtherTaxGroup)" & vbCrLf
'            s = s & "VALUES(0," & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
'            s = s & "      ," & DbQuote(Str, lblExternalJobID.Caption) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtNotes.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, GLPrefix) & vbCrLf
'            s = s & "      ," & DbQuote(Str, Community) & vbCrLf
'            s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtAddress.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtCity.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtProvince.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtPostal.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtPhone.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtFax.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Bit, cboStatus.ListIndex = 1) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtLot.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtBlock.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtLotPlan.Text) & vbCrLf
'
'            s = s & "      ," & DbQuote(Str, Model) & vbCrLf
'
'            s = s & "      ," & DbQuote(Str, .TextMatrix(1, .ColIndex("value"))) & vbCrLf
'            s = s & "      ," & DbQuote(Str, .TextMatrix(2, .ColIndex("value"))) & vbCrLf
'            s = s & "      ," & DbQuote(Str, .TextMatrix(3, .ColIndex("value"))) & vbCrLf
'            s = s & "      ," & DbQuote(Str, .TextMatrix(4, .ColIndex("value"))) & vbCrLf
'            s = s & "      ," & DbQuote(Str, .TextMatrix(5, .ColIndex("value"))) & vbCrLf
'            s = s & "      ," & DbQuote(Str, .TextMatrix(6, .ColIndex("value"))) & ")"
'            HFApp.SqlExec s
'            mJob = CleanJob(txtJob.Text)
'            Me.tag = mJob
'            txtJob.Enabled = False
'            cmdJob.Visible = False
'            txtJob.Width = 2865
'
'            Call HFApp.WriteJobToAccounting(CleanJob(txtJob.Text))
'
'            s = ""
'            s = s & "INSERT INTO tblcustomers (Customer_No,Job_no,Description,bal_sheet_Prefix,community,phase,address1,city,province,zip,phone,fax,purchased,approved,contract_assigned,cancelled,inactive,lot,block,lotplan,model,sale_posted)" & vbCrLf
'            s = s & "VALUES(" & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
'            s = s & "      ," & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, GLPrefix) & vbCrLf
'            s = s & "      ," & DbQuote(Str, Community) & vbCrLf
'            s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtAddress.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtCity.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtProvince.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtPostal.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtPhone.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtFax.Text) & ",1,1,1,0,0"
'            s = s & "      ," & DbQuote(Str, txtLot.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtBlock.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, txtLotPlan.Text) & vbCrLf
'            s = s & "      ," & DbQuote(Str, Model) & vbCrLf
'            s = s & "      ,1)"
'            HFApp.SqlExec s
'
'            s = ""
'            s = s & "INSERT INTO EstimateAssemblies(job,Customer_no,EstimateIndex,HFDescription,AssemblyType,OptionType,salesqty)" & vbCrLf
'            s = s & "VALUES(" & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
'            s = s & "      ," & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
'            s = s & "      ,0,'Manual Estimates',-1,-1,1)" & vbCrLf
'            HFApp.SqlExec s
'
'        Else
'            s = ""
'            s = s & "UPDATE tblJobs" & vbCrLf
'            s = s & "SET Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
'            s = s & "   ,Notes=" & DbQuote(Str, txtNotes.Text) & vbCrLf
'            s = s & "   ,GL_Prefix=" & DbQuote(Str, GLPrefix) & vbCrLf
'            s = s & "   ,Community=" & DbQuote(Str, Community) & vbCrLf
'            s = s & "   ,CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf
'            s = s & "   ,Municipal_Address=" & DbQuote(Str, txtAddress.Text) & vbCrLf
'            s = s & "   ,City=" & DbQuote(Str, txtCity.Text) & vbCrLf
'            s = s & "   ,Province=" & DbQuote(Str, txtProvince.Text) & vbCrLf
'            s = s & "   ,Zip=" & DbQuote(Str, txtPostal.Text) & vbCrLf
'            s = s & "   ,SitePhone=" & DbQuote(Str, txtPhone.Text) & vbCrLf
'            s = s & "   ,SiteFax=" & DbQuote(Str, txtFax.Text) & vbCrLf
'            s = s & "   ,Inactive=" & DbQuote(Bit, cboStatus.ListIndex = 1) & vbCrLf
'            s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
'            s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
'            s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
'            s = s & "   ,Model=" & DbQuote(Str, Model) & vbCrLf
'            s = s & "   ,LabourTaxGroup=" & DbQuote(Str, .TextMatrix(1, .ColIndex("value"))) & vbCrLf
'            s = s & "   ,MaterialTaxGroup=" & DbQuote(Str, .TextMatrix(2, .ColIndex("value"))) & vbCrLf
'            s = s & "   ,SubContractTaxGroup=" & DbQuote(Str, .TextMatrix(3, .ColIndex("value"))) & vbCrLf
'            s = s & "   ,EquipmentTaxGroup=" & DbQuote(Str, .TextMatrix(4, .ColIndex("value"))) & vbCrLf
'            s = s & "   ,OverheadTaxGroup=" & DbQuote(Str, .TextMatrix(5, .ColIndex("value"))) & vbCrLf
'            s = s & "   ,OtherTaxGroup=" & DbQuote(Str, .TextMatrix(6, .ColIndex("value"))) & vbCrLf
'            s = s & "WHERE Job_No=" & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
'            HFApp.SqlExec s
'
'
'            If HFApp.Options(SalesSystem) = SalesSystems.asNone Then
'                s = ""
'                s = s & "UPDATE tblcustomers" & vbCrLf
'                s = s & "SET Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
'                s = s & "   ,bal_sheet_Prefix=" & DbQuote(Str, GLPrefix) & vbCrLf
'                s = s & "   ,Community=" & DbQuote(Str, Community) & vbCrLf
'                s = s & "   ,Phase=" & DbQuote(Str, CommunityPhase) & vbCrLf
'                s = s & "   ,Address1=" & DbQuote(Str, txtAddress.Text) & vbCrLf
'                s = s & "   ,City=" & DbQuote(Str, txtCity.Text) & vbCrLf
'                s = s & "   ,Province=" & DbQuote(Str, txtProvince.Text) & vbCrLf
'                s = s & "   ,Zip=" & DbQuote(Str, txtPostal.Text) & vbCrLf
'                s = s & "   ,Phone=" & DbQuote(Str, txtPhone.Text) & vbCrLf
'                s = s & "   ,Fax=" & DbQuote(Str, txtFax.Text) & vbCrLf
'                s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
'                s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
'                s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
'                s = s & "   ,Model=" & DbQuote(Str, Model) & vbCrLf
'                s = s & "WHERE Customer_No=" & DbQuote(Str, txtJob.Text)
'                HFApp.SqlExec s
'            End If
'
'        End If
'
'
'    'save properties
'        s = ""
'        For r = 7 To .Rows - 1
'            If .TextMatrix(r, .ColIndex("DataType")) <> "" Then
'                s = s & "," & vbQuote & .TextMatrix(r, .ColIndex("Name")) & vbQuote & "=" & DbQuote(.ValueMatrix(r, .ColIndex("DataType")), .TextMatrix(r, .ColIndex("Value")))
'            End If
'        Next
'        If s <> "" Then
'            On Error Resume Next
'            Call HFApp.SqlExec("INSERT INTO JobCustomFields(Job_No) VALUES(" & DbQuote(Str, mJob) & ")", dbHomeFront)
'            On Error GoTo eh
'            Call HFApp.SqlExec("UPDATE JobCustomFields SET " & Mid(s, 2) & " WHERE Job_No=" & DbQuote(Str, mJob), dbHomeFront)
'        End If
'    End With
'
'
'    'save contacts
'    With gContacts
'        For r = 1 To .Rows - 1
'            s = "UPDATE tblJobs" & vbCrLf & _
'                "SET " & .TextMatrix(r, .ColIndex("ContactType")) & "=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PM"))) & vbCrLf & _
'                "WHERE Job_No=" & DbQuote(Str, mJob)
'            Call HFApp.SqlExec(s, dbHomeFront)
'        Next
'    End With
'
'
'    Call HFApp.WriteJobToAccounting(mJob)
'
'
    SaveData = True
    Dirty = False
'    Screen.MousePointer = vbDefault
'
'Exit Function
'eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Unable to save this job. " & vbQuote & txtJob.Text & vbQuote & " has already been used." & vbCrLf & vbCrLf & "Please enter a different job number.", vbExclamation, App.ProductName
'    Else
'        Call errHandler(SRCFILE & "SaveData", s)
'    End If
End Function







Private Sub Form_Resize()
On Error Resume Next
    txtComments.Move txtComments.Left, txtComments.Top, Me.ScaleWidth - txtComments.Left - 240, Me.ScaleHeight - txtComments.Height - 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub



Private Sub txtApprovedBy_Change()
    mDirty = True
End Sub
Private Sub txtApprovedDate_Change()
    mDirty = True
End Sub
Private Sub txtChangeDate_Change()
    mDirty = True
End Sub


Private Sub txtApprovedBy_GotFocus()
    SelectAll txtApprovedBy
End Sub
Private Sub txtApprovedDate_GotFocus()
    SelectAll txtApprovedDate
End Sub
Private Sub txtChangeDate_GotFocus()
    SelectAll txtChangeDate
End Sub

