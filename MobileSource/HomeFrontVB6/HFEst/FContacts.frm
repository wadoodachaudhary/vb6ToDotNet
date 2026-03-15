VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FContacts 
   Caption         =   "Contacts"
   ClientHeight    =   5730
   ClientLeft      =   3330
   ClientTop       =   2295
   ClientWidth     =   8730
   Icon            =   "FContacts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8730
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   3
      Top             =   5100
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7530
      TabIndex        =   2
      Top             =   5130
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid gContacts 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   8505
      _cx             =   15002
      _cy             =   2514
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
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FContacts.frx":000C
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
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
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
   Begin VSFlex8Ctl.VSFlexGrid gJobContacts 
      Height          =   2625
      Left            =   120
      TabIndex        =   4
      Top             =   2310
      Width           =   11295
      _cx             =   19923
      _cy             =   4630
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FContacts.frx":00F6
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
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label lblContacts 
      AutoSize        =   -1  'True
      Caption         =   "Personnel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FContacts"

Private mJob As String
Private mIsQuote As Boolean 'this is a quote or came from a quote therefore you can edit customer info

Private Dirty As Boolean

Private Sub cmdNav_Click(Index As Integer)
    If SaveData(Index = 1) Then Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gContacts)
    Call IniGetGrid(Me, gJobContacts)

    With gJobContacts
        .ColComboList(.ColIndex("SMSAddress")) = "|" & .BuildComboList(HFApp.SqlExec("select address,name from smscarriers order by 2"), "*name,address", "address")
        .ColComboList(.ColIndex("SendVia")) = "#1;Print|#3;Email|#4;Fax"
    End With


End Sub

Private Sub Form_Resize()
    Const margin = 120
    
    gContacts.Width = Me.ScaleWidth - 2 * margin
    gJobContacts.Move gJobContacts.Left, gJobContacts.Top, gContacts.Width, Me.ScaleHeight - gJobContacts.Top - 2 * margin - cmdNav(0).Height
    
    cmdNav(0).Move Me.ScaleWidth - cmdNav(0).Width * 2 - margin * 2, Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(1).Move Me.ScaleWidth - cmdNav(0).Width * 1 - margin * 1, Me.ScaleHeight - cmdNav(0).Height - margin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gContacts)
    Call IniPutGrid(Me, gJobContacts)
End Sub

Public Sub ShowForm(Job As String)
    mJob = Job
    Call LoadData
    Me.Show vbModal
End Sub

Private Sub LoadData()
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    
    s = "select isQuote,Quote_No from tblJobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(str, mJob)
    Set rs = HFApp.SqlExec(s)
    mIsQuote = "" & rs("Quote_No") <> "" Or "" & rs("IsQuote") = "True"
    
    
    
    s = ""
    s = s & "SELECT 'vendor' icon,'Project Manager' ContactType,p.pm ID,p.pmname Name,p.phone,p.cell,p.fax,p.email" & vbCrLf
    s = s & "  FROM system_setup s" & vbCrLf
    s = s & "       left outer join tblJobs j on(s.id = j.DivisionID and Job_No=" & DbQuote(str, mJob) & ")" & vbCrLf
    s = s & "       left outer JOIN tblProjectManager p ON(j.PM=p.PM)" & vbCrLf
    s = s & " where s.id = " & HFApp.DivisionID & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'vendor' icon,'Estimator' ContactType,p.pm ID,p.pmname Name,p.phone,p.cell,p.fax,p.email" & vbCrLf
    s = s & "  FROM system_setup s" & vbCrLf
    s = s & "       left outer join tblJobs j on(s.ID = j.DivisionID and Job_No=" & DbQuote(str, mJob) & ")" & vbCrLf
    s = s & "       LEFT JOIN tblProjectManager p ON(j.Estimator=p.PM)" & vbCrLf
    s = s & " where s.id = " & HFApp.DivisionID & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'vendor' icon,'Purchaser' ContactType,p.pm ID,p.pmname Name,p.phone,p.cell,p.fax,p.email" & vbCrLf
    s = s & "  FROM system_setup s" & vbCrLf
    s = s & "       left outer join tblJobs j on(s.ID = j.DivisionID and Job_No=" & DbQuote(str, mJob) & ")" & vbCrLf
    s = s & "       LEFT JOIN tblProjectManager p ON(j.Purchaser=p.PM)" & vbCrLf
    s = s & " where s.id = " & HFApp.DivisionID & vbCrLf
'    s = s & "UNION ALL" & vbCrLf
'    s = s & "select 'customer' icon,'Buyer',customer_no,customer_name + isnull(' ' + customer_lname,''),phone,cellphone,fax,email" & vbCrLf
'    s = s & "  from tblcustomers" & vbCrLf
'    s = s & " where job_no=" & DbQuote(Str, mJob) & vbCrLf
'    s = s & "UNION ALL" & vbCrLf
'    s = s & "select 'customer' icon,'Cobuyer',customer_no,cobuyer_name + isnull(' ' + cobuyer_lname,''),c_phone,c_cellphone,c_fax,c_email" & vbCrLf
'    s = s & "  from tblcustomers" & vbCrLf
'    s = s & " where job_no=" & DbQuote(Str, mJob) & vbCrLf
    With gContacts
        r = 0
        .Rows = 1
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            .TextMatrix(r, .ColIndex("ContactType")) = "" & rs("ContactType")
            .Cell(flexcpPicture, r, .ColIndex("ContactType")) = FMain.SmallIcons.ListImages("" & rs("icon")).Picture
            .TextMatrix(r, .ColIndex("ID")) = "" & rs("ID")
            .TextMatrix(r, .ColIndex("Name")) = "" & rs("Name")
            .TextMatrix(r, .ColIndex("Phone")) = "" & rs("Phone")
            .TextMatrix(r, .ColIndex("Cell")) = "" & rs("Cell")
            .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
            .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
            rs.MoveNext
        Wend
    End With
    
    Call LoadJobContacts
    
    Dirty = False

End Sub


Private Sub gContacts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContacts
        .ComboList = ""
        Select Case .TextMatrix(Row, .ColIndex("ContactType"))
            Case "Buyer", "Cobuyer"
                If Not mIsQuote Then
                    Cancel = True
                Else
                    Select Case .ColKey(Col)
                        Case "ContactType": Cancel = True
                        Case "ID":          Cancel = True
                        Case "Phone":       .EditMaxLength = 30
                        Case "Cell":        .EditMaxLength = 30
                        Case "Fax":         .EditMaxLength = 30
                        Case "Email":       .EditMaxLength = 50
                    End Select
                End If
            Case Else
                Select Case .ColKey(Col)
                    Case "ID":     .ComboList = "..."
                    Case "Name":   .ComboList = "..."
                    Case Else:     Cancel = True
                End Select
        End Select
    End With
End Sub

Private Sub gContacts_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton Then
        If gContacts.MouseRow = 0 Then
            Cancel = True
            Call FMain.ShowColumnMenu(gContacts)
        End If
    End If
End Sub



Private Sub gJobContacts_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    With gJobContacts
        For r = 1 To .Rows - 1
            If .TextMatrix(r, .ColIndex("ContactID")) = .TextMatrix(Row, .ColIndex("ContactID")) And r <> Row Then
                .TextMatrix(r, .ColIndex("Company")) = .TextMatrix(Row, .ColIndex("Company"))
                .TextMatrix(r, .ColIndex("Name")) = .TextMatrix(Row, .ColIndex("Name"))
                .TextMatrix(r, .ColIndex("Phone")) = .TextMatrix(Row, .ColIndex("Phone"))
                .TextMatrix(r, .ColIndex("Cell")) = .TextMatrix(Row, .ColIndex("Cell"))
                .TextMatrix(r, .ColIndex("Fax")) = .TextMatrix(Row, .ColIndex("Fax"))
                .TextMatrix(r, .ColIndex("Email")) = .TextMatrix(Row, .ColIndex("Email"))
                .TextMatrix(r, .ColIndex("SendVia")) = .TextMatrix(Row, .ColIndex("SendVia"))
                .TextMatrix(r, .ColIndex("SmsAddress")) = .TextMatrix(Row, .ColIndex("SmsAddress"))
            End If
        Next
    End With
End Sub

Private Sub gJobContacts_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton Then
        If gJobContacts.MouseRow = 0 Then
            Cancel = True
            Call FMain.ShowColumnMenu(gJobContacts)
        End If
    End If
End Sub

Private Sub gContacts_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gContacts
    
        If .TextMatrix(Row, .ColIndex("ContactType")) = "Project Manager" Then
            s = "SELECT PM,PMName Name,Phone,fax,email,cell FROM tblProjectManager WHERE Inactive<>1 AND 1=PrjMgr"
        Else
            s = "SELECT PM,PMName Name,Phone,fax,email,cell FROM tblProjectManager WHERE Inactive<>1 AND 1=" & .TextMatrix(Row, .ColIndex("ContactType"))
        End If
        
        If FPickList.Choose(HFApp.Databases(dbHomeFront), Choose(Row, "Project Manager", "Estimator", "Purchaser"), s, .TextMatrix(Row, .ColIndex("ID")), , , , "PM,Phone,fax,email,cell") Then
            .TextMatrix(Row, .ColIndex("ID")) = FPickList.SelectedItem("PM")
            .TextMatrix(Row, .ColIndex("Name")) = FPickList.SelectedItem("Name")
            .TextMatrix(Row, .ColIndex("Phone")) = FPickList.SelectedItem("Phone")
            .TextMatrix(Row, .ColIndex("Cell")) = FPickList.SelectedItem("Mobil")
            .TextMatrix(Row, .ColIndex("Fax")) = FPickList.SelectedItem("Fax")
            .TextMatrix(Row, .ColIndex("Email")) = FPickList.SelectedItem("cell")
            Dirty = True
        End If
    End With
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    
    Dim s As String
    Dim r As Long
    
    If Not Dirty Then
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
    
    Screen.MousePointer = vbHourglass
    With gContacts
        For r = 1 To .Rows - 1
            Select Case .TextMatrix(r, .ColIndex("ContactType"))
                Case "Buyer"
                    s = ""
                    s = s & "update tblcustomers" & vbCrLf
                    s = s & "set customer_name=" & DbQuote(str, GetFirstName(.TextMatrix(r, .ColIndex("name")))) & vbCrLf
                    s = s & "   ,customer_lname=" & DbQuote(str, GetLastName(.TextMatrix(r, .ColIndex("name")))) & vbCrLf
                    s = s & "   ,phone=" & DbQuote(str, .TextMatrix(r, .ColIndex("phone"))) & vbCrLf
                    s = s & "   ,cellphone=" & DbQuote(str, .TextMatrix(r, .ColIndex("cell"))) & vbCrLf
                    s = s & "   ,fax=" & DbQuote(str, .TextMatrix(r, .ColIndex("fax"))) & vbCrLf
                    s = s & "   ,email=" & DbQuote(str, .TextMatrix(r, .ColIndex("email"))) & vbCrLf
                    s = s & "where customer_no=" & DbQuote(str, .TextMatrix(r, .ColIndex("ID"))) & vbCrLf
                Case "Cobuyer"
                    s = ""
                    s = s & "update tblcustomers" & vbCrLf
                    s = s & "set cobuyer_name=" & DbQuote(str, GetFirstName(.TextMatrix(r, .ColIndex("name")))) & vbCrLf
                    s = s & "   ,cobuyer_lname=" & DbQuote(str, GetLastName(.TextMatrix(r, .ColIndex("name")))) & vbCrLf
                    s = s & "   ,c_phone=" & DbQuote(str, .TextMatrix(r, .ColIndex("phone"))) & vbCrLf
                    s = s & "   ,c_cellphone=" & DbQuote(str, .TextMatrix(r, .ColIndex("cell"))) & vbCrLf
                    s = s & "   ,c_fax=" & DbQuote(str, .TextMatrix(r, .ColIndex("fax"))) & vbCrLf
                    s = s & "   ,c_email=" & DbQuote(str, .TextMatrix(r, .ColIndex("email"))) & vbCrLf
                    s = s & "where customer_no=" & DbQuote(str, .TextMatrix(r, .ColIndex("ID"))) & vbCrLf
                Case "Project Manager":    s = "UPDATE tblJobs SET PM=" & DbQuote(str, .TextMatrix(r, .ColIndex("ID"))) & "WHERE DivisionID = " & HFApp.DivisionID & " and Job_No=" & DbQuote(str, mJob)
                Case "Estimator":          s = "UPDATE tblJobs SET Estimator=" & DbQuote(str, .TextMatrix(r, .ColIndex("ID"))) & "WHERE DivisionID = " & HFApp.DivisionID & " and Job_No=" & DbQuote(str, mJob)
                Case "Purchaser":          s = "UPDATE tblJobs SET Purchaser=" & DbQuote(str, .TextMatrix(r, .ColIndex("ID"))) & "WHERE DivisionID = " & HFApp.DivisionID & " and Job_No=" & DbQuote(str, mJob)
                Case Else:                 s = ""
            End Select
            If s <> "" Then Call HFApp.SqlExec(s, dbHomeFront)
        Next
    End With
    
    
    Call SaveJobContacts

    SaveData = True
    Dirty = False
    Screen.MousePointer = vbDefault
    
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function


Private Function GetFirstName(FullName As String) As String
    GetFirstName = Parse(FullName, 1, " ")
End Function
Private Function GetLastName(FullName As String) As String
    Dim i As Long
    i = InStr(1, FullName, " ")
    If i = 0 Then
        GetLastName = ""
    Else
        GetLastName = Mid(FullName, i + 1)
    End If
End Function

Private Sub gContacts_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContacts
    Select Case .ColKey(Col)
        Case "Phone", "Cell", "Fax"
            .EditText = FormatPhone(.EditText)
    End Select
    End With
    Dirty = True
End Sub












Private Sub gJobContacts_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim arCust As String
    On Error Resume Next
    arCust = "" & HFApp.SqlExec("Select isnull(AR_Customer_Deposit,'') from tblcustomers where customer_no = " & DbQuote(str, mJob))(0)
    
    With gJobContacts
    Select Case .ColKey(Col)
        Case "Name"
            s = ""
            s = s & "Customer Contacts" & Chr(1) & vbCrLf
            s = s & "   select Role,firstname Name,ContactID" & vbCrLf
            s = s & "   from contacts " & vbCrLf
            s = s & "   where contacttypeid=99 and arcustomer=" & DbQuote(str, arCust) & Chr(0)
            s = s & "Vendor Contacts" & Chr(1) & vbCrLf
            s = s & "   select v.vendor_name Company,c.Role,c.firstname Name,ContactID" & vbCrLf
            s = s & "   from contacts c" & vbCrLf
            s = s & "   join tblvendors v on(c.vendorcode=v.vendor_id)" & vbCrLf
            s = s & "   where contacttypeid=99 and vendorcode<>''" & vbCrLf
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Contact", s, , , , , "ContactID") Then
                Set rs = HFApp.SqlExec("select * from contacts where contactid=" & DbQuote(Num, FPickList.SelectedItem("ContactID")))
                r = Row
                .TextMatrix(r, .ColIndex("ContactID")) = "" & rs("ContactID")
                .TextMatrix(r, .ColIndex("Company")) = "" & rs("CompanyName")
                .TextMatrix(r, .ColIndex("Name")) = "" & rs("FirstName")
                .TextMatrix(r, .ColIndex("Phone")) = "" & rs("WorkPhone")
                .TextMatrix(r, .ColIndex("Cell")) = "" & rs("CellPhone")
                .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
                .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
                .TextMatrix(r, .ColIndex("SendVia")) = Val("" & rs("CommModeID"))
                .TextMatrix(r, .ColIndex("SmsAddress")) = "" & rs("SmsAddress")
                Dirty = True
                .RowData(r) = "DIRTY"
            End If
            
    End Select
    End With
End Sub

Private Sub gJobContacts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r As Long
    If KeyCode = vbKeyDelete And Shift <> 0 And gJobContacts.Rows > 1 And gJobContacts.Row <> gJobContacts.Rows - 1 Then
    With gJobContacts
        r = .Row
        .TextMatrix(r, .ColIndex("ContactID")) = ""
        .TextMatrix(r, .ColIndex("Name")) = ""
        .TextMatrix(r, .ColIndex("Phone")) = ""
        .TextMatrix(r, .ColIndex("Cell")) = ""
        .TextMatrix(r, .ColIndex("Fax")) = ""
        .TextMatrix(r, .ColIndex("Email")) = ""
        .TextMatrix(r, .ColIndex("SendVia")) = ""
        .TextMatrix(r, .ColIndex("SmsAddress")) = ""
        .RowData(r) = "DIRTY"
        Dirty = True
    End With
    End If
End Sub

Private Sub gJobContacts_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim r As Long
    
    With gJobContacts
        Select Case .ColKey(Col)
            Case "Cell":    .EditText = FormatPhone(.EditText)
            Case "Phone":   .EditText = FormatPhone(.EditText)
            Case "fax":     .EditText = FormatPhone(.EditText)
        End Select
        .RowData(Row) = "DIRTY"
        Dirty = True
    End With
End Sub

Private Sub gJobContacts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gJobContacts
        .EditMaxLength = 0
        .ComboList = ""
        If .TextMatrix(Row, .ColIndex("ContactID")) = "" Then
            Cancel = True
        End If
        Select Case .ColKey(Col)
            Case "Role":         Cancel = True
            Case "Company":      Cancel = True
            Case "Phone":        .EditMaxLength = 35
            Case "Cell":         .EditMaxLength = 35
            Case "Fax":          .EditMaxLength = 35
            Case "Email":        .EditMaxLength = 500
            Case "Name":         .ComboList = "...": Cancel = False
            Case "Role":         .EditMaxLength = 25
            Case "SmsAddress":   .EditMaxLength = 50
        End Select
    End With
End Sub

Private Sub LoadJobContacts()
On Error GoTo eh
    Dim roles As String
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    
    roles = HFApp.Options.ValueByName("JobContactRoles")
    With gJobContacts
        
        'add roles
        .Rows = 1
        For r = 1 To Parse(roles, , "|")
            s = Parse(roles, r, "|")
            If s <> "" Then
                .AddItem ""
                .TextMatrix(.Rows - 1, .ColIndex("Role")) = s
            End If
        Next
    
        'now look up contacts assigned to those roles
        s = ""
        s = s & "select j.Role JobRole,c.*" & vbCrLf
        s = s & "  from jobContacts j" & vbCrLf
        s = s & "  join contacts c on(j.contactid=c.contactid)" & vbCrLf
        s = s & " where j.Contactid<>0 and j.job=" & DbQuote(str, mJob) & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = FindRow("" & rs("JobRole"))
            If r = -1 Then
                .AddItem ""
                r = .Rows - 1
            End If
            .TextMatrix(r, .ColIndex("ContactID")) = "" & rs("ContactID")
            .TextMatrix(r, .ColIndex("Role")) = "" & rs("JobRole")
            .TextMatrix(r, .ColIndex("Company")) = "" & rs("CompanyName")
            .TextMatrix(r, .ColIndex("Name")) = "" & rs("FirstName")
            .TextMatrix(r, .ColIndex("Phone")) = "" & rs("WorkPhone")
            .TextMatrix(r, .ColIndex("Cell")) = "" & rs("CellPhone")
            .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
            .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
            .TextMatrix(r, .ColIndex("SendVia")) = Val("" & rs("CommModeID"))
            .TextMatrix(r, .ColIndex("SmsAddress")) = "" & rs("SmsAddress")
            rs.MoveNext
        Wend
    End With

Exit Sub
eh: errHandler ("LoadJobContacts")
End Sub

Private Function FindRow(Role As String) As Long
    Dim r As Long
    With gJobContacts
    FindRow = -1
    For r = 1 To .Rows - 1
        If .TextMatrix(r, .ColIndex("Role")) = Role Then
            FindRow = r
            Exit Function
        End If
    Next
    End With
End Function

Private Sub SaveJobContacts()
On Error GoTo eh
    Dim s As String
    Dim r As Long
    
    With gJobContacts
        For r = 1 To .Rows - 1
            If .RowData(r) = "DIRTY" Then
                
                
                s = ""
                s = s & "INSERT INTO JobContacts(ContactID,Role,Job)" & vbCrLf
                s = s & "VALUES(" & DbQuote(Num, .TextMatrix(r, .ColIndex("ContactID"))) & vbCrLf
                s = s & "      ," & DbQuote(str, .TextMatrix(r, .ColIndex("Role"))) & vbCrLf
                s = s & "      ," & DbQuote(str, mJob) & ")"
                On Error Resume Next
                Call HFApp.SqlExec(s)
                On Error GoTo eh
                
                s = ""
                s = s & "UPDATE JobContacts" & vbCrLf
                s = s & "SET ContactID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ContactID"))) & vbCrLf
                s = s & "WHERE Role=" & DbQuote(str, .TextMatrix(r, .ColIndex("Role"))) & vbCrLf
                s = s & "  AND Job=" & DbQuote(str, mJob) & vbCrLf
                Call HFApp.SqlExec(s)
            
            
                If .TextMatrix(r, .ColIndex("ContactID")) <> "" Then
                    s = ""
                    s = s & "UPDATE Contacts" & vbCrLf
                    s = s & "SET Role=" & DbQuote(str, .TextMatrix(r, .ColIndex("Role"))) & vbCrLf
                    s = s & "   ,FirstName=" & DbQuote(str, .TextMatrix(r, .ColIndex("Name"))) & vbCrLf
                    s = s & "   ,WorkPhone=" & DbQuote(str, .TextMatrix(r, .ColIndex("Phone"))) & vbCrLf
                    s = s & "   ,CellPhone=" & DbQuote(str, .TextMatrix(r, .ColIndex("Cell"))) & vbCrLf
                    s = s & "   ,Fax=" & DbQuote(str, .TextMatrix(r, .ColIndex("Fax"))) & vbCrLf
                    s = s & "   ,Email=" & DbQuote(str, .TextMatrix(r, .ColIndex("Email"))) & vbCrLf
                    s = s & "   ,CommModeID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("SendVia"))) & vbCrLf
                    s = s & "   ,SmsAddress=" & DbQuote(str, .TextMatrix(r, .ColIndex("SmsAddress"))) & vbCrLf
                    s = s & "WHERE ContactID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ContactID"))) & vbCrLf
                    Call HFApp.SqlExec(s)
                End If
                .RowData(r) = ""
            End If
        Next
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "SaveJobContacts", s)
End Sub


