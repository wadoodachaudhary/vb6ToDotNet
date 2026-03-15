VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FProjectManagers 
   Caption         =   "Project Managers"
   ClientHeight    =   5520
   ClientLeft      =   690
   ClientTop       =   2970
   ClientWidth     =   11505
   Icon            =   "FProjectManagers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   11505
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   10965
      _cx             =   19341
      _cy             =   3201
      Appearance      =   2
      BorderStyle     =   0
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FProjectManagers.frx":000C
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   1058
      ButtonWidth     =   1455
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    New    "
            Key             =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   4260
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   39
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":0220
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":0AFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":13D4
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":1CAE
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":2588
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":2E62
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":373C
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":4016
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":48F0
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":51CA
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":5AA4
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":637E
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":6C58
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":7532
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":7E0C
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":86E6
               Key             =   "New"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":8FC0
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":989A
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":A174
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":AA4E
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":B328
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":BC02
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":C4DC
               Key             =   "PricebookSearch"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":CDB6
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":D690
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":DF6A
               Key             =   "PricebookExport"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":E844
               Key             =   "PricebookImport"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":F11E
               Key             =   "PricebookNew"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":F9F8
               Key             =   "View"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":102D2
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":10BAC
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":11486
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":11D60
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":1207A
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":12954
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":1322E
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":13B08
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":143E2
               Key             =   "Models and Options"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FProjectManagers.frx":14CBC
               Key             =   "Generate"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FProjectManagers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FProjectManagers::"
Private mDirty As Boolean

Private mSMSCarriers As String


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyS And Shift = vbCtrlMask:       Call SaveData(False)
        Case KeyCode = vbKeyEscape:                         Unload Me
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo eh
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    Call LoadData
Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub


Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh:

    Dim rc As Long
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    rc = vbYes
    If prompt Then rc = MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, Me.Caption)
    Select Case rc
        Case vbNo:     SaveData = True:     Exit Function
        Case vbCancel: Exit Function
    End Select
    
    
    Screen.MousePointer = vbHourglass
    Dim r As Long
    Dim s As String
    Dim rs As Recordset
        
    With gData
        For r = 1 To .Rows - 1
            
            .Row = r
            Call .ShowCell(r, 0)
            .Refresh
            
            If .RowData(r) = "DIRTY" Then
            
                If Trim(.TextMatrix(r, .ColIndex("PM"))) = "" Then
                    Screen.MousePointer = vbDefault
                    Call MsgBox("The ID field cannot be blank", vbInformation, mProductName)
                    Exit Function
                End If
                
                If .Cell(flexcpData, r, .ColIndex("PM")) = "" Then
                    s = ""
                    s = s & "INSERT INTO tblProjectManager(PM,Password,PMName,Phone,Cell,Fax,Email,POLimit,PrimaryUserID,SecondaryUserID,PrjMgr,Estimator,Purchaser,InActive)" & vbCrLf
                    s = s & "VALUES(" & DbQuote(Str, .TextMatrix(r, .ColIndex("PM"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("Password"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("PMName"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Phone"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Cell"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Fax"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Email"))) & vbCrLf
                    s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("POLimit"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("PrimaryUserID"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("SecondaryUserID"))) & vbCrLf
                    s = s & "      ," & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("PrjMgr")) = flexChecked) & vbCrLf
                    s = s & "      ," & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("Estimator")) = flexChecked) & vbCrLf
                    s = s & "      ," & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("Purchaser")) = flexChecked) & vbCrLf
                    s = s & "      ," & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("Active")) = flexUnchecked) & ")"
                Else
                    s = ""
                    s = s & "UPDATE tblProjectManager" & vbCrLf
                    s = s & "SET PM=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PM"))) & vbCrLf
                    s = s & "   ,Password=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("Password"))) & vbCrLf
                    s = s & "   ,PMName=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PMName"))) & vbCrLf
                    s = s & "   ,Phone=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phone"))) & vbCrLf
                    s = s & "   ,Cell=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Cell"))) & vbCrLf
                    s = s & "   ,Fax=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Fax"))) & vbCrLf
                    s = s & "   ,Email=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Email"))) & vbCrLf
                    s = s & "   ,POLimit=" & DbQuote(Num, .TextMatrix(r, .ColIndex("POLimit"))) & vbCrLf
                    s = s & "   ,PrimaryUSerID=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PrimaryUserID"))) & vbCrLf
                    s = s & "   ,SecondaryUserID=" & DbQuote(Str, .TextMatrix(r, .ColIndex("SecondaryUserID"))) & vbCrLf
                    s = s & "   ,PrjMgr=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("PrjMgr")) = flexChecked) & vbCrLf
                    s = s & "   ,Estimator=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("Estimator")) = flexChecked) & vbCrLf
                    s = s & "   ,Purchaser=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("Purchaser")) = flexChecked) & vbCrLf
                    s = s & "   ,InActive=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("Active")) = flexUnchecked) & vbCrLf
                    s = s & "WHERE PM=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("PM"))) & vbCrLf
                End If
                Call HFApp.SqlExec(s)
                
                s = ""
                s = s & "update user_manager" & vbCrLf
                s = s & "set email=p.email" & vbCrLf
                s = s & "   ,firstname=" & DbQuote(Str, Parse(.TextMatrix(r, .ColIndex("PMName")), 1, " ")) & vbCrLf
                s = s & "   ,lastname=" & DbQuote(Str, Parse(.TextMatrix(r, .ColIndex("PMName")), 2, " ")) & vbCrLf
                s = s & "from tblprojectmanager p" & vbCrLf
                s = s & "join user_manager u on u.user_id=p.pm and u.pmuser=1" & vbCrLf
                s = s & "where p.pm=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("PM"))) & vbCrLf
                Call HFApp.SqlExec(s)
                
                
                'if id has changed then update jobs
                If .Cell(flexcpData, r, .ColIndex("PM")) <> .TextMatrix(r, .ColIndex("PM")) And .Cell(flexcpData, r, .ColIndex("PM")) <> "" Then
                    Call HFApp.SqlExec("UPDATE tblJobs SET PM=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PM"))) & " WHERE PM=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("PM"))), dbHomeFront)
                    Call HFApp.SqlExec("UPDATE tblJobs SET Estimator=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PM"))) & " WHERE Estimator=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("PM"))), dbHomeFront)
                    Call HFApp.SqlExec("UPDATE tblJobs SET Purchaser=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PM"))) & " WHERE Purchaser=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("PM"))), dbHomeFront)
                
                    'update the "original" id
                    .Cell(flexcpData, r, .ColIndex("PM")) = .TextMatrix(r, .ColIndex("PM"))
                End If
                
                'clear dirty flag
                .RowData(r) = ""
            End If
        Next
    End With
    
    mDirty = False
    SaveData = True
    Screen.MousePointer = vbDefault

Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Screen.MousePointer = vbDefault
        Call MsgBox("This ID is taken. Please choose something else.", vbExclamation, mProductName)
    Else
        Call errHandler(SRCFILE & "SaveData", s)
        Screen.MousePointer = vbDefault
    End If
End Function


Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim i As Long
    
        
    
    With gData
        .Redraw = flexRDNone
        .Rows = 1
        s = ""
        s = s & "SELECT *" & vbCrLf
        s = s & "  FROM tblProjectManager" & vbCrLf
'        s = s & " WHERE ISNULL(Inactive,0)=0" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            
            'save original id
            .Cell(flexcpData, r, .ColIndex("PM")) = "" & rs("PM")
            
            .TextMatrix(r, .ColIndex("PM")) = "" & rs("PM")
            
            .Cell(flexcpData, r, .ColIndex("Password")) = "" & rs("Password")
            .Cell(flexcpText, r, .ColIndex("Password")) = String(Len(.Cell(flexcpData, r, .ColIndex("Password"))), "*")
            
            .TextMatrix(r, .ColIndex("PMName")) = "" & rs("PMName")
            .TextMatrix(r, .ColIndex("Phone")) = "" & rs("Phone")
            .TextMatrix(r, .ColIndex("Cell")) = "" & rs("Cell")
            .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
            .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
            
            .TextMatrix(r, .ColIndex("POLimit")) = Val("" & rs("POLimit"))
            .TextMatrix(r, .ColIndex("PrimaryUserID")) = "" & rs("PrimaryUserID")
            .TextMatrix(r, .ColIndex("SecondaryUserID")) = "" & rs("SecondaryUserID")
            
            
            .Cell(flexcpChecked, r, .ColIndex("PrjMgr")) = "" & rs("PrjMgr") = "True"
            .Cell(flexcpChecked, r, .ColIndex("Estimator")) = "" & rs("Estimator") = "True"
            .Cell(flexcpChecked, r, .ColIndex("Purchaser")) = "" & rs("Purchaser") = "True"
            .Cell(flexcpChecked, r, .ColIndex("Active")) = "" & rs("InActive") <> "True"
            
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
    End With
    mDirty = False
    
    

End Sub


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
    .ComboList = ""
    Select Case .ColKey(Col)
        Case "PM":       .EditMaxLength = 10
        Case "Password": .EditMaxLength = 10
        Case "PMName":   .EditMaxLength = 50
        Case "Phone":    .EditMaxLength = 25
        Case "Cell":     .EditMaxLength = 25
        Case "Fax":      .EditMaxLength = 25
        Case "Email":    .EditMaxLength = 150
        Case "PrimaryUserID": .ComboList = "..."
        Case "SecondaryUserID": .ComboList = "..."
    End Select
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With gData
    Select Case .ColKey(Col)
        Case "PrimaryUserID", "SecondaryUserID":
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Purchasers", "select pm ID,pmname Name from tblprojectmanager where purchaser=1") Then
                .Text = FPickList.SelectedItem("ID")
                .RowData(Row) = "DIRTY"
                mDirty = True
            End If
    End Select
    End With
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mDirty = True
    
    With gData
    .RowData(Row) = "DIRTY"
    Select Case .ColKey(Col)
        Case "POLimit"
            .EditText = Val(.EditText)
            
        Case "Password":
            .Cell(flexcpData, Row, .ColIndex("Password")) = .EditText
            .Cell(flexcpText, Row, .ColIndex("Password")) = String(Len(.Cell(flexcpData, Row, .ColIndex("Password"))), "*")
            .EditText = .Cell(flexcpText, Row, .ColIndex("Password"))
    End Select
    End With
    
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim r As Long
    Select Case UCase(Trim(Button.Key))
        Case "NEW"
            With gData
                r = .Rows
                .AddItem ""
                .TextMatrix(r, .ColIndex("PMName")) = "(new manager)"
                .Cell(flexcpChecked, r, .ColIndex("PrjMgr")) = flexChecked
                .Cell(flexcpChecked, r, .ColIndex("Estimator")) = flexUnchecked
                .Cell(flexcpChecked, r, .ColIndex("Purchaser")) = flexUnchecked
                .Cell(flexcpChecked, r, .ColIndex("Active")) = flexChecked
                .RowData(r) = "DIRTY"
                mDirty = True
            End With
        Case "SAVE"
            Call SaveData(False)
    End Select
End Sub


