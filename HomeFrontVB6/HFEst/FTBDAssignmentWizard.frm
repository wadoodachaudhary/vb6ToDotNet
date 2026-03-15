VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FTBDAssignmentWizard 
   Caption         =   "PO Assignment Wizard"
   ClientHeight    =   8205
   ClientLeft      =   6750
   ClientTop       =   1275
   ClientWidth     =   11130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FTBDAssignmentWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11130
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   1588
      Caption         =   "Assign TBD Purchase Orders"
      Description     =   "Assign these POs to vendors or contractors."
      Icon            =   "FTBDAssignmentWizard.frx":000C
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11130
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7620
      Width           =   11130
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Save"
         Height          =   375
         Index           =   0
         Left            =   5085
         TabIndex        =   4
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   1
         Left            =   6180
         TabIndex        =   3
         Top             =   90
         Width           =   1095
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1860
      Left            =   105
      TabIndex        =   2
      Top             =   960
      Width           =   4845
      _cx             =   8546
      _cy             =   3281
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"FTBDAssignmentWizard.frx":08E6
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
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
      FrozenRows      =   1
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
End
Attribute VB_Name = "FTBDAssignmentWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FTBDAssignmentWizard::"

Dim mCancel As Boolean
Dim mDirty As Boolean

Public Function ShowForm() As Boolean
On Error GoTo eh:
    
    Dim s As String
    
    If Not HFApp.UserPermission("IssuePOs") Then
        MsgBox "You don't have permission to use this wizard. Contact an administrator to change your permissions.", vbInformation, App.ProductName
        Exit Function
    End If
    
    
    With gData
    
        s = ""
        s = s & "select Community,CommunityDesc,Job,JobDesc,POGroup,POIndex,PONumber,PODesc,PODate,null Vendor, null VendorDesc,Status" & vbCrLf
        s = s & "from purchaseorders" & vbCrLf
        s = s & "where Cancelled=0 and PostingBatch=0 and isTBDVendor=1" & vbCrLf
        s = s & "and divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        Call LoadSQL(s)
        
        Call IniGetGrid(Me, gData)
        Me.Show vbModal
    
        
    End With
    
    If mDirty And mCancel Then
        If MsgBox("Do you want to save these vendor assignments?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then mDirty = False
    End If
    
    If mDirty Then Call SaveData
    Unload Me
    
Exit Function
eh: Call errHandler(SRCFILE & "ShowForm")
End Function

Private Sub SaveData()
    Dim r As Long
    Dim AssignedCount As Long
    Dim ApprovedCount As Long
    Dim ApprovedPOs As String
    
    With gData
        
        'change the vendors
        ApprovedPOs = ""
        ApprovedCount = 0
        AssignedCount = 0
        For r = .Rows - 1 To 2 Step -1
        If Trim(.TextMatrix(r, .ColIndex("Vendor"))) <> "" Then
            AssignedCount = AssignedCount + 1
            Call ChangePOVendor(.TextMatrix(r, .ColIndex("PONumber")), .TextMatrix(r, .ColIndex("Vendor")), True)
            If .TextMatrix(r, .ColIndex("Status")) = "Approved" Then
                ApprovedCount = ApprovedCount + 1
                ApprovedPOs = ApprovedPOs & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("PONumber")))
            End If
            Call .RemoveItem(r)
        End If
        Next
        If AssignedCount = 0 Then Exit Sub
       
        
        'approve any that are pending?
        If ApprovedCount > 0 And HFApp.UserPermission("SendPO") Then
            If MsgBox(AssignedCount & " purchase orders have been assigned. " & ApprovedCount & " are approved. Do you want to send them now?", vbOKCancel + vbQuestion, App.ProductName) = vbOK Then
                Call SendingWizard("PO", , Mid(ApprovedPOs, 2))
            End If
        Else
            Call MsgBox(AssignedCount & " purchase orders have been assigned.", vbOK + vbInformation, App.ProductName)
        End If
    
    End With
    
    
End Sub

Private Sub LoadSQL(sql As String)
    Dim rs As Recordset
    Dim r As Long
    Dim c As Long
    
    With gData
        .Rows = 1 'clear filter
        .Rows = 2 'readd filterbar
        .FrozenRows = 1
        .Cols = 0
        
        Set rs = HFApp.SqlExec(sql)
        .Cols = rs.fields.Count
        For c = 0 To .Cols - 1
            .TextMatrix(0, c) = rs.fields(c).Name
            .ColKey(c) = rs.fields(c).Name
            If rs.fields(c).Type = adBoolean Then
                .ColDataType(c) = flexDTBoolean
            End If
        Next
        
        r = 1
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            For c = 0 To .Cols - 1
                .TextMatrix(r, c) = "" & rs(c)
            Next
            rs.MoveNext
        Wend
    
        Call .AutoSize(0, .Cols - 1)
    End With

End Sub

Private Sub ApplyFilter()
' apply filter to hide stuff that doesn't match
    Dim r As Long
    Dim c As Long

    Screen.MousePointer = vbHourglass
    With gData
        .Redraw = flexRDNone
        For r = 2 To .Rows - 1
            .RowHidden(r) = False
            For c = 0 To .Cols - 1
                If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                    .RowHidden(r) = True
                    Exit For
                End If
            Next
        Next
        .Redraw = flexRDBuffered
    
    End With
    Screen.MousePointer = vbDefault

End Sub
Private Sub cmdNav_Click(Index As Integer)
    mCancel = Index = 1
    Me.Hide
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60

    gData.Move margin, Me.WizHead1.Height + margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - WizHead1.Height - WizFoot.Height - 2 * margin
    
    cmdNav(0).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin

End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub



Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    If Row = 1 Then Call ApplyFilter
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With gData
    .ComboList = ""
    If Row = 1 Then
        .AutoSearch = flexSearchNone
    Else
        If IsIn(.ColKey(Col), "Vendor", "VendorDesc") Then
            .ComboList = "..."
        Else
            .AutoSearch = flexSearchFromCursor
            Cancel = True
        End If
    End If
    End With

End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If gData.MouseRow = 0 And Button = vbRightButton Then
        Cancel = True
        Call FMain.ShowColumnMenu(gData, , , , , False)
    End If
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim r As Long
    Dim s As String
        
    With gData
    
        s = "select vendor_id Vendor,vendor_name Name from tblvendors where isnull(istbd,0)=0 and inactive=0 and divisionid=" & DbQuote(Num, HFApp.DivisionID)
        If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s) Then
            For i = 0 To .SelectedRows - 1
                r = .SelectedRow(i)
                If Not .RowHidden(r) Then
                    .TextMatrix(r, .ColIndex("Vendor")) = FPickList.SelectedItem("Vendor")
                    .TextMatrix(r, .ColIndex("VendorDesc")) = FPickList.SelectedItem("Name")
                    mDirty = True
                End If
            Next
        End If
    
    End With
    
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim r As Long
        
    With gData
    
        If KeyCode <> vbKeyDelete Then Exit Sub
        If .Row < 1 Then Exit Sub
    
        If .Row = 1 Then
            .Text = ""
            Call ApplyFilter
        Else
            For i = 0 To .SelectedRows - 1
                r = .SelectedRow(i)
                If Not .RowHidden(r) Then
                    .TextMatrix(r, .ColIndex("Vendor")) = ""
                    .TextMatrix(r, .ColIndex("VendorDesc")) = ""
                End If
            Next
        End If
    End With
End Sub




