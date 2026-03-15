VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FPOPriceChangeWiz 
   Caption         =   "Purchase Order Price Update Wizard"
   ClientHeight    =   11760
   ClientLeft      =   3420
   ClientTop       =   1875
   ClientWidth     =   18045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FPOPriceChangeWiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11760
   ScaleWidth      =   18045
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   2
      Left            =   8385
      TabIndex        =   25
      Top             =   4140
      Visible         =   0   'False
      Width           =   7395
      Begin VSFlex8Ctl.VSFlexGrid gItems 
         Height          =   3285
         Left            =   1020
         TabIndex        =   26
         Top             =   1200
         Width           =   6165
         _cx             =   10874
         _cy             =   5794
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FPOPriceChangeWiz.frx":000C
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
      Begin HFEst.VBCombo cboCategory 
         Height          =   240
         Left            =   3075
         TabIndex        =   29
         Top             =   885
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.Label Label144 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Variance Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1380
         TabIndex        =   30
         Top             =   915
         Width           =   1575
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verify the new item rates"
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
         Index           =   5
         Left            =   900
         TabIndex        =   28
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   2115
      End
      Begin VB.Label lblDescription 
         Caption         =   $"FPOPriceChangeWiz.frx":00E1
         Height          =   465
         Index           =   4
         Left            =   1050
         TabIndex        =   27
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5805
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   3
         Left            =   255
         Picture         =   "FPOPriceChangeWiz.frx":0174
         Top             =   255
         Width           =   480
      End
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   18045
      _ExtentX        =   31829
      _ExtentY        =   1588
      Caption         =   "Update Purchase Orders Prices"
      Description     =   "The PO Price Update wizard will help you change the vendor costs on purchase orders"
      Icon            =   "FPOPriceChangeWiz.frx":0A3E
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   7395
      Begin VSFlex8Ctl.VSFlexGrid gPOs 
         Height          =   3585
         Left            =   1020
         TabIndex        =   21
         Top             =   900
         Width           =   6165
         _cx             =   10874
         _cy             =   6324
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FPOPriceChangeWiz.frx":1318
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
      Begin VB.Label lblDescription 
         Caption         =   "Review the list below to verifying the purchase orders you have selected."
         Height          =   465
         Index           =   3
         Left            =   1050
         TabIndex        =   23
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5805
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verify which POs will be updated"
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
         Index           =   2
         Left            =   900
         TabIndex        =   22
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   2805
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   1
         Left            =   255
         Picture         =   "FPOPriceChangeWiz.frx":13ED
         Top             =   255
         Width           =   480
      End
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   18045
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   11175
      Width           =   18045
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Commi&t"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   -60
         X2              =   26420
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   870
      Visible         =   0   'False
      Width           =   7395
      Begin VB.TextBox txtPONumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   5
         Top             =   3030
         Width           =   3915
      End
      Begin VB.TextBox txtPOIndex 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   4
         Top             =   2790
         Width           =   3915
      End
      Begin VB.TextBox txtPOGroup 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   3
         Top             =   2550
         Width           =   3915
      End
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   2
         Top             =   2310
         Width           =   3915
      End
      Begin VB.TextBox txtJob 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   1
         Top             =   2070
         Width           =   3915
      End
      Begin VB.TextBox txtCommunity 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   0
         Top             =   1830
         Width           =   3915
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FPOPriceChangeWiz.frx":1CB7
         Height          =   1125
         Index           =   1
         Left            =   1050
         TabIndex        =   20
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5865
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PO Number"
         Height          =   195
         Index           =   5
         Left            =   1470
         TabIndex        =   19
         Top             =   3060
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PO Index"
         Height          =   195
         Index           =   4
         Left            =   1635
         TabIndex        =   18
         Top             =   2820
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PO Group"
         Height          =   195
         Index           =   3
         Left            =   1590
         TabIndex        =   17
         Top             =   2580
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendor"
         Height          =   195
         Index           =   2
         Left            =   1785
         TabIndex        =   16
         Top             =   2340
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Job"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   15
         Top             =   2100
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Community"
         Height          =   195
         Index           =   0
         Left            =   1530
         TabIndex        =   14
         Top             =   1860
         Width           =   765
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Purchase Orders"
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
         Index           =   0
         Left            =   900
         TabIndex        =   13
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   1905
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FPOPriceChangeWiz.frx":1E84
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Remove"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FPOPriceChangeWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FPOPriceChangeWiz::"

Private mRecordset As Recordset
Private mPONumber As String

Public Function ShowForm(Optional PONumber As String)
    cmdNav(2).Enabled = True
    mPONumber = PONumber
    Me.Show vbModal
End Function


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
On Error GoTo eh
    Dim i As Long
    Dim c As Long
    Dim r As Long
    Dim s As String
    Dim rs As Recordset
    
    Screen.MousePointer = vbHourglass
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
        
        
    Select Case RHS
        Case 0 'filter pos
            Set gPOs.DataSource = Nothing
            gPOs.Rows = 1
            Set gItems.DataSource = Nothing
            gItems.Rows = 1
            
        Case 1 'POs
            With gPOs
                .Rows = 1
                Set rs = HFApp.SqlExec(QueryPOs, dbHomefront)
                
                'load headings
                .Cols = rs.fields.Count
                .ColDataType(0) = flexDTBoolean
                For c = 0 To gPOs.Cols - 1
                    .TextMatrix(0, c) = rs.fields(c).Name
                    .ColKey(c) = rs.fields(c).Name
                    .ColDataType(c) = flexDTString
                Next
                .TextMatrix(0, .ColIndex("Community")) = FMain.CD_Community
                .TextMatrix(0, .ColIndex("CommunityDesc")) = FMain.CD_Community & " Desc"
                .ColDataType(.ColIndex("PODate")) = flexDTDate
                .ColFormat(.ColIndex("PODate")) = HFApp.Options(DateFormat)
                
                'load data
                r = 0
                While Not rs.EOF
                    .AddItem ""
                    r = r + 1
                    For c = 0 To .Cols - 1
                        .TextMatrix(r, c) = "" & rs(c)
                    Next
                    rs.MoveNext
                Wend
                
                'size columns
                Call .AutoSize(0, .Cols - 1)
            End With
            
        Case 2 'items
            With gItems
                .Rows = 1
                Set rs = HFApp.SqlExec(QueryItems, dbHomefront)
                
                'load headings
                .Cols = rs.fields.Count
                For c = 0 To .Cols - 1
                    .TextMatrix(0, c) = rs.fields(c).Name
                    .ColKey(c) = rs.fields(c).Name
                    .ColDataType(c) = flexDTString
                Next
                
                'load data
                r = 0
                While Not rs.EOF
                    .AddItem ""
                    r = r + 1
                    For c = 0 To .Cols - 1
                        .TextMatrix(r, c) = "" & rs(c)
                    Next
                    
                    If .TextMatrix(r, .ColIndex("NewRate")) = "" Then .TextMatrix(r, .ColIndex("NewRate")) = .TextMatrix(r, .ColIndex("OrigRate"))
                    
                    If .TextMatrix(r, .ColIndex("NewRate")) <> .TextMatrix(r, .ColIndex("OrigRate")) Then .Cell(flexcpForeColor, r, .ColIndex("NewRate")) = vbRed
                    
                    rs.MoveNext
                Wend
                
                'size columns
                Call .AutoSize(0, .Cols - 1)
            End With
            
    End Select
    
    Screen.MousePointer = vbDefault
    
Exit Property
eh: Call errHandler(SRCFILE & "CurrentFrame")
End Property





Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3:
            If ValidateData() Then
                Call SaveData
                Unload Me
            End If
    End Select
End Sub

Private Sub Form_Activate()
    
    If mPONumber = "" Then Exit Sub
    
    txtPONumber.Text = mPONumber
    mPONumber = ""
    CurrentFrame = 2
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).BorderStyle = 0
        WizFrame(i).Move 0, WizHead1.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead1.Height - WizFoot.Height
    Next
    
    gPOs.Move gPOs.Left, gPOs.Top, WizFrame(0).Width - gPOs.Left - 180, WizFrame(0).Height - gPOs.Top - 180
    gItems.Move gItems.Left, gItems.Top, WizFrame(0).Width - gItems.Left - 180, WizFrame(0).Height - gItems.Top - 180
    
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadCustomDescriptions
    Call LoadCategories
    
    CurrentFrame = 0
End Sub


Private Sub LoadCategories()
    Dim s As String
    s = ""
    s = s & "SELECT isnull(description,category),category,0,isvariance FROM standardcategories " & vbCrLf
    s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
    s = s & "order by 4 desc,1" & vbCrLf
    Call LoadComboBox(cboCategory, HFApp.Databases(dbHomefront), s)


    s = HFApp.Options.ValueByName("ChangeWizCategory")
    If s <> "" Then
        Call SetListIndex(cboCategory, , s)
    Else
        cboCategory.ListIndex = 0
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    HFApp.Options.ValueByName("ChangeWizCategory") = cboCategory.Text
    Call IniPutForm(Me)
End Sub

Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim r As Long
    With gItems
    For r = 1 To .Rows - 1
        If .TextMatrix(r, .ColIndex("NewRate")) <> .TextMatrix(r, .ColIndex("OrigRate")) Then
            .Cell(flexcpForeColor, r, .ColIndex("NewRate")) = vbRed
        Else
            .Cell(flexcpForeColor, r, .ColIndex("NewRate")) = vbWindowText
        End If
    Next
    End With
End Sub

Private Sub gItems_RowColChange()
    With gItems
    
        If .ColKey(.Col) = "NewRate" Then
            .Editable = flexEDKbdMouse
            .AutoSearch = flexSearchNone
        Else
            .Editable = flexEDNone
            .AutoSearch = flexSearchFromCursor
        End If
    
    End With
End Sub

Private Sub gPOs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = gPOs.ColKey(Col) <> "Selected"
End Sub

Private Function WhereClause(Optional TableAlias As String) As String
    Dim s As String
    
    If TableAlias <> "" Then TableAlias = TableAlias & "."
    
    If txtCommunity.Text <> "" Then s = s & "       and (" & TableAlias & "community like " & DbQuote(Str, txtCommunity.Text) & " or " & TableAlias & "communitydesc like " & DbQuote(Str, txtCommunity.Text) & ")" & vbCrLf
    If txtPOIndex.Text <> "" Then s = s & "       and (" & TableAlias & "poindex like " & DbQuote(Str, txtPOIndex.Text) & " or " & TableAlias & "poindexdescription like " & DbQuote(Str, txtPOIndex.Text) & ")" & vbCrLf
    If txtPOGroup.Text <> "" Then s = s & "       and isnull(" & TableAlias & "pogroup,'') like " & DbQuote(Str, txtPOGroup.Text) & vbCrLf
    If txtPONumber.Text <> "" Then s = s & "       and " & TableAlias & "ponumber like " & DbQuote(Str, txtPONumber.Text) & vbCrLf
    
    If TableAlias = "v." Then
        If txtJob.Text <> "" Then s = s & "       and (" & TableAlias & "job_no like " & DbQuote(Str, txtJob.Text) & " or " & TableAlias & "jobdesc like " & DbQuote(Str, txtJob.Text) & ")" & vbCrLf
        If txtVendor.Text <> "" Then s = s & "       and (" & TableAlias & "povendor like " & DbQuote(Str, txtVendor.Text) & " or " & TableAlias & "povendorname like " & DbQuote(Str, txtVendor.Text) & ")" & vbCrLf
    Else
        If txtJob.Text <> "" Then s = s & "       and (" & TableAlias & "job like " & DbQuote(Str, txtJob.Text) & " or " & TableAlias & "jobdesc like " & DbQuote(Str, txtJob.Text) & ")" & vbCrLf
        If txtVendor.Text <> "" Then s = s & "       and (" & TableAlias & "vendor like " & DbQuote(Str, txtVendor.Text) & " or " & TableAlias & "vendordesc like " & DbQuote(Str, txtVendor.Text) & ")" & vbCrLf
    End If
    
    WhereClause = Trim(s)
End Function


Private Function ValidateData() As Boolean
    Dim r As Long
    Dim c As Long
    
    If GetComboBoxListKey(cboCategory) = "" Then
        MsgBox "Variance category is required", vbExclamation, App.ProductName
        Exit Function
    End If
    
    
    With gItems
    For r = 1 To .Rows - 1
        If .TextMatrix(r, .ColIndex("NewRate")) <> .TextMatrix(r, .ColIndex("OrigRate")) Then c = c + 1
    Next
    End With
    If c < 1 Then
        MsgBox "No prices have changed", vbInformation, App.ProductName
        Exit Function
    End If
    
    
    ValidateData = True
    
End Function


Private Sub SaveData()
On Error GoTo eh
    
    Dim r As Long
    Dim s As String
    Dim SessionKey As String
    Dim PriceDelta As Double
    Dim POs As String
    Dim rs As Recordset
    
    SessionKey = HFApp.LoginID & format(Now(), "--yyyymmddhhmmss")


    With gItems
        For r = 1 To .Rows - 1
        
            PriceDelta = .ValueMatrix(r, .ColIndex("NewRate")) - .ValueMatrix(r, .ColIndex("OrigRate"))
            If PriceDelta <> 0 Then
        
                'mark pos with the changeorderaudit tag so pos cant be repriced a second time
                s = ""
                s = s & "update pm set" & vbCrLf
                s = s & "datesenttobuildpro=null" & vbCrLf
                s = s & ",ChangeOrderAudit=" & DbQuote(Str, SessionKey) & vbCrLf
                s = s & "from EstimateItems i " & vbCrLf
                s = s & "join EstimateAssemblies a ON(i.EstAssemblyID=a.EstAssemblyID)" & vbCrLf
                s = s & "join EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                s = s & "join NextPOChangeOrders n on(i.divisionid=n.divisionid and i.ponumber=n.ponumber)" & vbCrLf
                s = s & "join PurchaseOrders p on i.divisionid=p.divisionid and i.ponumber=p.ponumber" & vbCrLf
                s = s & "join pomaster pm on i.divisionid=pm.divisionid and i.ponumber=pm.ponumber" & vbCrLf
                s = s & "left join invoiceitems ii on p.divisionid=ii.divisionid and p.ponumber=ii.commitment" & vbCrLf
                s = s & "" & vbCrLf
                s = s & "where ii.invoice is null" & vbCrLf
                s = s & "and p.Cancelled=0" & vbCrLf
                s = s & "and i.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & WhereClause("v")
                s = s & "and v.povendor=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor"))) & vbCrLf
                s = s & "and i.phase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase"))) & vbCrLf
                s = s & "and i.item=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Item"))) & vbCrLf
                s = s & "and i.description=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                s = s & "and i.porate=" & DbQuote(Num, .TextMatrix(r, .ColIndex("OrigRate"))) & vbCrLf
                s = s & "and isnull(i.RowType,'') not in('Original - Cancelled','Original - Reversal')" & vbCrLf
                HFApp.SqlExec s
                
                
                'create new items for price changes
                s = ""
                s = s & "insert estimateitems(" & vbCrLf
                s = s & " ChangeOrderAudit" & vbCrLf
                s = s & ",ChangeOrderOrigEstItemID" & vbCrLf
                s = s & ",POChangeOrderNumber" & vbCrLf
                s = s & ",VarianceJCCategory" & vbCrLf
                s = s & ",CreatedDate, CreatedBy, ModifiedDate, ModifiedBy" & vbCrLf
                s = s & "" & vbCrLf
                s = s & ",BudgetVendor, BudgetQty, BudgetTaxGroup, BudgetJCTaxRate, BudgetNJCTaxRate" & vbCrLf
                s = s & "" & vbCrLf
                s = s & ",BudgetGenerated" & vbCrLf
                s = s & ",BudgetPostingBatch" & vbCrLf
                s = s & ",BudgetDeleted" & vbCrLf
                s = s & ",BudgetOverridden" & vbCrLf
                s = s & ",BudgetPostingDate" & vbCrLf
                s = s & ",BudgetRate" & vbCrLf
                s = s & ",BudgetPretax" & vbCrLf
                s = s & ",BudgetJCTax" & vbCrLf
                s = s & ",BudgetNJCTax" & vbCrLf
                s = s & "" & vbCrLf
                s = s & ",POVendor, POQty, POTaxGroup, POJCTaxRate, PONJCTaxRate, PONumber, ExcludeFromPO, PODeleted, POOverridden" & vbCrLf
                s = s & ",PORate" & vbCrLf
                s = s & ",POPretax" & vbCrLf
                s = s & ",POJCTax" & vbCrLf
                s = s & ",PONJCTax" & vbCrLf
                s = s & "" & vbCrLf
                s = s & ",POGenBatch, EstAssemblyID, POIndex, Phase, Item, Job, JCExtra, JCCostCode"
                s = s & ",JCCategory, OriginalJCCategory"
                s = s & ",SortOrder, Description, Comments, TakeoffQty, TakeoffUOM, ConversionFactor, OrderUOM" & vbCrLf
                s = s & ",Assembly, AssemblyDescription, Model, Location, RFP, Formula, Unit, BillingItemID, SalesQty, DivisionID" & vbCrLf
                s = s & ",RequestDetailID, Sequence " & vbCrLf
                s = s & ",WBS01, WBS02, WBS03, WBS04, WBS05, WBS06, WBS07, WBS08, WBS09, WBS10 " & vbCrLf
                s = s & ",WBS11, WBS12, WBS13, WBS14, WBS15, WBS16, WBS17, WBS18, WBS19, WBS20 " & vbCrLf
                s = s & ",WBS21, WBS22, WBS23, WBS24, WBS25, WBS26, WBS27, WBS28, WBS29, WBS30 " & vbCrLf
                s = s & ",WBS31, WBS32, WBS33, WBS34, WBS35, WBS36, WBS37, WBS38, WBS39, WBS40)" & vbCrLf
                s = s & "" & vbCrLf '-------------------------------------------------------------------------------------------------
                s = s & "select " & vbCrLf
                s = s & DbQuote(Str, SessionKey) & vbCrLf
                s = s & ",i.EstItemID ChangeOrderOrigEstItemID" & vbCrLf
                s = s & ",n.NextChangeNumber ChangeOrder" & vbCrLf
                
                If PostPOQtyToAccounting Then
                    s = s & ",null VarianceJCCategory" & vbCrLf
                Else
                    s = s & "," & DbQuote(Str, GetComboBoxListKey(cboCategory)) & " VarianceJCCategory" & vbCrLf
                End If
            
                s = s & ",getdate() CreatedDate, " & DbQuote(Str, HFApp.LoginID) & " CreatedBy, getdate() ModifiedDate, " & DbQuote(Str, HFApp.LoginID) & " ModifiedBy" & vbCrLf
                s = s & "" & vbCrLf
                s = s & ",i.BudgetVendor, i.BudgetQty, i.BudgetTaxGroup, i.BudgetJCTaxRate, i.BudgetNJCTaxRate" & vbCrLf
                s = s & "" & vbCrLf
                s = s & ",0 BudgetGenerated" & vbCrLf
                s = s & ",0 BudgetPostingBatch" & vbCrLf
                s = s & ",a.Budgetslocked BudgetDeleted" & vbCrLf
                s = s & ",0 BudgetOverridden" & vbCrLf
                s = s & ",null BudgetPostingDate " & vbCrLf
                s = s & ",case when a.Budgetslocked=1 then 0 else " & DbQuote(Num, PriceDelta) & " end BudgetRate" & vbCrLf
                s = s & ",round(case when a.Budgetslocked=1 then 0 else " & DbQuote(Num, PriceDelta) & " end * i.budgetqty,2) BudgetPretax " & vbCrLf
                
                s = s & ",round(round(case when a.Budgetslocked=1 then 0 else " & DbQuote(Num, PriceDelta) & " end * i.budgetqty,2) * i.BudgetJCTaxRate/100,2)  BudgetJCTax" & vbCrLf
                s = s & ",round(round(case when a.Budgetslocked=1 then 0 else " & DbQuote(Num, PriceDelta) & " end * i.budgetqty,2) * i.BudgetNJCTaxRate/100,2)  BudgetNJCTax" & vbCrLf
                s = s & "" & vbCrLf
                s = s & ",i.POVendor, i.POQty, i.POTaxGroup, i.POJCTaxRate, i.PONJCTaxRate, i.PONumber, i.ExcludeFromPO, i.PODeleted, i.POOverridden" & vbCrLf
                
                s = s & "," & DbQuote(Num, PriceDelta) & " PORate" & vbCrLf
                s = s & ",round(i.poqty* " & DbQuote(Num, PriceDelta) & ",2) POPretax " & vbCrLf
                s = s & ",round(round(i.poqty* " & DbQuote(Num, PriceDelta) & ",2)*i.POJCTaxRate/100,2)  POJCTax " & vbCrLf
                s = s & ",round(round(i.poqty* " & DbQuote(Num, PriceDelta) & ",2)*i.PONJCTaxRate/100,2)  PONJCTax " & vbCrLf
                
                
                s = s & "" & vbCrLf
                s = s & ",i.POGenBatch, i.EstAssemblyID, i.POIndex, i.Phase, i.Item, i.Job, i.JCExtra, i.JCCostCode"
                s = s & "," & DbQuote(Str, GetComboBoxListKey(cboCategory)) & " JCCategory" & vbCrLf
                s = s & "," & DbQuote(Str, GetComboBoxListKey(cboCategory)) & " OriginalJCCategory" & vbCrLf
                s = s & ",i.SortOrder, i.Description, i.Comments, i.TakeoffQty, i.TakeoffUOM, i.ConversionFactor, i.OrderUOM" & vbCrLf
                s = s & ",i.Assembly, i.AssemblyDescription, i.Model, i.Location, i.RFP, i.Formula, i.Unit, i.BillingItemID, i.SalesQty, i.DivisionID" & vbCrLf
                s = s & ",i.RequestDetailID, i.Sequence" & vbCrLf
                s = s & ",i.WBS01, i.WBS02, i.WBS03, i.WBS04, i.WBS05, i.WBS06, i.WBS07, i.WBS08, i.WBS09, i.WBS10" & vbCrLf
                s = s & ",i.WBS11, i.WBS12, i.WBS13, i.WBS14, i.WBS15, i.WBS16, i.WBS17, i.WBS18, i.WBS19, i.WBS20 " & vbCrLf
                s = s & ",i.WBS21, i.WBS22, i.WBS23, i.WBS24, i.WBS25, i.WBS26, i.WBS27, i.WBS28, i.WBS29, i.WBS30 " & vbCrLf
                s = s & ",i.WBS31, i.WBS32, i.WBS33, i.WBS34, i.WBS35, i.WBS36, i.WBS37, i.WBS38, i.WBS39, i.WBS40" & vbCrLf
                s = s & "" & vbCrLf
                s = s & "from EstimateItems i " & vbCrLf
                s = s & "join EstimateAssemblies a ON(i.EstAssemblyID=a.EstAssemblyID)" & vbCrLf
                s = s & "join EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                s = s & "join NextPOChangeOrders n on(i.divisionid=n.divisionid and i.ponumber=n.ponumber)" & vbCrLf
                s = s & "join PurchaseOrders p on i.divisionid=p.divisionid and i.ponumber=p.ponumber" & vbCrLf
                s = s & "left join invoiceitems ii on p.divisionid=ii.divisionid and p.ponumber=ii.commitment" & vbCrLf
                s = s & "" & vbCrLf
                s = s & "where ii.invoice is null" & vbCrLf
                s = s & "and isnull(i.pochangeordernumber,'')=''" & vbCrLf
                s = s & "and p.Cancelled=0" & vbCrLf
                s = s & "and i.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "and isnull(i.RowType,'') not in('Original - Cancelled','Original - Reversal')" & vbCrLf
                s = s & "and isnull(i.ItemReversed,0)=0" & vbCrLf
                s = s & "and isnull(i.IsReversingItem,0)=0" & vbCrLf
                
                s = s & WhereClause("v")
                s = s & "and v.povendor=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor"))) & vbCrLf
                s = s & "and i.phase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase"))) & vbCrLf
                s = s & "and i.item=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Item"))) & vbCrLf
                s = s & "and i.description=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                s = s & "and i.porate=" & DbQuote(Num, .TextMatrix(r, .ColIndex("OrigRate"))) & vbCrLf
                HFApp.SqlExec s
        
            End If
        Next
    End With
    
    Call SaveChangeOrder(SessionKey)
    
    If HFApp.Options.ValueByName("BuildProSendPOsImmediately") = "true" Then
        'get affected PO numbers
        s = "select ponumber from pomaster where ChangeOrderAudit=" & DbQuote(Str, SessionKey)
        Set rs = HFApp.SqlExec(s, dbHomefront)
        While Not rs.EOF
            POs = POs & "," & DbQuote(Str, "" & rs("ponumber"))
            rs.MoveNext
        Wend
        POs = Mid(POs, 2)
        'send to BP
        Call SendBuildProPOs("", POs)
    End If
    
    MsgBox "POs have been updated", vbInformation, App.ProductName
    
Exit Sub
eh: Call errHandler(SRCFILE & "SaveData")
Stop
Resume
End Sub

Private Sub SaveChangeOrder(SessionKey As String)
Dim s As String

    'write change order
    s = ""
    s = s & "insert pochangeorders(PONumber,DivisionID,ChangeOrder,Description,CODate,UStmp,TStmp)" & vbCrLf
    s = s & "select distinct" & vbCrLf
    s = s & " ponumber" & vbCrLf
    s = s & ",divisionid" & vbCrLf
    s = s & ",pochangeordernumber" & vbCrLf
    s = s & ",'price change'" & vbCrLf
    s = s & ",getdate()" & vbCrLf
    s = s & "," & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & ",getdate()" & vbCrLf
    s = s & "from estimateitems where changeorderaudit=" & DbQuote(Str, SessionKey)
    Call HFApp.SqlExec(s, dbHomefront)
    
    'write change order items
    s = ""
    s = s & "insert into pochangeorderitems(DivisionID, PONumber, EstItemID, ChangeOrder, Description, Qty, UOM, Rate, Pretax, TaxGroup, JCTax, NJCTax, Job, JCExtra, JCCostCode, JCCategory, LineNumber)" & vbCrLf
    s = s & "select " & vbCrLf
    s = s & " c.divisionid" & vbCrLf
    s = s & ",c.ponumber" & vbCrLf
    s = s & ",c.EstItemID" & vbCrLf
    s = s & ",c.pochangeordernumber" & vbCrLf
    s = s & ",c.description" & vbCrLf
    s = s & ",c.poqty" & vbCrLf
    s = s & ",c.orderuom" & vbCrLf
    s = s & ",c.porate" & vbCrLf
    s = s & ",c.popretax" & vbCrLf
    s = s & ",c.potaxgroup" & vbCrLf
    s = s & ",c.pojctax" & vbCrLf
    s = s & ",c.ponjctax" & vbCrLf
    s = s & ",c.job" & vbCrLf
    s = s & ",c.jcextra" & vbCrLf
    s = s & ",c.jccostcode" & vbCrLf
    If PostPOQtyToAccounting Then
        s = s & ",c.jccategory" & vbCrLf
    Else
        s = s & ",c.variancejccategory" & vbCrLf
    End If
    s = s & ",i.LineNumber" & vbCrLf
    s = s & "from estimateitems c " & vbCrLf
    s = s & "join poitems i on c.ChangeOrderOrigEstItemId=i.estitemid and i.ponumber=c.ponumber" & vbCrLf
    s = s & "where c.changeorderaudit=" & DbQuote(Str, SessionKey)
    Call HFApp.SqlExec(s, dbHomefront)

    'update adjusted columns on poitems table
    s = ""
    s = s & "update pi set " & vbCrLf
    s = s & " adjrate = pi.rate+i.porate" & vbCrLf
    s = s & ",adjpretax = pi.pretax+i.popretax" & vbCrLf
    s = s & ",adjjctax = pi.jctax+i.pojctax" & vbCrLf
    s = s & ",adjnjctax = pi.njctax+i.ponjctax" & vbCrLf
    s = s & "from estimateitems i" & vbCrLf
    s = s & "join poitems pi on pi.estitemid=i.changeorderorigestitemid" & vbCrLf
    s = s & "where i.changeorderaudit=" & DbQuote(Str, SessionKey)
    Call HFApp.SqlExec(s, dbHomefront)


End Sub


Private Function QueryItems() As String
    Dim s As String
    s = ""
    s = s & "select distinct" & vbCrLf
    s = s & " v.povendor Vendor,v.povendorname Name,i.Phase,i.Item,i.Description,i.porate OrigRate" & vbCrLf
's = s & " v.povendor Vendor,v.povendorname Name,i.Phase,i.Item,i.Description,pi.AdjRate OrigRate--,i.porate OrigRate" & vbCrLf
    
    s = s & ",dbo.Purch_GetItemRate(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,v.Sequence,v.POVendor,getdate(),i.divisionid) NewRate" & vbCrLf
    s = s & "from EstimateItems i " & vbCrLf
    s = s & "join EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
    s = s & "join PurchaseOrders p on i.divisionid=p.divisionid and i.ponumber=p.ponumber" & vbCrLf
    s = s & "left join invoiceitems ii on p.divisionid=ii.divisionid and p.ponumber=ii.commitment" & vbCrLf
    
's = s & "join poitems pi on i.estitemid=pi.estitemid" & vbCrLf
    
    s = s & "where ii.invoice is null" & vbCrLf
s = s & "and p.IsRepriced=0" & vbCrLf
    s = s & "and isnull(i.RowType,'') not in('Original - Cancelled','Original - Reversal')" & vbCrLf
    s = s & "and isnull(i.pochangeordernumber,'')=''" & vbCrLf
    s = s & "and p.Cancelled=0" & vbCrLf
    s = s & "and p.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & WhereClause("v")
    
    QueryItems = s
End Function




Private Sub LoadCustomDescriptions()
    Label2(0).Caption = FMain.CD_Community
End Sub



Private Function QueryPOs() As String
    Dim s As String
    s = ""
    s = s & "select p.Community,p.CommunityDesc,p.Job,p.JobDesc,p.Vendor,p.VendorDesc,p.POGroup,p.POIndex,p.POIndexDescription,p.PONumber,cast(p.PODate as date) PODate" & vbCrLf
    s = s & "from purchaseorders p" & vbCrLf
    s = s & "left join invoiceitems i on p.divisionid=i.divisionid and p.ponumber=i.commitment" & vbCrLf
    s = s & "where i.invoice is null" & vbCrLf
s = s & "and p.IsRepriced=0" & vbCrLf
    s = s & "and p.Cancelled=0" & vbCrLf
    s = s & "and p.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & WhereClause("p")
    s = s & "order by 1,2,3,4,5,6,7" & vbCrLf
    QueryPOs = s
End Function

