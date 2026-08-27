VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FAssemblyCosts 
   Caption         =   "FAssemblyCosts"
   ClientHeight    =   7224
   ClientLeft      =   1752
   ClientTop       =   1548
   ClientWidth     =   9936
   Icon            =   "FAssemblyCosts.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7224
   ScaleWidth      =   9936
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   5415
      Left            =   150
      TabIndex        =   0
      Top             =   1230
      Width           =   9075
      _cx             =   1986412167
      _cy             =   1986405711
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      SelectionMode   =   1
      GridLines       =   1
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
      FormatString    =   $"FAssemblyCosts.frx":058A
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
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   8010
      Picture         =   "FAssemblyCosts.frx":071B
      TabIndex        =   7
      ToolTipText     =   "Cancel"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   96
      Picture         =   "FAssemblyCosts.frx":0CA5
      Top             =   120
      Width           =   384
   End
   Begin VB.Label lblAssembly 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2070
      TabIndex        =   6
      Top             =   750
      UseMnemonic     =   0   'False
      Width           =   75
   End
   Begin VB.Label lblModelOption 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2070
      TabIndex        =   5
      Top             =   510
      UseMnemonic     =   0   'False
      Width           =   45
   End
   Begin VB.Label lblCommunity 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2070
      TabIndex        =   4
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Assembly"
      Height          =   195
      Index           =   18
      Left            =   1320
      TabIndex        =   3
      Top             =   750
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Model/Option"
      Height          =   195
      Index           =   17
      Left            =   1005
      TabIndex        =   2
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Community/Phase"
      Height          =   195
      Index           =   16
      Left            =   690
      TabIndex        =   1
      Top             =   270
      Width           =   1290
   End
End
Attribute VB_Name = "FAssemblyCosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FAssemblyCosts::"

Public Sub ShowForm(WorksheetID As Long, community As String, CommunityPhase As String, Model As String, OptionID As String, Assembly As String)
On Error Resume Next
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    Dim sum As Double
    
    Screen.MousePointer = vbHourglass
    
    s = ""
    s = s & "select community + isnull('.' + communityphase,'') + ' -- ' + CommunityDesc + isnull(' ' + CommunityPhaseDesc,'')" & vbCrLf
    s = s & "      ,model + isnull(' / ' + nullif(optionid,''),11)" & vbCrLf
    s = s & "      ,AssemblyDesc" & vbCrLf
    s = s & "  from assemblylist" & vbCrLf
    s = s & " where community=" & DbQuote(Str, community) & vbCrLf
    s = s & "   and communityphase=" & DbQuote(Str, CommunityPhase) & vbCrLf
    s = s & "   and assembly=" & DbQuote(Str, Assembly) & vbCrLf
    s = s & "   and model=" & DbQuote(Str, Model) & vbCrLf
    s = s & "   and optionid=" & DbQuote(Str, OptionID) & vbCrLf
    s = s & "   and DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s, dbHomefront)
    If Not rs.EOF Then
        lblCommunity.Caption = "" & rs(0)
        lblModelOption.Caption = "" & rs(1)
        lblAssembly.Caption = "" & rs(2)
    End If
    
    
    s = ""
    s = s & "SELECT i.poindex POIndex" & vbCrLf
    s = s & "      ,ISNULL(p.description,'** unknown **') PODesc" & vbCrLf
    s = s & "      ,c.Phase" & vbCrLf
    s = s & "      ,c.Item" & vbCrLf
    s = s & "      ,ISNULL(i.description,'** unknown **') ItemDesc" & vbCrLf
    s = s & "      ,c.OrderQty" & vbCrLf
    s = s & "      ,c.Rate" & vbCrLf
    s = s & "      ,i.OrderUOM" & vbCrLf
    s = s & "      ,c.Vendor" & vbCrLf
    s = s & "      ,ISNULL(v.Vendor_name ,'** unknown **') VendorDesc" & vbCrLf
    s = s & "  FROM tblSalesSheetDetails d" & vbCrLf
    s = s & "       JOIN tblSalesSheetMaster m on (m.worksheet = d.Worksheet)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblSalesSheetCosts c ON(d.Worksheet=c.Worksheet and d.community=c.community and d.communityphase=c.communityphase and d.model=c.model and d.optionid=c.optionid and d.assembly=c.assembly)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPhaseItem i ON (i.DivisionID = m.DivisionID and c.Phase=i.Phase AND c.Item=i.Item)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPOIndex p ON (i.DivisionID = p.DivisionID and i.poindex=p.poindex)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblVendors v ON (c.Vendor=v.Vendor_ID and v.DivisionID = m.DivisionID)" & vbCrLf
    s = s & " WHERE d.Worksheet=" & DbQuote(Num, WorksheetID) & vbCrLf
    s = s & "   and d.community=" & DbQuote(Str, community) & vbCrLf
    s = s & "   and d.communityphase=" & DbQuote(Str, CommunityPhase) & vbCrLf
    s = s & "   and d.model=" & DbQuote(Str, Model) & vbCrLf
    s = s & "   AND d.optionid=" & DbQuote(Str, OptionID) & vbCrLf
    s = s & "   and d.assembly=" & DbQuote(Str, Assembly) & vbCrLf
    s = s & "   and m.DivisionID = " & HFApp.DivisionID
    s = s & "order by p.poindex,c.sequence" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gData
        .Rows = 1
        r = 1
        sum = 0
        While Not rs.EOF
            Call .AddItem("")
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(r, .ColIndex("PODesc")) = "" & rs("PODesc")
            .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
            .TextMatrix(r, .ColIndex("ItemDesc")) = "" & rs("ItemDesc")
            .TextMatrix(r, .ColIndex("OrderQty")) = "" & rs("OrderQty")
            .TextMatrix(r, .ColIndex("Rate")) = "" & rs("Rate")
            .TextMatrix(r, .ColIndex("Total")) = Round(Val("" & rs("OrderQty")) * Val("" & rs("Rate")), 2)
            sum = sum + Round(Val("" & rs("OrderQty")) * Val("" & rs("Rate")), 2)
            
            .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(r, .ColIndex("Vendor")) = "" & rs("Vendor")
            .TextMatrix(r, .ColIndex("VendorDesc")) = "" & rs("VendorDesc")
            r = r + 1
            Call rs.MoveNext
        Wend
        Call .AddItem("")
        .TextMatrix(r, .ColIndex("Total")) = sum
        .Cell(flexcpFontBold, r, .ColIndex("Total")) = True
        End With
    
    
    Screen.MousePointer = vbDefault
        

    Call IniGetGrid(Me, gData)
    Me.Show vbModal
End Sub
                
                


Private Sub cmdNav_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    gData.Move margin, gData.Top, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - gData.Top - 2 * margin - cmdNav(1).Height
    cmdNav(1).Move Me.ScaleWidth - margin - cmdNav(1).Width, Me.ScaleHeight - margin - cmdNav(1).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton Then
        If gData.MouseRow = 0 Then
            Cancel = True
            Call FMain.ShowColumnMenu(gData)
        End If
    End If
End Sub

