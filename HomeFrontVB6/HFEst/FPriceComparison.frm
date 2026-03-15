VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FPriceComparison 
   Caption         =   "Price Comparison"
   ClientHeight    =   7230
   ClientLeft      =   7545
   ClientTop       =   3600
   ClientWidth     =   9930
   Icon            =   "FPriceComparison.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9930
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   8010
      Picture         =   "FPriceComparison.frx":058A
      TabIndex        =   21
      ToolTipText     =   "Cancel"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.PictureBox picLegend 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   150
      ScaleHeight     =   2025
      ScaleWidth      =   9675
      TabIndex        =   4
      Top             =   4440
      Width           =   9675
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Best fit for selected vendor and assembly."
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
         Index           =   21
         Left            =   5940
         TabIndex        =   32
         Top             =   360
         Width           =   3585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pricing for differrent assembly."
         ForeColor       =   &H80000011&
         Height          =   195
         Index           =   20
         Left            =   5940
         TabIndex        =   31
         Top             =   990
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pricing for alternate vendor."
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   15
         Left            =   5940
         TabIndex        =   30
         Top             =   780
         Width           =   1950
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Legend:"
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
         Index           =   1
         Left            =   5850
         TabIndex        =   29
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Other pricing for this vendor and assembly."
         Height          =   195
         Index           =   7
         Left            =   5940
         TabIndex        =   28
         Top             =   570
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7 - other"
         Height          =   195
         Index           =   14
         Left            =   180
         TabIndex        =   27
         Top             =   1620
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Prices for other assemblies."
         Height          =   195
         Index           =   6
         Left            =   1800
         TabIndex        =   26
         Top             =   1620
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate for this item in any community or phase."
         Height          =   195
         Index           =   13
         Left            =   1800
         TabIndex        =   17
         Top             =   1410
         Width           =   3150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Community specific rate for this item."
         Height          =   195
         Index           =   12
         Left            =   1800
         TabIndex        =   16
         Top             =   1200
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Community and phase specific rate for this item."
         Height          =   195
         Index           =   11
         Left            =   1800
         TabIndex        =   15
         Top             =   990
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Quote for this assembly in any community or phase."
         Height          =   195
         Index           =   10
         Left            =   1800
         TabIndex        =   14
         Top             =   780
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Community specific quote for this assembly."
         Height          =   195
         Index           =   9
         Left            =   1800
         TabIndex        =   13
         Top             =   570
         Width           =   3045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Community and phase specific quote for this assembly."
         Height          =   195
         Index           =   8
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   3840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate types and order of retrieval:"
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
         Left            =   90
         TabIndex        =   11
         Top             =   120
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6 - Global Rate"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   10
         Top             =   1410
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5 - Community Rate"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   9
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4 - Phase Rate"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   990
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3 - Global Quote"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2 - Community Quote"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   570
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1 - Phase Quote"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   1155
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gPrices 
      Height          =   3105
      Left            =   150
      TabIndex        =   0
      Top             =   1230
      Width           =   9075
      _cx             =   16007
      _cy             =   5477
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FPriceComparison.frx":0B14
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
      ExplorerBar     =   0
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
   Begin VB.Label lblVendor 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   6120
      TabIndex        =   34
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   195
      Index           =   22
      Left            =   5505
      TabIndex        =   33
      Top             =   270
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Community"
      Height          =   195
      Index           =   26
      Left            =   5220
      TabIndex        =   25
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Assembly"
      Height          =   195
      Index           =   25
      Left            =   5310
      TabIndex        =   24
      Top             =   750
      Width           =   675
   End
   Begin VB.Label lblCommunity 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   6090
      TabIndex        =   23
      Top             =   510
      UseMnemonic     =   0   'False
      Width           =   75
   End
   Begin VB.Label lblAssembly 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   6090
      TabIndex        =   22
      Top             =   750
      UseMnemonic     =   0   'False
      Width           =   75
   End
   Begin VB.Label lblDefaultRate 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1320
      TabIndex        =   20
      Top             =   750
      UseMnemonic     =   0   'False
      Width           =   3840
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   19
      Top             =   510
      UseMnemonic     =   0   'False
      Width           =   3840
   End
   Begin VB.Label lblPhaseItem 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1320
      TabIndex        =   18
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   3840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Rate"
      Height          =   195
      Index           =   18
      Left            =   150
      TabIndex        =   3
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   17
      Left            =   435
      TabIndex        =   2
      Top             =   510
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phase/Item"
      Height          =   195
      Index           =   16
      Left            =   405
      TabIndex        =   1
      Top             =   270
      Width           =   825
   End
End
Attribute VB_Name = "FPriceComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FPriceListUpdate::"

Public Sub ShowForm(Vendor As String, community As String, CommunityPhase As String, Assembly As String, Model As String, Phase As String, Item As String, ItemDesc As String)
On Error Resume Next
    Dim b As Boolean
    Dim s As String
    Dim rs As Recordset
    
Screen.MousePointer = vbHourglass
    
    s = ""
    s = s & "SELECT i.Price" & vbCrLf
    s = s & "      ,i.OrderUOM" & vbCrLf
    s = s & "      ,c.Description + ISNULL(' ' + p.Description,'') Community" & vbCrLf
    s = s & "      ,' -- ' + a.Description AssemblyDesc" & vbCrLf
    s = s & "      ,v.Vendor_Name" & vbCrLf
    s = s & "from tblphaseitem i" & vbCrLf
    s = s & "     left outer join tblVendors v on(v.DivisionID = " & HFApp.DivisionID & " and v.vendor_id=" & DbQuote(str, Vendor) & ")" & vbCrLf
    s = s & "     left outer join tbllocality c on (c.area=" & DbQuote(str, community) & ")" & vbCrLf
    s = s & "     left outer join communityphase p on (p.community=" & DbQuote(str, community) & " and p.communityphase=" & DbQuote(str, CommunityPhase) & ")" & vbCrLf
    s = s & "     left outer join DistinctAssemblies a on(a.assembly=" & DbQuote(str, Assembly) & ")" & vbCrLf
    s = s & "where i.DivisionID = " & HFApp.DivisionID & " And i.Phase = " & DbQuote(str, Phase) & "" & vbCrLf
    s = s & "  and item=" & DbQuote(str, Item) & "" & vbCrLf
    
    Set rs = HFApp.SqlExec(s)
    lblPhaseItem.Caption = Phase & " / " + Item
    lblDescription.Caption = ItemDesc
    lblDefaultRate.Caption = format("" & rs("Price"), "$#,##0.00##########") & " per " & rs("OrderUOM")
    lblVendor.Caption = "" & rs("Vendor_Name")
    lblCommunity.Caption = "" & rs("Community")
    lblAssembly.Caption = Assembly & rs("AssemblyDesc")
    
    
    
    s = ""
    s = s & "select case when isnull(vc.community,'')=" & DbQuote(str, community) & " and isnull(vc.communityphase,'')=" & DbQuote(str, CommunityPhase) & " and isnull(vc.assembly,'')=" & DbQuote(str, Assembly) & " and isnull(vc.model,'')=" & DbQuote(str, Model) & " then  '1 - Phase Quote' " & vbCrLf
    s = s & "            when isnull(vc.community,'')=" & DbQuote(str, community) & " and isnull(vc.communityphase,'')=''  and isnull(vc.assembly,'')=" & DbQuote(str, Assembly) & " and isnull(vc.model,'')=" & DbQuote(str, Model) & " then  '2 - Community Quote'" & vbCrLf
    s = s & "            when isnull(vc.community,'')=''    and isnull(vc.communityphase,'')=''  and isnull(vc.assembly,'')=" & DbQuote(str, Assembly) & " and isnull(vc.model,'')=" & DbQuote(str, Assembly) & " then  '3 - Global Quote'" & vbCrLf
    s = s & "            when isnull(vc.community,'')=" & DbQuote(str, community) & " and isnull(vc.communityphase,'')=" & DbQuote(str, CommunityPhase) & " and isnull(vc.assembly,'')=''         and isnull(vc.model,'')=''         then  '4 - Phase Rate'" & vbCrLf
    s = s & "            when isnull(vc.community,'')=" & DbQuote(str, community) & " and isnull(vc.communityphase,'')=''  and isnull(vc.assembly,'')=''         and isnull(vc.model,'')=''         then  '5 - Community Rate'" & vbCrLf
    s = s & "            when isnull(vc.community,'')=''    and isnull(vc.communityphase,'')=''  and isnull(vc.assembly,'')=''         and isnull(vc.model,'')=''         then  '6 - Global Rate' " & vbCrLf
    s = s & "            else '7 - other' end" & vbCrLf
    s = s & "      ,vc.Vendor" & vbCrLf
    s = s & "      ,Vendor_Name" & vbCrLf
    s = s & "      ,Current_cost" & vbCrLf
    s = s & "      ,c.Description + ISNULL(' ' + p.Description,'')" & vbCrLf
    s = s & "      ,vc.assembly + ISNULL(' -- ' + a.Description,'')" & vbCrLf
    s = s & "from tblvendorcost vc" & vbCrLf
    s = s & "     left outer join DistinctAssemblies a on(a.assembly=vc.assembly)" & vbCrLf
    s = s & "     left outer join tblVendors v on(vc.DivisionID = v.DivisionID and vc.vendor=v.vendor_id and v.inactive=0)" & vbCrLf
    s = s & "     left outer join tbllocality c on (c.area=vc.community)" & vbCrLf
    s = s & "     left outer join communityphase p on (p.community=vc.community and p.communityphase=vc.communityphase)" & vbCrLf
    s = s & "where vc.phase=" & DbQuote(str, Phase) & "" & vbCrLf
    s = s & "  and vc.item=" & DbQuote(str, Item) & "" & vbCrLf
    s = s & "  and (isnull(vc.community,'')=" & DbQuote(str, community) & " or isnull(vc.community,'')='')" & vbCrLf
    s = s & "  and (isnull(vc.communityphase,'')=" & DbQuote(str, CommunityPhase) & " or isnull(vc.communityphase,'')='')" & vbCrLf
    s = s & "  and vc.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "order by 1,2,3,4,5" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With Me.gPrices
        .Rows = 1
        While Not rs.EOF
            Call .AddItem("" & rs(0) & vbTab & rs(2) & vbTab & format("" & rs(3), "#,##0.00##########") & vbTab & rs(4) & vbTab & rs(5))
            
            Select Case True
                Case "" & rs(0) = "7 - other"
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .cols - 1) = vbGrayText
                Case "" & rs(1) = Vendor 'highlight first vendor price
                    If Not b Then
                        b = True
                        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .cols - 1) = True
                    End If
                Case Else ' highlight other vendors
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .cols - 1) = vbHighlight
            End Select
            
            
            Call rs.MoveNext
        Wend
        Call .AutoSize(0, .cols - 1)
    End With
    
    
Screen.MousePointer = vbDefault
        

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
    
    gPrices.Move margin, gPrices.Top, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - gPrices.Top - 3 * margin - cmdNav(1).Height - picLegend.Height
    picLegend.Move margin, gPrices.Top + gPrices.Height + margin, gPrices.Width, Me.ScaleHeight - gPrices.Top - gPrices.Height - 3 * margin - cmdNav(1).Height
    cmdNav(1).Move Me.ScaleWidth - margin - cmdNav(1).Width, Me.ScaleHeight - margin - cmdNav(1).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

