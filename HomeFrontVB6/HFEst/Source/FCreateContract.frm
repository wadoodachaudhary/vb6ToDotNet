VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FCreateContract 
   Caption         =   "Add To Contract"
   ClientHeight    =   8535
   ClientLeft      =   8595
   ClientTop       =   2355
   ClientWidth     =   13170
   Icon            =   "FCreateContract.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   13170
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   5475
      Width           =   16305
      Begin VB.CommandButton cmdNav 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   10620
         Picture         =   "FCreateContract.frx":000C
         TabIndex        =   13
         ToolTipText     =   "Login"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   1
         Left            =   11850
         Picture         =   "FCreateContract.frx":0596
         TabIndex        =   12
         ToolTipText     =   "Cancel"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame frmChangeorder 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   11775
         Begin VB.TextBox txtDescription 
            Height          =   615
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   6885
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Change Order Comments "
            Height          =   435
            Left            =   120
            TabIndex        =   11
            Top             =   180
            Width           =   1110
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gBillingItems 
      Height          =   3885
      Left            =   7380
      TabIndex        =   2
      Top             =   360
      Width           =   4335
      _cx             =   7646
      _cy             =   6853
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FCreateContract.frx":0B20
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   7
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
      ExplorerBar     =   0
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   915
      Left            =   4860
      TabIndex        =   4
      Top             =   3705
      Width           =   2445
      Begin VB.OptionButton optAddon 
         Caption         =   "Billed individually"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   630
         Value           =   -1  'True
         Width           =   2565
      End
      Begin VB.OptionButton optAddon 
         Caption         =   "Prorated (hidden)"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   390
         Width           =   2595
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Addons"
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
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   645
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gAssemblies 
      Height          =   3885
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   4335
      _cx             =   7646
      _cy             =   6853
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
      AllowUserResizing=   0
      SelectionMode   =   3
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
      FormatString    =   $"FCreateContract.frx":0BFC
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   7
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
      ExplorerBar     =   0
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
   Begin VSFlex8Ctl.VSFlexGrid gGrouping 
      Height          =   2850
      Left            =   4860
      TabIndex        =   15
      Top             =   915
      Width           =   2175
      _cx             =   3836
      _cy             =   5027
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FCreateContract.frx":0D3C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   7
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Items Preview"
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
      Left            =   7380
      TabIndex        =   3
      Top             =   120
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Level"
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
      Left            =   4860
      TabIndex        =   1
      Top             =   630
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unassigned Costs"
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
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "FCreateContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FCreateContract::"
Public EventTraps As Collection

Private Type ViewDefs
    Name         As String
    KeyFlds      As String
    SortFlds     As String
    DisplayFlds  As String
    IconKeys     As String
End Type
Private mViews() As ViewDefs
Private mViewIndex As Long
Private mViewQuery As String

Private mCancel As Boolean
Private mJob As String
Private mChangeOrder  As Boolean

Private Sub LoadViews()
    Dim i As Long
    
    
    If mChangeOrder Then
    
        ReDim mViews(0) As ViewDefs
        
        mViews(i).Name = "Change Requests"
        mViews(i).KeyFlds = "CORSort,EstAssemblyID"
        mViews(i).DisplayFlds = "CORNumber,HFDescription"
        mViews(i).SortFlds = mViews(i).KeyFlds
        mViews(i).IconKeys = "option,assembly"
        i = i + 1
        
    Else
    
        ReDim mViews(5) As ViewDefs
        
        mViews(i).Name = "Customer"
        mViews(i).KeyFlds = "Customer_No,CORSort,EstAssemblyID"
        mViews(i).DisplayFlds = "CustomerDesc,CORNumber,HFDescription"
        mViews(i).SortFlds = mViews(i).KeyFlds
        mViews(i).IconKeys = "customer,option,assembly"
        i = i + 1
        
        mViews(i).Name = "PO Group"
        mViews(i).KeyFlds = "POGroup,POIndex,PONumberVendor"
        mViews(i).DisplayFlds = "POGroup,case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,PONumber + ' - ' + POVendorName"
        mViews(i).SortFlds = mViews(i).KeyFlds
        mViews(i).IconKeys = "folder,purchaseorder,sendpos"
        i = i + 1
    
        mViews(i).Name = "PO Index"
        mViews(i).KeyFlds = "POIndex,PONumberVendor"
        mViews(i).DisplayFlds = "case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end,PONumber + ' - ' + POVendorName"
        mViews(i).SortFlds = mViews(i).KeyFlds
        mViews(i).IconKeys = "purchaseorder,sendpos"
        i = i + 1
    
        mViews(i).Name = "Purchase Orders"
        mViews(i).KeyFlds = "PONumber"
        mViews(i).DisplayFlds = "PONumber + CASE PONumber WHEN '' THEN '' ELSE ' -- ' + POVendorName END"
        mViews(i).SortFlds = mViews(i).KeyFlds
        mViews(i).IconKeys = "sendpos"
        i = i + 1
        
        mViews(i).Name = "Vendor"
        mViews(i).KeyFlds = "POVendor,PONumber"
        mViews(i).DisplayFlds = "POVendorName,PONUmber + ' - ' + case when POIndex=POIndexDescription then POIndex else POIndex + ' ' + POIndexDescription end"
        mViews(i).SortFlds = mViews(i).KeyFlds
        mViews(i).IconKeys = "vendor,sendpos"
        i = i + 1
    
        mViews(i).Name = "Cost Code"
        mViews(i).KeyFlds = "JCExtra,JCCostCode,JCCategory"
        If HFApp.Options(AccountingSystem) = asquickbooks Then
            mViews(i).DisplayFlds = "JCExtra,JCCostCodeDesc,JCCategoryDesc"
        Else
            mViews(i).DisplayFlds = "JCExtra,JCCostCode+' - '+JCCostCodeDesc,JCCategoryDesc"
        End If
        mViews(i).SortFlds = mViews(i).KeyFlds
        mViews(i).IconKeys = "folder,folder,folder"
        i = i + 1
    End If
    
End Sub


Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then Call SaveData
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadViews
    Call LoadFields
        
    mViewIndex = Val(IniGet(AppIni, "Options", "BillingItemsView"))
    If mChangeOrder Then mViewIndex = 0
    frmChangeorder.Visible = mChangeOrder
    
    Call LoadAssemblies(True)
    Call LoadItems
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    Const margin = 270
    
    Frame2.Move 0, Me.ScaleHeight - Frame2.Height, Me.ScaleWidth
    frmChangeorder.Move 0, 0, Frame2.Width
    cmdNav(0).Left = Me.ScaleWidth - 2 * cmdNav(0).Width - margin - 30
    cmdNav(1).Left = Me.ScaleWidth - 1 * cmdNav(0).Width - margin
    txtDescription.Width = cmdNav(0).Left - txtDescription.Left - margin
    gBillingItems.Move gBillingItems.Left, gBillingItems.Top, Me.ScaleWidth - gBillingItems.Left - margin, Me.ScaleHeight - gBillingItems.Top - Frame2.Height
    gAssemblies.Height = gBillingItems.Height


    With gBillingItems
        Call .AutoSize(0, .cols - 1)
        .ColWidth(.ColIndex("description")) = .Width - .ColWidth(.ColIndex("price")) - 360
    End With


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Call IniPutForm(Me)
    Call IniPut(AppIni, "Options", "BillingItemsView", mViewIndex)
    For i = 0 To gGrouping.Rows - 1
        Call IniPut(AppIni, "Options", "BillingLevel" & i, IIf(gGrouping.Cell(flexcpChecked, i, 0) = flexChecked, 1, 0))
    Next
End Sub

Public Function Edit(ChangeOrder As Boolean, Job As String) As Boolean
    mJob = Job
    mChangeOrder = ChangeOrder
    mCancel = True
    Me.Show vbModal
    Edit = Not mCancel
End Function

Private Sub LoadItems()
    Dim s As String
    Dim selectStr As String
    Dim groupStr As String
    
    Dim rs As Recordset
    Dim rsJobs As Recordset
    Dim i As Long
    
    'for proration
    Dim Total As Double
    Dim AddonApplied As Double
    Dim AddonTotal As Double
    
    
    gBillingItems.Rows = 1
   
    
    selectStr = ""
    With gGrouping
    For i = 0 To .Rows - 1
        If .Cell(flexcpChecked, i, 0) = flexChecked Then
            selectStr = selectStr & " + ', ' + isnull(" & .Cell(flexcpText, i, 1) & ",'')"
            groupStr = groupStr & ",isnull(" & .Cell(flexcpText, i, 1) & ",'')"
        End If
    Next
    End With
    If groupStr = "" Then Exit Sub
        
    
    s = ""
    s = s & "select artaxgroup,holdbackrate," & Mid(selectStr, 11) & vbCrLf
    s = s & "      ,sum(isnull(popretax,0)+isnull(pojctax,0)),revenueaccount" & vbCrLf
    s = s & "  from estimateditems" & vbCrLf
    If mViewIndex = 0 Then
        If mChangeOrder Then
            s = s & " where ischangerequest=1 and " & WhereClause & vbCrLf
        Else
            s = s & " where (ischangerequest=0 or ischange=0) and " & WhereClause & vbCrLf
        End If
    Else
        s = s & " where ischange=0 and " & WhereClause & vbCrLf
    End If
    s = s & "   and billingitemid=0" & vbCrLf
    s = s & "group by revenueaccount,artaxgroup,holdbackrate," & Mid(groupStr, 2) & vbCrLf
    If optAddon(1) Then
        s = s & "union all" & vbCrLf
        s = s & "select distinct artaxgroup,holdbackrate,a.title,a.amount,revenueaccount" & vbCrLf
        s = s & "  from estimateditems i join addons a on a.assemblyid=i.estassemblyid" & vbCrLf
        If mChangeOrder Then
            s = s & " where i.ischangerequest=1 and " & WhereClause("i") & vbCrLf
        Else
            s = s & " where i.ischangerequest=0 and " & WhereClause("i") & vbCrLf
        End If
        s = s & "   and i.billingitemid=0" & vbCrLf
        s = s & "   and a.Basis<>'Cost' and a.Basis<>'SubTotal'" & vbCrLf
        s = s & "   and a.amount<>0" & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select distinct artaxgroup,holdbackrate,a.title,a.amount,revenueaccount" & vbCrLf
        s = s & "  from estimateditems i join addons a on a.job=i.job_no" & vbCrLf
        s = s & " where i.ischangerequest=0 and " & WhereClause("i") & vbCrLf
        s = s & "   and i.billingitemid=0" & vbCrLf
        s = s & "   and a.Job=" & DbQuote(str, mJob) & vbCrLf
        s = s & "   and a.assemblyid=0" & vbCrLf
        s = s & "   and a.Basis<>'Cost' and a.Basis<>'SubTotal'" & vbCrLf
        s = s & "   and a.amount<>0" & vbCrLf
    End If
    s = s & "order by 3"
    Set rs = HFApp.SqlExec(s)
    With gBillingItems
        .Rows = 1
        Total = 0
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, .ColIndex("taxgroup")) = "" & rs("artaxgroup")
            .TextMatrix(.Rows - 1, .ColIndex("revenueaccount")) = "" & rs("revenueaccount")
            .TextMatrix(.Rows - 1, .ColIndex("holdbackrate")) = "" & rs("holdbackrate")
            
            .TextMatrix(.Rows - 1, .ColIndex("Description")) = "" & rs(2)
            .TextMatrix(.Rows - 1, .ColIndex("Price")) = format("" & rs(3), "#,###.00")
            Total = Total + Val("" & rs(3))
            
            If mChangeOrder Then
                .TextMatrix(.Rows - 1, .ColIndex("whereclause")) = WhereClause & " and ischangerequest=1 and " & Mid(selectStr, 11) & " = " & DbQuote(str, "" & rs(2))
            Else
                .TextMatrix(.Rows - 1, .ColIndex("whereclause")) = WhereClause & " and " & Mid(selectStr, 11) & " = " & DbQuote(str, "" & rs(2))
            End If
            
            rs.MoveNext
        Wend
        
        If Total <> 0 Then
        If Not mChangeOrder And optAddon(0) Then
            'update amounts... prorate total addon value
            
            s = ""
            s = s & "select " & vbCrLf
            s = s & "isnull((select sum(distinct a.amount)" & vbCrLf
            s = s & "  from estimateditems i join addons a on a.assemblyid=i.estassemblyid" & vbCrLf
            s = s & "  where i.ischangerequest=0 and " & WhereClause("i") & vbCrLf
            s = s & "   and i.billingitemid=0" & vbCrLf
            s = s & "   and a.Basis<>'Cost' and a.Basis<>'SubTotal'" & vbCrLf
            s = s & "   and a.amount<>0),0)" & vbCrLf
            s = s & "+" & vbCrLf
            s = s & "isnull((select sum(distinct a.amount)" & vbCrLf
            s = s & "  from estimateditems i join addons a on a.job=i.job_no" & vbCrLf
            s = s & "  where i.ischangerequest=0 and " & WhereClause("i") & vbCrLf
            s = s & "   and i.billingitemid=0" & vbCrLf
            s = s & "   and a.Job=" & DbQuote(str, mJob) & vbCrLf
            s = s & "   and a.assemblyid=0" & vbCrLf
            s = s & "   and a.Basis<>'Cost' and a.Basis<>'SubTotal'" & vbCrLf
            s = s & "   and a.amount<>0),0)" & vbCrLf
            On Error Resume Next
            AddonTotal = 0
            AddonApplied = 0
            AddonTotal = Val(HFApp.SqlExec(s)(0))
            On Error GoTo 0
            If AddonTotal <> 0 Then
                For i = 1 To .Rows - 1
                    AddonApplied = AddonApplied + Round(Val(.ValueMatrix(i, 1)) / Total * AddonTotal, 2)
                    .TextMatrix(i, .ColIndex("Price")) = format(Val(.ValueMatrix(i, .ColIndex("Price"))) + Round(Val(.ValueMatrix(i, .ColIndex("Price"))) / Total * AddonTotal, 2), "#,###.00")
                Next
                .TextMatrix(.Rows - 1, .ColIndex("Price")) = format(Val(.ValueMatrix(.Rows - 1, .ColIndex("Price"))) + AddonTotal - AddonApplied, "#,###.00")
            End If
            
        Else 'If mChangeOrder And optAddon(0) Then
'             'update amounts... prorate total addon value
'
'            s = ""
'            s = s & "select " & vbCrLf
'            s = s & "isnull((select sum(distinct a.amount)" & vbCrLf
'            s = s & "  from estimateditems i join addons a on a.assemblyid=i.estassemblyid" & vbCrLf
'            s = s & "  where i.ischangerequest=1 and " & WhereClause("i") & vbCrLf
'            s = s & "   and i.billingitemid=0" & vbCrLf
'            s = s & "   and a.Basis<>'Cost' and a.Basis<>'SubTotal'" & vbCrLf
'            s = s & "   and a.amount<>0),0)" & vbCrLf
'            s = s & "+" & vbCrLf
'            s = s & "isnull((select sum(distinct a.amount)" & vbCrLf
'            s = s & "  from estimateditems i join addons a on a.job=i.job_no" & vbCrLf
'            s = s & "  where i.ischangerequest=0 and " & WhereClause("i") & vbCrLf
'            s = s & "   and i.billingitemid=0" & vbCrLf
'            s = s & "   and a.Job=" & DbQuote(Str, mJob) & vbCrLf
'            s = s & "   and a.assemblyid=0" & vbCrLf
'            s = s & "   and a.Basis<>'Cost' and a.Basis<>'SubTotal'" & vbCrLf
'            s = s & "   and a.amount<>0),0)" & vbCrLf
'            On Error Resume Next
'            AddonTotal = 0
'            AddonApplied = 0
'            AddonTotal = Val(HFApp.SqlExec(s)(0))
'            On Error GoTo 0
'            If AddonTotal <> 0 Then
'                For i = 1 To .Rows - 1
'                    AddonApplied = AddonApplied + Round(Val(.ValueMatrix(i, 1)) / Total * AddonTotal, 2)
'                    .TextMatrix(i, .ColIndex("Price")) = format(Val(.ValueMatrix(i, .ColIndex("Price"))) + Round(Val(.ValueMatrix(i, .ColIndex("Price"))) / Total * AddonTotal, 2), "#,###.00")
'                Next
'                .TextMatrix(.Rows - 1, .ColIndex("Price")) = format(Val(.ValueMatrix(.Rows - 1, .ColIndex("Price"))) + AddonTotal - AddonApplied, "#,###.00")
'            End If
        End If
        
        
        If mChangeOrder And optAddon(0) Then
            'update amounts... prorate total addon value
            s = ""
            s = s & "select sum(amount)" & vbCrLf
            s = s & "  from addons" & vbCrLf
            s = s & " where Basis<>'Cost' and a.Basis<>'SubTotal' and isnull(title,'')<>''" & vbCrLf
            s = s & "   and amount<>0" & vbCrLf
            s = s & "   and assemblyid in(select distinct estassemblyid" & vbCrLf
            s = s & "                       from estimateditems" & vbCrLf
            s = s & "                      where ischangerequest=1 and billingitemid=0 and " & WhereClause() & vbCrLf
            s = s & "                    )" & vbCrLf
           
            On Error Resume Next
            AddonTotal = 0
            AddonApplied = 0
            AddonTotal = Val(HFApp.SqlExec(s)(0))
            On Error GoTo 0
            If AddonTotal <> 0 Then
                For i = 1 To .Rows - 2
                    AddonApplied = AddonApplied + Round(Val(.ValueMatrix(i, 1)) / Total * AddonTotal, 2)
                    .TextMatrix(i, .ColIndex("Price")) = format(Val(.ValueMatrix(i, .ColIndex("Price"))) + Round(Val(.ValueMatrix(i, .ColIndex("Price"))) / Total * AddonTotal, 2), "#,###.00")
                Next
                .TextMatrix(.Rows - 1, .ColIndex("Price")) = format(Val(.ValueMatrix(.Rows - 1, .ColIndex("Price"))) + AddonTotal - AddonApplied, "#,###.00")
            End If
            
        
        End If
        End If
    
        Call Form_Resize
        Me.cmdNav(0).Enabled = .Rows > 1
    End With

End Sub


Private Sub LoadAssemblies(ClearTree As Boolean)
On Error GoTo eh
    Dim rs As Recordset
    Dim r As Long
    Dim s       As String
    Dim i       As Long
    Dim levels  As Long
    Dim Level   As Long
    Dim n       As VSFlexNode
    Dim WhereClause As String
    Dim KeyFld     As String
    Dim DisplayFld As String
    
    Dim SortFld    As String
    Dim KeyValue   As String
    Dim DisplayValue   As String
    Dim IconKey As String

'PROPERTY                  CONTAINS
'.rowdata()                Where clause
'.cell(flexcpText, r, 0)   DisplayFld Value
'.cell(flexcpData, r, 0)   DisplayFld Field Name (also used as tooltip)
'.Cell(flexcpText, r, 1)   KeyFld Value
'.Cell(flexcpData, r, 1)   KeyFld Field Name
'.Cell(flexcpText, r, 2)   Completed  As MultiStateEnum
'.Cell(flexcpText, r, 3)   Selected   as MultiStateEnum

    With gAssemblies
        .Redraw = flexRDNone
        
        .TextMatrix(0, 0) = mViews(mViewIndex).Name
        
        levels = Parse(mViews(mViewIndex).DisplayFlds)

        
        If ClearTree Then .Rows = 1


        If .Rows = 1 Then
            Level = -1
        Else
            Level = .RowOutlineLevel(.Row)
        End If

        If Level + 1 = levels Then
            'user has opened the lowest level so do nothing
        Else
            On Error Resume Next
            If .TextMatrix(.GetNodeRow(.Row, flexNTFirstChild), 0) = "dummy" Then
                .RemoveItem .GetNodeRow(.Row, flexNTFirstChild)
            Else
                .Redraw = flexRDBuffered
                Exit Sub
            End If
            On Error GoTo eh

            KeyFld = Parse(mViews(mViewIndex).KeyFlds, Level + 2)
            DisplayFld = Parse(mViews(mViewIndex).DisplayFlds, Level + 2)
            IconKey = Parse(mViews(mViewIndex).IconKeys, Level + 2)
            SortFld = Parse(mViews(mViewIndex).SortFlds, Level + 2)
            

            s = ""
            s = s & "SELECT " & KeyFld & "," & DisplayFld & vbCrLf
            s = s & "      ,COUNT(*) TotalItems" & vbCrLf
            s = s & "      ,SUM(CASE WHEN IsChangeRequest=0 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) Budgeted" & vbCrLf
            s = s & "      ,SUM(CASE WHEN IsChangeRequest=1 and BudgetDeleted<>1 THEN BudgetPretax + BudgetJCTax ELSE 0 END) Changes" & vbCrLf
            s = s & "      ,SUM(CASE WHEN POGenBatch=0 or poDeleted=1 THEN 0 ELSE POPretax+POJCTax END) Commited" & vbCrLf
            s = s & "      ,MIN(CAST(BudgetsLocked AS INT)) BudgetsLocked" & vbCrLf
            s = s & "      ,MIN(HFLocation) HFLocation" & vbCrLf
            s = s & "      ,MIN(IsChangeRequest) IsChangeRequest" & vbCrLf
            s = s & "  FROM Estimateditems" & vbCrLf
            On Error Resume Next
            WhereClause = .RowData(.Row)
            On Error GoTo eh
            
            
            
            If WhereClause = "" Then
                If mViewIndex = 0 Then
                    If mChangeOrder Then
                        s = s & " WHERE isnull(IsChangeRequest,0) = 1 and isnull(billingitemid,0)=0 and Job_No=" & DbQuote(str, mJob) & vbCrLf
                    Else
                        s = s & " WHERE (isnull(IsChangeRequest,0) = 1 or isnull(IsChange,0) = 0) and isnull(billingitemid,0)=0 and Job_No=" & DbQuote(str, mJob) & vbCrLf
                    End If
                Else
                    s = s & " WHERE isnull(ischange,0)=0 and isnull(billingitemid,0)=0 and Job_No=" & DbQuote(str, mJob) & vbCrLf
                End If
            Else
                If mViewIndex = 0 Then
                    If mChangeOrder Then
                        s = s & " WHERE isnull(IsChangeRequest,0)= 1 and isnull(billingitemid,0)=0 and " & Mid(WhereClause, 6) & vbCrLf
                    Else
                        s = s & " WHERE (isnull(IsChangeRequest,0) = 1 or isnull(IsChange,0) = 0) and isnull(billingitemid,0)=0 and " & Mid(WhereClause, 6) & vbCrLf
                    End If
                Else
                    s = s & " WHERE ischange=0 and isnull(billingitemid,0)=0 and " & Mid(WhereClause, 6) & vbCrLf
                End If
                
            End If
            
            
            
            Select Case KeyFld
                Case ""
                Case "EstAssemblyID"
                    s = s & "GROUP BY " & KeyFld & "," & DisplayFld & ",seq" & vbCrLf
                Case "ChangeOrder"
                    s = s & "GROUP BY " & KeyFld & vbCrLf
                Case "POIndex"
                    s = s & "GROUP BY POIndex,POIndexDescription" & vbCrLf
                Case Else
                    s = s & "GROUP BY " & KeyFld & "," & DisplayFld & vbCrLf
            End Select
            
            
            Select Case SortFld
                Case "ChangeOrder", "CORSort"
                    s = s & "ORDER BY 1"
                Case "Seq"
                    s = s & "ORDER BY seq"
                Case Else
                    s = s & "ORDER BY 2" ' & SortFld
            End Select
            

            mViewQuery = s
 
            Set rs = HFApp.SqlExec(s)
            
            
            'remove any children this node already has
            If .Row > 0 Then
                Set n = .GetNode().GetNode(flexNTFirstChild)
                While Not n Is Nothing
                    Call n.RemoveNode
                    Set n = .GetNode().GetNode(flexNTFirstChild)
                Wend
            End If
            
            'now add new children
            While Not rs.EOF
                KeyValue = "" & rs(0)
                DisplayValue = Trim("" & rs(1))
                If DisplayValue = "" Then DisplayValue = " -- "

                If Level = -1 Then
                    r = .Rows
                    Call .AddItem(DisplayValue, r)
                    
                    .Cell(flexcpText, r, .ColIndex("Budgeted")) = format(Val("" & rs("budgeted")), "#,###.00")
                    .Cell(flexcpText, r, .ColIndex("Changes")) = format(Val("" & rs("Changes")), "#,###.00")
                    .Cell(flexcpText, r, .ColIndex("Commited")) = format(Val("" & rs("commited")), "#,###.00")
                    
                    .Cell(flexcpText, r, .ColIndex("IsChangeRequest")) = IIf(mViewIndex = 0, Val("" & rs("IsChangeRequest")), 0)
                    .Cell(flexcpText, r, .ColIndex("BudgetsLocked")) = Val("" & rs("BudgetsLocked"))
                    .Cell(flexcpText, r, .ColIndex("hflocation")) = Val("" & rs("hflocation"))
                    If Val("" & rs("BudgetsLocked")) <> 0 Then .Cell(flexcpForeColor, r, .ColIndex("Budgeted")) = vbGrayText
                    

                    .Cell(flexcpData, r, 0) = DisplayFld
                    .Cell(flexcpText, r, 1) = KeyValue
                    .Cell(flexcpData, r, 1) = KeyFld

                    Select Case True
                        Case Val("" & rs(2)) = Val("" & rs(3)): .Cell(flexcpText, r, .ColIndex("completed")) = msAll
                        Case Val("" & rs(3)) > 0:               .Cell(flexcpText, r, .ColIndex("completed")) = msSome
                        Case Else:                              .Cell(flexcpText, r, .ColIndex("completed")) = msNone
                    End Select

                    .Cell(flexcpText, r, 3) = msNone
                    Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(.Cell(flexcpValue, r, .ColIndex("completed")), .Cell(flexcpValue, r, .ColIndex("selected")), .Cell(flexcpText, r, .ColIndex("IsChangeRequest")))

                    .IsSubtotal(r) = True
                    .RowOutlineLevel(r) = Level + 1

                    If Level + 2 <> levels Then
                        Call .GetNode(r).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                    End If

                    Set n = .GetNode(r)
                    n.Expanded = False


                    WhereClause = " AND Job_No=" & DbQuote(str, mJob)

                    .RowData(r) = WhereClause & " AND " & KeyFld & " = " & DbQuote(str, KeyValue)

                Else
                    r = .Row
                    Set n = .GetNode(r).AddNode(flexNTLastChild, DisplayValue)

                    .Cell(flexcpText, n.Row, .ColIndex("Budgeted")) = format(Val("" & rs("budgeted")), "#,###.00")
                    .Cell(flexcpText, n.Row, .ColIndex("Changes")) = format(Val("" & rs("Changes")), "#,###.00")
                    .Cell(flexcpText, n.Row, .ColIndex("Commited")) = format(Val("" & rs("commited")), "#,###.00")
                    
                    .Cell(flexcpText, n.Row, .ColIndex("IsChangeRequest")) = IIf(mViewIndex = 0, Val("" & rs("IsChangeRequest")), 0)
                    .Cell(flexcpText, n.Row, .ColIndex("BudgetsLocked")) = Val("" & rs("BudgetsLocked"))
                    If Val("" & rs("BudgetsLocked")) <> 0 Then .Cell(flexcpForeColor, n.Row, .ColIndex("Budgeted")) = vbGrayText
                    .Cell(flexcpText, n.Row, .ColIndex("hflocation")) = Val("" & rs("hflocation"))
                    .Cell(flexcpData, n.Row, 0) = DisplayFld
                    .Cell(flexcpText, n.Row, 1) = KeyValue
                    .Cell(flexcpData, n.Row, 1) = KeyFld
                    Select Case True
                        Case rs(2) = rs(3): .Cell(flexcpText, n.Row, .ColIndex("completed")) = msAll
                        Case rs(3) > 0:     .Cell(flexcpText, n.Row, .ColIndex("completed")) = msSome
                        Case Else:          .Cell(flexcpText, n.Row, .ColIndex("completed")) = msNone
                    End Select
                    .Cell(flexcpText, n.Row, .ColIndex("selected")) = .Cell(flexcpText, .Row, .ColIndex("selected"))
                    Set .Cell(flexcpPicture, n.Row, 0) = MultiStateIcon(.Cell(flexcpValue, n.Row, .ColIndex("completed")), .Cell(flexcpValue, n.Row, .ColIndex("selected")), .Cell(flexcpText, n.Row, .ColIndex("IsChangeRequest")))


                    If Level + 2 < levels Then
                        Call .GetNode(n.Row).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                    End If

                    n.Expanded = False


                    Select Case KeyFld
                        Case "EstAssemblyID", "EstimateIndex", "POGroup"
                            .RowData(n.Row) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Num, KeyValue)
                        Case Else
                            .RowData(n.Row) = WhereClause & " AND " & KeyFld & " = " & DbQuote(str, KeyValue)
                    End Select

                End If

                rs.MoveNext
            Wend
        End If

On Error Resume Next
.Cell(flexcpPicture, 0, 0) = FMain.SmallIcons.ListImages("combo1").Picture

        Call .AutoSize(0, .cols - 1)
        .Redraw = flexRDBuffered
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "LoadAssemblies", s)
End Sub



Private Sub gAssemblies_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Row < 0 Then Exit Sub
    
    With gAssemblies
        If .Cell(flexcpText, Row, 3) = msSome Then .Cell(flexcpText, Row, 3) = msNone
        Set .Cell(flexcpPicture, Row, 0) = MultiStateIcon(msNone, .Cell(flexcpValue, Row, 3), .Cell(flexcpValue, Row, .ColIndex("IsChangeRequest")))
    End With
    Call ResetChildrenChecked(Row)
    Call ResetParentChecked(Row)
    Call LoadItems
    
    
End Sub

Private Sub gAssemblies_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
    Dim r As Long
    Dim levels As Long
    Dim Level As Long
    
    With gAssemblies
        If State = flexOutlineExpanded Then
            .Row = Row
            Call LoadAssemblies(False)
        End If
    End With
End Sub

Private Sub gAssemblies_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub gAssemblies_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error Resume Next
    gAssemblies.Row = gAssemblies.MouseRow
End Sub

Private Sub gAssemblies_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gAssemblies
        If .Row < 0 Then Exit Sub
        Select Case True
        
            Case KeyCode = vbKeyF And Shift = vbCtrlMask
                Call FFind.ShowForm(gAssemblies)
            
            Case KeyCode = vbKeySpace
                'toggle selected value
                i = .Row
                .Cell(flexcpText, i, 3) = IIf(.Cell(flexcpValue, i, 3) = msAll, msNone, msAll)
                Set .Cell(flexcpPicture, i, 0) = MultiStateIcon(.Cell(flexcpValue, i, 2), .Cell(flexcpValue, i, 3), .Cell(flexcpValue, i, .ColIndex("IsChangeRequest")))
                Call gAssemblies_AfterEdit(i, 0)
            
            Case KeyCode = vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
                
            Case KeyCode = vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If
                
        End Select
    End With
End Sub

Private Sub gAssemblies_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim r As Long
    Dim ParentMenu As Long
    
    With gAssemblies
        .SetFocus
        r = .MouseRow
        If r < 0 Then Exit Sub
        If r = 0 Then
            ParentMenu = FMain.PopMenu.MenuIndex("mnuEstimateItemViews")
            Call FMain.PopMenu.ClearSubMenusOfItem(ParentMenu)
            For i = 0 To UBound(mViews)
                FMain.PopMenu.AddItem mViews(i).Name, "EstimateItemViews" & i, , i, ParentMenu, , i = mViewIndex
            Next
            PopupMenu FMain.mnuEstimateItemViews, , .Left, .Top + .RowHeight(0)
        Else
            'they clicked on the checkbox
            .Cell(flexcpText, r, 3) = IIf(.Cell(flexcpValue, r, 3) = msAll, msNone, msAll)
            Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(msNone, .Cell(flexcpValue, r, 3), False)
            Call gAssemblies_AfterEdit(r, 0)
        End If
    End With
    
End Sub

Private Sub gAssemblies_RowColChange()
On Error Resume Next
    With gAssemblies
        .Col = 0
        Call .ShowCell(.Row, 0)
    End With
End Sub


Private Property Get MultiStateIcon(Completed As MultiStateEnum, Selected As MultiStateEnum, IsChangeRequest As Boolean) As IPictureDisp
    Set MultiStateIcon = FMain.MultiStateIcons.ListImages("K0" & Selected).Picture
End Property

Private Sub ResetChildrenChecked(ByVal ParentRow As Long)
    Dim r As Long
    If ParentRow < 0 Then Exit Sub
    With gAssemblies
        'for all my children set checkbox to same as me
        r = .GetNodeRow(ParentRow, flexNTFirstChild)
        While r <> -1
            .Cell(flexcpText, r, 3) = .Cell(flexcpText, ParentRow, 3)
            Set .Cell(flexcpPicture, r, 0) = MultiStateIcon(.Cell(flexcpValue, r, 2), .Cell(flexcpValue, r, 3), .Cell(flexcpValue, r, .ColIndex("IsChangeRequest")))
            Call ResetChildrenChecked(r)
            r = .GetNodeRow(r, flexNTNextSibling)
        Wend
    End With
End Sub

Private Sub ResetParentChecked(ByVal ChildRow As Long)
    Dim ParentRow As Long
    Dim r As Long
    With gAssemblies
        If ChildRow = -1 Then Exit Sub
        ParentRow = .GetNodeRow(ChildRow, flexNTParent)
        If ParentRow = -1 Then Exit Sub
        
        .Cell(flexcpText, ParentRow, 3) = .Cell(flexcpText, ChildRow, 3)
        Set .Cell(flexcpPicture, ParentRow, 0) = MultiStateIcon(.Cell(flexcpValue, ParentRow, 2), .Cell(flexcpValue, ParentRow, 3), .Cell(flexcpValue, ParentRow, .ColIndex("IsChangeRequest")))
        
        r = .GetNodeRow(ChildRow, flexNTFirstSibling)
        While r <> -1
            If .Cell(flexcpText, r, 3) <> .Cell(flexcpText, ChildRow, 3) Then
                .Cell(flexcpText, ParentRow, 3) = msSome
                Set .Cell(flexcpPicture, ParentRow, 0) = MultiStateIcon(.Cell(flexcpValue, ParentRow, 2), .Cell(flexcpValue, ParentRow, 3), .Cell(flexcpValue, ParentRow, .ColIndex("IsChangeRequest")))
                r = -1
            Else
                r = .GetNodeRow(r, flexNTNextSibling)
            End If
        Wend
        Call ResetParentChecked(ParentRow)
    End With
End Sub

Public Sub mnuEstimateItemViewsSub_Click(Index As Integer)
    
    mViewIndex = Index - 1
    Call LoadAssemblies(True)
    Call LoadItems
    
End Sub


Private Function WhereClause(Optional prefix As String)
    Dim i As Long
    Dim s As String
    
    If prefix <> "" Then prefix = prefix & "."
    
    With gAssemblies
        i = 1
        While i < .Rows And i >= 0
            If .Cell(flexcpText, i, 3) = msAll And .RowData(i) <> "" Then
                s = s & " OR (" & Mid(Replace(.RowData(i), " AND ", " AND " & prefix), 6) & ")"
                i = .GetNodeRow(i, flexNTNextSibling)
            Else
                i = i + 1
            End If
        Wend
    End With
    
    
    
    If s = "" Then
        WhereClause = "1=2"
    Else
        WhereClause = "(" & Mid(s, 5) & ")"
    End If


    
End Function






Private Sub SaveData()
    
    Dim s As String
    Dim i As Long
    Dim ID As Long
    Dim rs As Recordset
    Dim co  As String
    
    
    With gBillingItems
        
        If mChangeOrder Then
            s = ""
            s = s & "update tblcustomers set last_co=isnull(cast(last_co as integer),0)+1 where job_no=" & DbQuote(str, mJob) & vbCrLf
            s = s & "insert into changeordermaster(customer_no,change_order_no,purchasingco,usealtjccodes,Comments) select customer_no,'CO '+last_co,1,1," & DbQuote(str, Me.txtDescription) & " from tblcustomers where job_no=" & DbQuote(str, mJob) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
            co = "CO " & HFApp.SqlExec("select last_co from tblcustomers where job_no=" & DbQuote(str, mJob), dbHomefront)(0)
        End If
        
        For i = 1 To .Rows - 1
            
            s = ""
            s = s & "select billingitemid" & vbCrLf
            s = s & "from billingitems" & vbCrLf
            s = s & "where job=" & DbQuote(str, mJob) & vbCrLf
            s = s & "  and description=" & DbQuote(str, .TextMatrix(i, .ColIndex("description"))) & vbCrLf
            Set rs = HFApp.SqlExec(s)
            If Not rs.EOF Then
                ID = rs(0)
            Else
                s = ""
                s = s & "insert into billingitems(job,description,price,amtbilled,sortorder,taxgroup,revenueaccount,holdbackrate)" & vbCrLf
                s = s & "values(" & DbQuote(str, mJob) & vbCrLf
                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("description"))) & vbCrLf
                s = s & "      ,0,0,99999" & vbCrLf
                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("taxgroup"))) & vbCrLf
                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("revenueaccount"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("holdbackrate"))) & ")" & vbCrLf
                s = s & "update billingitems set sortorder=billingitemid where billingitemid=@@identity" & vbCrLf
                Call HFApp.SqlExec(s)
                ID = HFApp.SqlIdentity("billingitems")
            End If
                        
            s = ""
            s = s & "update billingitems" & vbCrLf
            s = s & "   set Price=Price+" & DbQuote(Num, .TextMatrix(i, .ColIndex("price"))) & vbCrLf
            s = s & "where billingitemid=" & DbQuote(Num, ID) & vbCrLf
            Call HFApp.SqlExec(s)
            
            s = ""
            s = s & "update estimateitems" & vbCrLf
            s = s & "set billingitemid=" & DbQuote(Num, ID) & vbCrLf
            s = s & "where estitemid in(select estitemid from estimateditems where billingitemid=0 and " & .TextMatrix(i, .ColIndex("whereclause")) & ")"
            Call HFApp.SqlExec(s)
            
            If mChangeOrder Then
                s = ""
                s = s & "update estimateassemblies" & vbCrLf
                s = s & "set changerequeststatus='', changeorder=" & DbQuote(str, co) & vbCrLf
                s = s & "where estassemblyid in(select distinct estassemblyid from estimateditems where " & .TextMatrix(i, .ColIndex("whereclause")) & ")"
                Call HFApp.SqlExec(s)
            End If
            
        Next
    End With
    
    mCancel = False

End Sub

Private Sub gGrouping_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call LoadItems

End Sub

Private Sub optAddon_Click(Index As Integer)
    Call LoadItems
End Sub


Private Sub LoadFields()
    Dim i As Long
    Dim rs As Recordset
    
    Set rs = HFApp.SqlExec("select * from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no = " & DbQuote(str, mJob))
    With gGrouping
        .Rows = 6
        For i = 1 To 40
            If "" & rs.fields("WBSdesc" & format(i, "00")) <> "" Then
                .AddItem rs.fields("WBSdesc" & format(i, "00")) & vbTab & "wbs" & format(i, "00")
            End If
        Next
        
        For i = 0 To .Rows - 1
            .Cell(flexcpChecked, i, 0) = IIf(IniGet(AppIni, "Options", "BillingLevel" & i, 0) = "1", flexChecked, flexUnchecked)
        Next
        
    End With
    
End Sub
