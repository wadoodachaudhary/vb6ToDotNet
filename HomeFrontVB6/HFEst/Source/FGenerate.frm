VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FGenerate 
   Caption         =   " "
   ClientHeight    =   5040
   ClientLeft      =   9900
   ClientTop       =   1875
   ClientWidth     =   11490
   Icon            =   "FGenerate.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   11490
   Begin VB.CommandButton cmdNav 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   6060
      Picture         =   "FGenerate.frx":000C
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   1
      Left            =   4800
      Picture         =   "FGenerate.frx":0596
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   4560
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3465
      Left            =   210
      TabIndex        =   0
      Top             =   990
      Width           =   11115
      _cx             =   19606
      _cy             =   6112
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FGenerate.frx":0B20
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
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   630
      Picture         =   "FGenerate.frx":0C72
      Top             =   390
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   300
      Picture         =   "FGenerate.frx":11FC
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"FGenerate.frx":1AC6
      Height          =   615
      Left            =   1110
      TabIndex        =   1
      Top             =   270
      Width           =   5115
   End
End
Attribute VB_Name = "FGenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Cancel As Boolean


Public Function VerifyBudgets(WhereClause As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    
    s = ""
    s = s & "SELECT JCCostCodeDesc CostCode,JCCategoryDesc Category" & vbCrLf
    s = s & "      ,convert(varchar(50),cast(SUM(isnull(BudgetPretax,0)) as money),1) Budgeted" & vbCrLf
    s = s & "      ,convert(varchar(50),cast(SUM(isnull(POPretax,0)) as money),1) Committed" & vbCrLf
    s = s & "      ,convert(varchar(50),cast(SUM(isnull(POPretax,0))-SUM(isnull(BudgetPretax,0)) as money),1) OverBudget" & vbCrLf
    s = s & "      ,case sum(isnull(budgetpretax,0)) when 0 then 100 else round((SUM(POPretax)-SUM(BudgetPretax))/SUM(BudgetPretax)*100,2) end Percentage" & vbCrLf
    s = s & "FROM EstimatedItems i" & vbCrLf
    s = s & "WHERE PODeleted = 0 And POGenBatch = 0 And " & WhereClause
    s = s & "GROUP BY JCCostCode,JCCostCodeDesc,JCCategory,JCCategoryDesc" & vbCrLf
    s = s & "HAVING SUM(isnull(POPretax,0))>SUM(isnull(BudgetPretax,0))" & vbCrLf
    s = s & "ORDER BY 1,2" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        VerifyBudgets = True
    Else
        Load Me
        Me.Caption = "Warning"
        Me.Label1.Caption = "The PO's you are about to generate will exceed the jobs budgeted amounts. Are you sure you want to continue?"
        Me.cmdNav(0).Visible = True
        Call LoadData(rs)
        Beep
        Cancel = False
        Me.Show vbModal
        VerifyBudgets = Not Cancel
    End If

End Function

Public Function ValidateData(WhereClause As String, Budgets As Boolean) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select " & IIf(Budgets, "Budget", "PO") & "vendor Vendor" & vbCrLf
    s = s & "      ,POIndex" & vbCrLf
    s = s & "      ,EstPhase" & vbCrLf
    s = s & "      ,EstItem" & vbCrLf
    s = s & "      ,ItemDesc" & vbCrLf
    s = s & "      ,JCCostCode" & vbCrLf
    s = s & "      ,JCCategory" & vbCrLf
    s = s & "      ," & IIf(Budgets, "Budget", "PO") & "TaxGroup TaxGroup" & vbCrLf
    s = s & "      ,substring(case when jccostcodedesc='' then ', Cost code is missing or invalid' else '' end +" & vbCrLf
    s = s & "       case when jccategorydesc='' then ', Category is missing or invalid' else '' end +" & vbCrLf
    s = s & "       case when poindex='' then ', PO index is missing or invalid' else '' end +" & vbCrLf
    If Not Budgets Then
        s = s & "       case when variancejccategorydesc ='' and variancepretax<>0 and BudgetsLocked = 1 then ', Variance category is missing or invalid' else '' end +" & vbCrLf
    End If
    s = s & "       case when " & IIf(Budgets, "Budget", "PO") & "vendorname='' then ', Vendor is missing or invalid' else '' end +" & vbCrLf
    s = s & "       case when " & IIf(HFApp.Options(TaxGroupRequired), "1=1", "1=0") & " and " & IIf(Budgets, "Budget", "PO") & "taxgroup='' then ', Tax group is missing or invalid' else '' end,3,999) Problem" & vbCrLf
    s = s & " from estimateditems i" & vbCrLf
    s = s & " where (JCcostcodeDesc ='' or " & vbCrLf
    s = s & "        jccategorydesc ='' or " & vbCrLf
    s = s & "        jccategorydesc ='' or " & vbCrLf
    
    If Not Budgets Then
        s = s & "        (variancejccategorydesc ='' and variancepretax<>0  and BudgetsLocked = 1) or " & vbCrLf
    End If
    s = s & "        poindex ='' or " & vbCrLf
    s = s & "        " & IIf(Budgets, "Budget", "PO") & "vendorname='' or " & vbCrLf
    s = s & "        (" & IIf(HFApp.Options(TaxGroupRequired), "1=1", "1=0") & " and " & IIf(Budgets, "Budget", "PO") & "taxgroup=''))" & vbCrLf
    s = s & "   and " & IIf(Budgets, "Budget", "PO") & "deleted=0" & vbCrLf
    If Not Budgets Then
        s = s & "and pogenbatch = 0" & vbCrLf
    End If
    s = s & "   and " & WhereClause
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        ValidateData = True
    Else
        Load Me
        ValidateData = False
        Call LoadData(rs)
        Beep
        Me.Show vbModal
    End If
End Function

Private Sub LoadData(rs As Recordset)
    Dim r As Long
    Dim c As Long
    
    With gData
        .Rows = 1
        .Cols = rs.fields.Count
        
        r = 0
        For c = 0 To rs.fields.Count - 1
            .TextMatrix(r, c) = rs.fields(c).Name
        Next
        
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            For c = 0 To rs.fields.Count - 1
                .TextMatrix(r, c) = "" & rs(c).value
            Next
            rs.MoveNext
            
        Wend
        Call .AutoSize(0, .Cols - 1)
    End With
End Sub


Private Sub cmdNav_Click(Index As Integer)
    Cancel = Index = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    gData.Move margin, gData.Top, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - gData.Top - 2 * margin - cmdNav(1).Height
    
    If cmdNav(0).Visible Then
        cmdNav(0).Cancel = True
        cmdNav(0).Default = True
        cmdNav(1).Move Me.ScaleWidth / 2 - margin / 2 - cmdNav(1).Width, Me.ScaleHeight - margin - cmdNav(1).Height
        cmdNav(0).Move Me.ScaleWidth / 2 + margin / 2, Me.ScaleHeight - margin - cmdNav(1).Height
    Else
        cmdNav(1).Move Me.ScaleWidth / 2 - cmdNav(1).Width / 2, Me.ScaleHeight - margin - cmdNav(1).Height
        cmdNav(1).Cancel = True
        cmdNav(1).Default = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Dim r As Long
    If Button = vbRightButton Then
        If gData.MouseRow = 0 Then
            Cancel = True
            Call FMain.ShowColumnMenu(gData)
        End If
    End If
End Sub
