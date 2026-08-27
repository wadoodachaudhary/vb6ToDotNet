VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FItemChart 
   Caption         =   "Item Chart"
   ClientHeight    =   5475
   ClientLeft      =   2190
   ClientTop       =   2355
   ClientWidth     =   10215
   Icon            =   "FItemChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10215
   Begin VB.TextBox txtFormula 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1230
      TabIndex        =   10
      Top             =   1200
      Width           =   3225
   End
   Begin VB.TextBox txtName 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1230
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   3225
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1058
      ButtonWidth     =   1376
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Delete  "
            Key             =   "Delete"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3465
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   10005
      _cx             =   17648
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"FItemChart.frx":000C
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
   Begin VB.ComboBox cboVariable 
      Height          =   315
      Index           =   1
      Left            =   5730
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2865
   End
   Begin VB.ComboBox cboVariable 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5730
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1215
      Width           =   2865
   End
   Begin VB.ComboBox cboVariable 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   5730
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1470
      Width           =   2865
   End
   Begin VB.Image cmdFormula 
      Height          =   240
      Left            =   4470
      Picture         =   "FItemChart.frx":00D4
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Formula"
      Height          =   195
      Left            =   555
      TabIndex        =   11
      Top             =   1230
      Width           =   555
   End
   Begin VB.Image cmdVariable 
      Enabled         =   0   'False
      Height          =   240
      Index           =   2
      Left            =   8610
      Picture         =   "FItemChart.frx":021E
      Top             =   1215
      Width           =   240
   End
   Begin VB.Image cmdVariable 
      Height          =   240
      Index           =   1
      Left            =   8610
      Picture         =   "FItemChart.frx":0368
      Top             =   960
      Width           =   240
   End
   Begin VB.Image cmdVariable 
      Enabled         =   0   'False
      Height          =   240
      Index           =   3
      Left            =   8610
      Picture         =   "FItemChart.frx":04B2
      Top             =   1470
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Variable 2"
      Height          =   195
      Index           =   2
      Left            =   4920
      TabIndex        =   9
      Top             =   1260
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Variable 1"
      Height          =   195
      Index           =   0
      Left            =   4920
      TabIndex        =   8
      Top             =   1005
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Variable 3"
      Height          =   195
      Index           =   1
      Left            =   4920
      TabIndex        =   7
      Top             =   1515
      Width           =   705
   End
   Begin VB.Label Label112 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Chart Name"
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   990
      Width           =   840
   End
End
Attribute VB_Name = "FItemChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public EventTraps As Collection
Const SRCFILE = "FItemChart::"

Private mChart    As String
Private mDirty    As Boolean

Private MinValue(3) As Double
Private MaxValue(3) As Double




Private Property Get Dirty() As Boolean
    Dirty = mDirty
End Property
Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
    Toolbar.Buttons("Save").Enabled = mDirty
    Toolbar.Buttons("Delete").Enabled = mChart <> ""
End Property

Private Sub cboVariable_Click(Index As Integer)
    Dirty = True

    Select Case Index
        Case 1
            cboVariable(2).Enabled = cboVariable(1).Text <> ""
            cboVariable(3).Enabled = cboVariable(2).Text <> ""
        Case 2
            cboVariable(3).Enabled = cboVariable(2).Text <> ""
    End Select

    cmdVariable(1).Enabled = cboVariable(1).Enabled
    cmdVariable(2).Enabled = cboVariable(2).Enabled
    cmdVariable(3).Enabled = cboVariable(3).Enabled

    Call ConfigDimensions

End Sub

Private Sub cboVariable_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cboVariable(Index).ListIndex = -1
        Call ConfigDimensions
        Dirty = True
    End If
    
End Sub

Private Sub cmdFormula_Click()
    Dim s As String
    s = Me.txtFormula.Text
    If FFormulaEditor.EditFormula("", s) Then
        Me.txtFormula.Text = s
        Dirty = True
    End If
End Sub

Private Sub cmdVariable_Click(Index As Integer)
    If FVariable.EditVariable(cboVariable(Index).Text) Then LoadVariables
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbCtrlMask And KeyCode = vbKeyO:  Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case Shift = vbCtrlMask And KeyCode = vbKeyS:  Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub

Public Sub Edit(Chart As String)
    mChart = Chart
    Me.Show vbModal

End Sub

Private Sub Form_Load()
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    Call LoadVariables
    Call ConfigDimensions
    gData.Rows = 2
    
    If mChart <> "" Then LoadData
    
    Dirty = False
End Sub

Private Sub LoadVariables()
    Dim rs  As ADODB.Recordset
    Dim i1 As Long
    Dim i2 As Long
    Dim i3 As Long
    
    
    Set rs = HFApp.SqlExec("select name from variables order by 1")
    
    i1 = cboVariable(1).ListIndex
    i2 = cboVariable(2).ListIndex
    i3 = cboVariable(3).ListIndex
    
    cboVariable(1).Clear
    cboVariable(2).Clear
    cboVariable(3).Clear
    
    While Not rs.EOF
        cboVariable(1).AddItem Trim("" & rs(0))
        cboVariable(2).AddItem Trim("" & rs(0))
        cboVariable(3).AddItem Trim("" & rs(0))
        rs.MoveNext
    Wend

On Error Resume Next
    cboVariable(1).ListIndex = i1
    cboVariable(2).ListIndex = i2
    cboVariable(3).ListIndex = i3
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    gData.Move margin, gData.Top, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - gData.Top - margin
End Sub



Private Sub gData_AfterSort(ByVal Col As Long, Order As Integer)
    Call gData.AddItem("")
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        Select Case .ColKey(Col)
            Case "Phase", "Item", "Description"
                .ComboList = "..."
            Case Else
                .ComboList = ""
        End Select
        Cancel = False
    End With
End Sub

Private Sub gData_BeforeSort(ByVal Col As Long, Order As Integer)
    Call gData.RemoveItem(gData.Rows - 1)
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim i As Long
    With gData
        Select Case .ColKey(Col)
            Case "Phase", "Item", "Description"
                s = "select Phase,Item,Description,POIndex from tblphaseitem where DivisionID = " & HFApp.DivisionID
                If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s, , , , , , True) Then Exit Sub
                
                For i = 1 To FPickList.SelectedItems
                
                    If i <> 1 Then
                        Row = .Rows - 1
                        .AddItem ""
                    End If
                
                    .TextMatrix(Row, .ColIndex("phase")) = FPickList.SelectedItem("phase", i)
                    .TextMatrix(Row, .ColIndex("item")) = FPickList.SelectedItem("item", i)
                    .TextMatrix(Row, .ColIndex("description")) = FPickList.SelectedItem("description", i)
                Next
                
                Dirty = True
                
        End Select
    End With
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyDelete And Shift <> 0 And gData.Row > 0 And gData.Row < gData.Rows - 1
            gData.RemoveItem
            gData.Row = gData.Row - 1
            Dirty = True
        
        Case KeyCode = vbKeyO And Shift <> 0:    Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case KeyCode = vbKeyS And Shift <> 0:    Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub

Private Sub gData_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        If Row = .Rows - 1 Then
            .AddItem ""
        End If
    End With
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    With gData
        Select Case .ColKey(Col)
            Case "D1", "D2", "D3"
                If .ColComboList(Col) = "" Then
                    Cancel = Not IsNumeric(.EditText)
                    i = Val(Right(.ColKey(Col), 1))
                    If Val(.EditText) < MinValue(i) Then
                        MsgBox "The minimum value for " & cboVariable(i).Text & " is '" & format(MinValue(i), "0") & "'. The number you have entered is too small.", vbInformation, App.ProductName
                        Cancel = True
                    End If
                    If Val(.EditText) > MaxValue(i) Then
                        MsgBox "The maximum value for " & cboVariable(i).Text & " is '" & format(MaxValue(i), "0") & "'. The number you have entered is too large.", vbInformation, App.ProductName
                        Cancel = True
                    End If
                End If
                
                
                                
                
                
        End Select
    End With
    If Not Cancel Then Dirty = True
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim f As Form
    Dim s As String

    Select Case Button.Key
        Case "Delete"
            If mChart <> "" Or txtName.Text <> "" Then
                If MsgBox("Are you sure you want to delete this chart?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then Exit Sub
                Call DeleteChart
            End If
            
            
        Case "New"
            If Not SaveData(True) Then Exit Sub
            mChart = ""
            Call LoadData
        
        Case "Open"
            If Not SaveData(True) Then Exit Sub

            s = "SELECT Name from itemcharts"
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Item Chart", s, mChart) Then
                mChart = FPickList.SelectedItem("Name")
                Call LoadData
            End If


         Case "Save"
            Call SaveData(False)


    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh

    Dim s As String
    Dim r As Long
    Dim dupString As String

    If Not Dirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox(Me.Caption & " has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If



    'validate
    txtName.Text = Trim(txtName.Text)
    If Trim(txtName.Text) = "" Then
        MsgBox "Chart name is required.", vbExclamation, App.ProductName
        Call SetCtrlFocus(txtName)
        Exit Function
    End If
    If Trim(cboVariable(1).Text) = "" Then
        MsgBox "At least one variable must be specified.", vbExclamation, App.ProductName
        Call SetCtrlFocus(cboVariable(1))
        Exit Function
    End If
    With gData
        For r = 1 To .Rows - 2
            If .TextMatrix(r, .ColIndex("Phase")) = "" Or .TextMatrix(r, .ColIndex("Item")) = "" Then
                MsgBox "You must specify a phase/item on each line.", vbExclamation, App.ProductName
                Exit Function
            End If
        Next
    
    End With




    Screen.MousePointer = vbHourglass
    dupString = "Unable to save this chart. " & vbQuote & txtName.Text & vbQuote & " has already been used." & vbCrLf & vbCrLf & "Please enter a different name."
    If mChart = "" Then
        s = ""
        s = s & "insert into itemcharts(name,formula,variable1,variable2,variable3)" & vbCrLf
        s = s & "values(" & DbQuote(Str, txtName.Text) & "," & DbQuote(Str, txtFormula.Text) & "," & DbQuote(Str, cboVariable(1).Text) & "," & DbQuote(Str, cboVariable(2).Text) & "," & DbQuote(Str, cboVariable(3).Text) & ")"
        Call HFApp.SqlExec(s)
        mChart = txtName.Text
    Else
        s = ""
        s = s & "update itemcharts" & vbCrLf
        s = s & "set name=" & DbQuote(Str, txtName.Text) & vbCrLf
        s = s & "   ,formula=" & DbQuote(Str, txtFormula.Text) & vbCrLf
        s = s & "   ,variable1=" & DbQuote(Str, cboVariable(1).Text, , True) & vbCrLf
        s = s & "   ,variable2=" & DbQuote(Str, cboVariable(2).Text, , True) & vbCrLf
        s = s & "   ,variable3=" & DbQuote(Str, cboVariable(3).Text, , True) & vbCrLf
        s = s & "where name=" & DbQuote(Str, mChart) & vbCrLf
        Call HFApp.SqlExec(s)
        mChart = txtName.Text
    End If

    
    Call HFApp.SqlExec("delete from itemchartdetails where name=" & DbQuote(Str, txtName.Text))
    With gData
    For r = 1 To .Rows - 2
        
        dupString = "Unable to save row " & r & ". The variable values have already been used."
        
        s = ""
        s = s & "insert into itemchartdetails(name,dimension1,dimension2,dimension3,phase,item)" & vbCrLf
        s = s & "values(" & DbQuote(Str, mChart) & vbCrLf
        s = s & "      ," & DbQuote(Str, IIf(.ColHidden(.ColIndex("D1")), "", .TextMatrix(r, .ColIndex("D1")))) & vbCrLf
        s = s & "      ," & DbQuote(Str, IIf(.ColHidden(.ColIndex("D2")), "", .TextMatrix(r, .ColIndex("D2")))) & vbCrLf
        s = s & "      ," & DbQuote(Str, IIf(.ColHidden(.ColIndex("D3")), "", .TextMatrix(r, .ColIndex("D3")))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item"))) & ")"
        Call HFApp.SqlExec(s, dbHomefront)
    Next
    End With



    SaveData = True
    Dirty = False
    Screen.MousePointer = vbDefault

Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Screen.MousePointer = vbDefault
        MsgBox dupString, vbExclamation, App.ProductName
    Else
        Call errHandler(SRCFILE & "SaveData", s)
    End If
End Function


Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset

    s = ""
    s = s & "select *" & vbCrLf
    s = s & "  from itemcharts" & vbCrLf
    s = s & " where name=" & DbQuote(Str, mChart)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        txtName.Text = ""
        txtFormula.Text = ""
        cboVariable(1).ListIndex = -1
        cboVariable(2).ListIndex = -1
        cboVariable(3).ListIndex = -1
    Else
        txtName.Text = "" & rs("Name")
        txtFormula.Text = "" & rs("Formula")
        cboVariable(1).ListIndex = -1
        cboVariable(2).ListIndex = -1
        cboVariable(3).ListIndex = -1
        Call SetListIndex(cboVariable(1), , "" & rs("Variable1"))
        Call SetListIndex(cboVariable(2), , "" & rs("Variable2"))
        Call SetListIndex(cboVariable(3), , "" & rs("Variable3"))
    End If
    
    
    
    With gData
        .Rows = 1
    
        s = ""
        s = s & "select i.*,d.description" & vbCrLf
        s = s & "  from itemchartdetails i" & vbCrLf
        s = s & "  join tblphaseitem d on(d.DivisionID = " & HFApp.DivisionID & " and i.phase=d.phase and i.item=d.item)" & vbCrLf
        s = s & " where name=" & DbQuote(Str, mChart)
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, .ColIndex("D1")) = "" & rs("Dimension1")
            .TextMatrix(.Rows - 1, .ColIndex("D2")) = "" & rs("Dimension2")
            .TextMatrix(.Rows - 1, .ColIndex("D3")) = "" & rs("Dimension3")
            .TextMatrix(.Rows - 1, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(.Rows - 1, .ColIndex("Item")) = "" & rs("Item")
            .TextMatrix(.Rows - 1, .ColIndex("Description")) = "" & rs("Description")
            rs.MoveNext
        Wend
    
        .AddItem ""
    End With
    
    
    

    
    Dirty = False
    
End Sub



Private Sub txtFormula_Change()
    Dirty = True
End Sub

Private Sub txtName_Change()
    Dirty = True
End Sub

Private Sub txtName_GotFocus()
    SelectAll txtName
End Sub

Private Sub ConfigDimensions()
    Dim i As Long
    Dim c As Long
    Dim rs As Recordset
    Dim s As String
    
    With gData
        For i = 1 To 3
            c = .ColIndex("D" & i)
            If cboVariable(i).Text = "" Then
                .ColHidden(c) = True
            Else
                s = "select * from variables where name=" & DbQuote(Str, cboVariable(i).Text)
                Set rs = HFApp.SqlExec(s)
                
                .ColHidden(c) = False
                .TextMatrix(0, c) = "" & rs("name")
                MinValue(i) = Val("" & rs("minimumvalue"))
                MaxValue(i) = Val("" & rs("maximumvalue"))
                
                
                
                s = "" & rs("listofvalues")
                If s = "" Then
                    .ColComboList(c) = ""
                Else
                    .ColComboList(c) = " |" & s
                End If
            End If
        Next
    End With
End Sub

Private Sub DeleteChart()
    
    If mChart <> "" Then
        Call HFApp.SqlExec("delete from itemcharts where name=" & DbQuote(Str, mChart))
        Call HFApp.SqlExec("delete from itemchartdetails where name=" & DbQuote(Str, mChart))
        Call HFApp.SqlExec("delete from tbldbassemblydetails where itemchart=" & DbQuote(Str, mChart))
        mChart = ""
    End If
    Call LoadData
End Sub
