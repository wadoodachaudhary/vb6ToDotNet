VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FColumns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Columns"
   ClientHeight    =   6240
   ClientLeft      =   1950
   ClientTop       =   3615
   ClientWidth     =   4815
   Icon            =   "FColumns.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdShow 
      Caption         =   "Restore Default Layout"
      Height          =   870
      Index           =   4
      Left            =   3600
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid gList 
      Height          =   4605
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   3255
      _cx             =   5741
      _cy             =   8123
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
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show &All"
      Height          =   315
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Hide All"
      Height          =   315
      Index           =   3
      Left            =   3600
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Hide"
      Height          =   315
      Index           =   2
      Left            =   3600
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Height          =   315
      Index           =   0
      Left            =   3600
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move &Down"
      Height          =   315
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move &Up"
      Height          =   315
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4005
      IntegralHeight  =   0   'False
      ItemData        =   "FColumns.frx":000C
      Left            =   240
      List            =   "FColumns.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   5700
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   5700
      Width           =   1095
   End
   Begin VB.PictureBox ULine1 
      Height          =   30
      Left            =   240
      ScaleHeight     =   30
      ScaleWidth      =   4335
      TabIndex        =   10
      Top             =   5580
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   $"FColumns.frx":0010
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FColumns::"
Public Grid As Object 'VSFlexGrid





Private Sub cmdMove_Click(Index As Integer)
On Error GoTo eh
    
    Dim r  As Long
    Dim NewRow  As Long
    
    Dim data    As String
    Dim Text    As String
    Dim checked As CellCheckedSettings
    
    r = gList.Row
    data = gList.Cell(flexcpData, r, 0)
    Text = gList.Cell(flexcpText, r, 0)
    checked = gList.Cell(flexcpChecked, r, 0)
    
    Select Case Index
        Case 0: NewRow = r - 1  'up
        Case 1: NewRow = r + 1  'down
    End Select
    gList.RemoveItem r
    
    NewRow = Max(NewRow, 0)
    NewRow = Min(NewRow, gList.Rows)
    
    
    gList.AddItem "", NewRow
    gList.Cell(flexcpText, NewRow, 0) = Text
    gList.Cell(flexcpData, NewRow, 0) = data
    gList.Cell(flexcpChecked, NewRow, 0) = checked
    
    Call gList.Select(NewRow, 0)
    Call gList.ShowCell(NewRow, 0)
    
    
Exit Sub
eh: Call errHandler(SRCFILE & "cmdMove_Click")
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    Dim i As Integer
    Select Case Index
        Case 0 'OK
            With Grid
                For i = 0 To gList.Rows - 1
                
                    .TextMatrix(0, .ColIndex(gList.Cell(flexcpData, i, 0))) = gList.Cell(flexcpText, i, 0)

                    .ColPosition(.ColIndex(gList.Cell(flexcpData, i, 0))) = i
                    .ColHidden(.ColIndex(gList.Cell(flexcpData, i, 0))) = gList.Cell(flexcpChecked, i, 0) = flexUnchecked
                Next
            End With
            Unload Me
        Case 1 'cancel
            Unload Me
    End Select
    Exit Sub
eh: Call errHandler(SRCFILE & "cmdNav_Click")
End Sub




Private Sub cmdShow_Click(Index As Integer)
On Error GoTo eh
    Select Case Index
        Case 0: gList.Cell(flexcpChecked, gList.Row, 0, gList.RowSel, 0) = flexChecked    'show current
        Case 1: gList.Cell(flexcpChecked, 0, 0, gList.Rows - 1, 0) = flexChecked             'show all
        Case 2: gList.Cell(flexcpChecked, gList.Row, 0, gList.RowSel, 0) = flexUnchecked  'hide current
        Case 3: gList.Cell(flexcpChecked, 0, 0, gList.Rows - 1, 0) = flexUnchecked             'hide all
        Case 4: Call RestoreDefaultLayout
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "cmdShow_Click")
End Sub

Private Sub RestoreDefaultLayout()
    Dim s As String
    Dim rs As Recordset
    Dim i As Long
    Dim r As Long
    
    If Grid.Tag = "" Then Exit Sub

    s = ""
    s = s & "select colkey,caption,hidden" & vbCrLf
    s = s & "from AppGridLayout " & vbCrLf
    s = s & "where uid = char(1)" & vbCrLf
    s = s & "and gridname=" & DbQuote(Str, Grid.Tag) & vbCrLf
    s = s & "order by colindex" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    If rs.EOF Then Exit Sub
    
    With gList
    While Not rs.EOF
        
        r = .FindRow("" & rs("colkey"))
        If r <> -1 Then
            'restore label and visibility
            .Cell(flexcpText, r, 0) = "" & rs("caption")
            .Cell(flexcpChecked, r, 0) = IIf("" & rs("hidden") = "True", flexUnchecked, flexChecked)
            'restore column order
            .RowPosition(r) = i
            
            i = i + 1
        End If
        rs.MoveNext
    Wend
    End With
    
End Sub

Private Sub Form_Load()
On Error GoTo eh
    Call IniGetForm(Me)
    LoadColumns
    Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub

Private Sub LoadColumns()
On Error GoTo eh
    Dim i As Integer
    gList.Rows = 0
    With Grid
        For i = .FixedCols To .Cols - 1
            If .ColKey(i) <> "" And .TextMatrix(0, i) <> "" Then
                gList.AddItem ""
                gList.RowData(gList.Rows - 1) = .ColKey(i) 'this is only used to enable findrow in RestoreDefaultLayout
                gList.Cell(flexcpText, gList.Rows - 1, 0) = .TextMatrix(0, i)
                gList.Cell(flexcpData, gList.Rows - 1, 0) = .ColKey(i)
                gList.Cell(flexcpChecked, gList.Rows - 1, 0) = IIf(.ColHidden(i), flexUnchecked, flexChecked)
            End If
        Next
    End With
    If gList.Rows > 0 Then gList.Row = 0
    Exit Sub
eh: Call errHandler(SRCFILE & "LoadColumns")
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error GoTo eh
    Call IniPutForm(Me)
    Exit Sub
eh: Call errHandler(SRCFILE & "Form_Unload")
End Sub

Private Sub gList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As Long

    With gList
        r = .Row
        .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbRed
        r = .DragRow(r)
        .Cell(flexcpCustomFormat, r, 0, r, .Cols - 1) = False
    End With

End Sub



