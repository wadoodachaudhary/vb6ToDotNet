VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FTableLocks 
   Caption         =   "Table Locks"
   ClientHeight    =   5520
   ClientLeft      =   15765
   ClientTop       =   11490
   ClientWidth     =   11505
   Icon            =   "FTableLocks.frx":0000
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FTableLocks.frx":000C
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
End
Attribute VB_Name = "FTableLocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FTableLocks::"
Private mDirty As Boolean


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyS And Shift = vbCtrlMask:       Call SaveData(False)
        Case KeyCode = vbKeyEscape:                         Unload Me
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo eh
    Call WindowOnTop(Me, True)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    Call LoadData
Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
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
        For r = .Rows - 1 To 1 Step -1
            If .RowData(r) = "DIRTY" And .Cell(flexcpChecked, r, .ColIndex("locked")) = flexUnchecked Then
            
                s = ""
                s = s & "UPDATE tablelog" & vbCrLf
                s = s & "SET recordlocked=0" & vbCrLf
                s = s & "WHERE tablename=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("Tablename"))) & vbCrLf
                s = s & "  AND keyrecord=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("key"))) & vbCrLf
                Call HFApp.SqlExec(s)
                
                Call .RemoveItem(r)
            End If
        Next
    End With
    
    mDirty = False
    SaveData = True
    Screen.MousePointer = vbDefault

Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function


Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim i As Long
    
        
    
    With gData
        .Redraw = flexRDNone
        .Rows = 1
        s = "select *from tablelog where recordlocked=1"
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            
            .Cell(flexcpChecked, r, .ColIndex("Locked")) = flexChecked
            .TextMatrix(r, .ColIndex("Tablename")) = "" & rs("tablename")
            .TextMatrix(r, .ColIndex("Key")) = "" & rs("keyrecord")
            .TextMatrix(r, .ColIndex("LockDate")) = "" & rs("lockdate")
            .TextMatrix(r, .ColIndex("User")) = "" & rs("user_id")
            .TextMatrix(r, .ColIndex("Workstation")) = "" & rs("workstation_name")
            
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
    End With
    mDirty = False
    
    

End Sub


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = gData.ColKey(Col) <> "Locked"
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mDirty = True
    gData.RowData(Row) = "DIRTY"
End Sub
