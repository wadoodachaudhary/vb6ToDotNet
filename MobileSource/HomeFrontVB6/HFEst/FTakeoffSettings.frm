VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FTakeoffSettings 
   Caption         =   "Takeoff Settings"
   ClientHeight    =   6270
   ClientLeft      =   885
   ClientTop       =   2400
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "FTakeoffSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   4815
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2310
      TabIndex        =   1
      Top             =   5730
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3450
      TabIndex        =   2
      Top             =   5730
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4275
      _cx             =   7541
      _cy             =   9340
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
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   42
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FTakeoffSettings.frx":000C
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
Attribute VB_Name = "FTakeoffSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCancel   As Boolean
Private mSettings As String
Private mJob      As String

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then Call SaveData
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    
    gData.Move margin, margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - 3 * margin - cmdNav(0).Height
    
    cmdNav(0).Move Me.ScaleWidth - 2 * margin - 2 * cmdNav(0).Width, Me.ScaleHeight - margin - cmdNav(1).Height
    cmdNav(1).Move Me.ScaleWidth - margin - 1 * cmdNav(0).Width, Me.ScaleHeight - margin - cmdNav(1).Height
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Public Function GetSettings(Job As String, Settings As String) As String
    Dim s As String
    Dim rs As Recordset
    Dim i As Long
    
    mJob = Job
    mSettings = Settings
    
    'load wbs descriptions
    s = "select * from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(str, mJob)
    Set rs = HFApp.SqlExec(s)
    With gData
        For i = 1 To 40
            .TextMatrix(i + 1, 0) = "" & rs("WBSDesc" & format(i, "00"))
            .RowHidden(i + 1) = .TextMatrix(i + 1, 0) = ""
        Next
        
        'set current values
        For i = 1 To 41
            .TextMatrix(i, 1) = Parse(mSettings, i, Chr(1))
        Next
        
        Call .AutoSize(0, .cols - 1)
        
    End With
    
    mSettings = ""
    
    Me.Show vbModal
    GetSettings = mSettings
    
End Function

Private Sub SaveData()
    Dim i As Long
    mSettings = ""
    For i = 1 To 41
        mSettings = mSettings & Chr(1) & gData.TextMatrix(i, 1)
    Next
    mSettings = Mid(mSettings, 2)
    
    
    'if empty return empty
    If Trim(Replace(mSettings, Chr(1), "")) = "" Then
        mSettings = ""
    End If
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim fieldname As String
    
    With gData
        .EditMaxLength = 50
        
        If Row = 1 Then
            fieldname = "Location"
        Else
            fieldname = "WBS" & format(Row - 1, "00")
        End If
        .ComboList = GetComboList(fieldname)
    End With
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then gData.Clip = ""
End Sub

Private Function GetComboList(fieldname As String) As String
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    s = "select distinct " & fieldname & " from estimateitems where job=" & DbQuote(str, mJob) & " order by 1"
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    s = ""
    While Not rs.EOF
        If Trim("" & rs(0)) <> "" Then s = s & "|" & rs(0)
        rs.MoveNext
    Wend
    If s = "|" Then s = ""
    GetComboList = s
    
End Function

