VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FWBSCodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WBS Codes"
   ClientHeight    =   6255
   ClientLeft      =   5775
   ClientTop       =   2445
   ClientWidth     =   4815
   Icon            =   "FWBSCodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3510
      TabIndex        =   2
      Top             =   5730
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2340
      TabIndex        =   1
      Top             =   5730
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   5295
      Left            =   270
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
      Rows            =   41
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FWBSCodes.frx":000C
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
Attribute VB_Name = "FWBSCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mJob As String

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then Call SaveData
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    gData.AutoSearch = flexSearchNone
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Public Sub ShowForm(Optional Job As String)
    Dim s As String
    Dim rs As Recordset
    Dim i As Long
    
    mJob = Job
    
    If mJob = "" Then
        'get standards
        s = "select Item,custom_description Description from customDescriptions where item like 'WBS__' order by item"
    Else
        'get job list
        s = ""
        For i = 1 To 40
            s = s & "union select 'WBS" & format(i, "00") & "' Item, WBSDesc" & format(i, "00") & " Description from tblJobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(str, mJob) & vbCrLf
        Next
        s = Mid(s, 7)
    End If
    
    
    Set rs = HFApp.SqlExec(s)
    With gData
    While Not rs.EOF
        .TextMatrix(Val(Mid("" & rs("item"), 4)), 1) = "" & rs("Description")
        rs.MoveNext
    Wend
    End With
    
    Me.Show vbModal
    
End Sub

Private Sub SaveData()
    Dim i As Long
    Dim s As String
    With gData
    For i = 1 To 40
        If mJob = "" Then
            s = "update customdescriptions set custom_description=" & DbQuote(str, .TextMatrix(i, 1)) & " where item=" & DbQuote(str, "WBS" & format(i, "00"))
        Else
            s = "update tbljobs set wbsdesc" & format(i, "00") & "=" & DbQuote(str, .TextMatrix(i, 1)) & " where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(str, mJob)
        End If
        Call HFApp.SqlExec(s)
    Next
    End With
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    gData.EditMaxLength = 50
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then gData.Clip = ""
End Sub
