VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   8535
   ClientTop       =   1545
   ClientWidth     =   6585
   Icon            =   "FProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   1860
      Picture         =   "FProperties.frx":000C
      TabIndex        =   1
      ToolTipText     =   "Login"
      Top             =   2820
      Width           =   1140
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   2445
      Left            =   870
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _cx             =   8705
      _cy             =   4313
      Appearance      =   1
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FProperties.frx":0596
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "FProperties.frx":05D2
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "FProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMode As String


Public Sub ShowArray(ParamArray p())
    Dim r As Long
    mMode = "tip"
    
    With gData
        .Rows = 0
        For r = 0 To UBound(p) Step 2
            If "" & p(r) = "PropertyGroup" Then
                If r > 0 Then .AddItem ""
                .AddItem "" & p(r + 1)
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 1) = vbInactiveTitleBar
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, 1) = vbInactiveTitleBarText
            Else
                .AddItem "" & p(r) & vbTab & p(r + 1)
            End If
        Next
    End With
    Call Form_Resize
    Me.Show vbModal
End Sub

Public Sub ShowSQL(Caption As String, SQL As String, Optional Database As Connections)
    
    Dim i As Long
    Dim fname As String
    Dim rs As Recordset
    Dim section As String
    
    mMode = "form"
    Set rs = HFApp.SqlExec(SQL, Database)
    If rs.EOF Then Exit Sub
    Me.Caption = Caption
    With gData
        .Rows = 0
        For i = 0 To rs.fields.Count - 1
            If UCase(Left(rs(i).Name, 7)) = "SECTION" Then
                If .Rows > 0 Then .AddItem ""
                .AddItem rs(i)
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 1) = vbInactiveTitleBar
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, 1) = vbInactiveTitleBarText
            Else
                .AddItem SpaceCase(rs(i).Name) & vbTab & rs(i)
            End If
        Next
    End With
    Me.Show
End Sub


Public Sub ShowForm(Caption As String, ParamArray p())
    Dim r As Long
    mMode = "form"
    
    r = (UBound(p) + 1) \ 2
    If r * 2 < UBound(p) + 1 Then r = r + 1
    With gData
        .Rows = r
        For r = 0 To .Rows - 1
            .TextMatrix(r, 0) = "" & p(r * 2) & " "
            .TextMatrix(r, 1) = "" & p(r * 2 + 1)
        Next
    End With
    Me.Caption = Caption
    Me.Show vbModal
End Sub


Private Sub cmdNav_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_Resize()
    Const Margin = 120
    
    With gData
        Call .AutoSize(0, 1)
        .Width = .ColWidth(0) + .ColWidth(1)
        .Height = .RowHeight(0) * .Rows
    End With
    
    Select Case mMode
        Case "tip"
            Image1.Visible = False
            cmdNav(0).Visible = False
            gData.FixedCols = 0
            gData.Move 0, 0
            gData.GridLines = flexGridNone
            Me.Height = gData.Height + (Me.Height - Me.ScaleHeight)
            Me.Width = gData.Left + (Me.Width - Me.ScaleWidth)
            
        Case "form"
            Image1.Visible = True
            cmdNav(0).Visible = True
            gData.FixedCols = 1
            gData.GridLines = flexGridFlat
            Me.Height = gData.Height + 3 * Margin + cmdNav(0).Height + (Me.Height - Me.ScaleHeight)
            Me.Width = gData.Left + gData.Width + Margin + (Me.Width - Me.ScaleWidth)
            cmdNav(0).Move (Me.Width - cmdNav(0).Width) / 2, Me.ScaleHeight - Margin - cmdNav(0).Height
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

