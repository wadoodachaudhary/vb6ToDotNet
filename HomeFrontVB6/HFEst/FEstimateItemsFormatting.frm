VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FEstimateItemsFormatting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formatting Rules"
   ClientHeight    =   1140
   ClientLeft      =   4110
   ClientTop       =   1725
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid gFormatting 
      Height          =   960
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4065
      _cx             =   7170
      _cy             =   1693
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FEstimateItemsFormatting.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
End
Attribute VB_Name = "FEstimateItemsFormatting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub ShowBidRules(AcceptedForeColor As Long, AcceptedBackColor As Long, AcceptedStyle As String, _
                        DeclinedForeColor As Long, DeclinedBackColor As Long, DeclinedStyle As String, _
                        MinForeColor As Long, MinBackColor As Long, MinStyle As String, _
                        MaxForeColor As Long, MaxBackColor As Long, MaxStyle As String)

    With gFormatting
        .Cell(flexcpText, 0, 0) = "Accepted Bid"
        .Cell(flexcpText, 1, 0) = "Declined Bid"
        .Cell(flexcpText, 2, 0) = "Minimum Bid"
        .Cell(flexcpText, 3, 0) = "Maximim Bid"
        
        .Cell(flexcpText, 0, 1) = AcceptedStyle
        .Cell(flexcpText, 1, 1) = DeclinedStyle
        .Cell(flexcpText, 2, 1) = MinStyle
        .Cell(flexcpText, 3, 1) = MaxStyle
        
        .Cell(flexcpBackColor, 0, 2) = AcceptedBackColor
        .Cell(flexcpBackColor, 1, 2) = DeclinedBackColor
        .Cell(flexcpBackColor, 2, 2) = MinBackColor
        .Cell(flexcpBackColor, 3, 2) = MaxBackColor

        .Cell(flexcpBackColor, 0, 3) = AcceptedForeColor
        .Cell(flexcpBackColor, 1, 3) = DeclinedForeColor
        .Cell(flexcpBackColor, 2, 3) = MinForeColor
        .Cell(flexcpBackColor, 3, 3) = MaxForeColor
    End With
    Call ApplyFormating
    Me.Show vbModal
    With gFormatting
        AcceptedStyle = .Cell(flexcpText, 0, 1)
        DeclinedStyle = .Cell(flexcpText, 1, 1)
        MinStyle = .Cell(flexcpText, 2, 1)
        MaxStyle = .Cell(flexcpText, 3, 1)
        
        AcceptedBackColor = .Cell(flexcpBackColor, 0, 2)
        DeclinedBackColor = .Cell(flexcpBackColor, 1, 2)
        MinBackColor = .Cell(flexcpBackColor, 2, 2)
        MaxBackColor = .Cell(flexcpBackColor, 3, 2)

        AcceptedForeColor = .Cell(flexcpBackColor, 0, 3)
        DeclinedForeColor = .Cell(flexcpBackColor, 1, 3)
        MinForeColor = .Cell(flexcpBackColor, 2, 3)
        MaxForeColor = .Cell(flexcpBackColor, 3, 3)
    End With
    Unload Me

End Sub

Public Sub ShowPurchasingRules()
    With gFormatting
        .Cell(flexcpText, 0, 1) = HFApp.Options.Value(Format_QtyRateEQZero_FontStyle)
        .Cell(flexcpText, 1, 1) = HFApp.Options.Value(Format_QtyRateLTZero_FontStyle)
        .Cell(flexcpText, 2, 1) = HFApp.Options.Value(Format_NonSysRate_FontStyle)
        .Cell(flexcpText, 3, 1) = HFApp.Options.Value(Format_InvalidData_FontStyle)
        
        .Cell(flexcpBackColor, 0, 2) = HFApp.Options.Value(Format_QtyRateEQZero_BackColor)
        .Cell(flexcpBackColor, 1, 2) = HFApp.Options.Value(Format_QtyRateLTZero_BackColor)
        .Cell(flexcpBackColor, 2, 2) = HFApp.Options.Value(Format_NonSysRate_BackColor)
        .Cell(flexcpBackColor, 3, 2) = HFApp.Options.Value(Format_InvalidData_BackColor)

        .Cell(flexcpBackColor, 0, 3) = HFApp.Options.Value(Format_QtyRateEQZero_ForeColor)
        .Cell(flexcpBackColor, 1, 3) = HFApp.Options.Value(Format_QtyRateLTZero_ForeColor)
        .Cell(flexcpBackColor, 2, 3) = HFApp.Options.Value(Format_NonSysRate_ForeColor)
        .Cell(flexcpBackColor, 3, 3) = HFApp.Options.Value(Format_InvalidData_ForeColor)
    End With
    Call ApplyFormating
    Me.Show vbModal
    With gFormatting
        HFApp.Options.Value(Format_QtyRateEQZero_FontStyle) = .Cell(flexcpText, 0, 1)
        HFApp.Options.Value(Format_QtyRateLTZero_FontStyle) = .Cell(flexcpText, 1, 1)
        HFApp.Options.Value(Format_NonSysRate_FontStyle) = .Cell(flexcpText, 2, 1)
        HFApp.Options.Value(Format_InvalidData_FontStyle) = .Cell(flexcpText, 3, 1)
        
        HFApp.Options.Value(Format_QtyRateEQZero_BackColor) = .Cell(flexcpBackColor, 0, 2)
        HFApp.Options.Value(Format_QtyRateLTZero_BackColor) = .Cell(flexcpBackColor, 1, 2)
        HFApp.Options.Value(Format_NonSysRate_BackColor) = .Cell(flexcpBackColor, 2, 2)
        HFApp.Options.Value(Format_InvalidData_BackColor) = .Cell(flexcpBackColor, 3, 2)

        HFApp.Options.Value(Format_QtyRateEQZero_ForeColor) = .Cell(flexcpBackColor, 0, 3)
        HFApp.Options.Value(Format_QtyRateLTZero_ForeColor) = .Cell(flexcpBackColor, 1, 3)
        HFApp.Options.Value(Format_NonSysRate_ForeColor) = .Cell(flexcpBackColor, 2, 3)
        HFApp.Options.Value(Format_InvalidData_ForeColor) = .Cell(flexcpBackColor, 3, 3)
    End With
    Call HFApp.Options.SaveData
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub gFormatting_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call ApplyFormating
End Sub

Private Sub gFormatting_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = 0
End Sub

Private Sub gFormatting_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim c As Long
    c = gFormatting.Cell(flexcpBackColor, Row, Col)
    If VBChooseColor(c, False, , , Me.hwnd) Then
        gFormatting.Cell(flexcpBackColor, Row, Col) = c
        Call ApplyFormating
    End If
End Sub

Private Sub ApplyFormating()
    Dim r As Long
    With gFormatting
        For r = 0 To 3
        
            'stupid flexgrid reads 0 as white - substitue something "almost" black
            If .Cell(flexcpBackColor, r, 3) = 0 Then .Cell(flexcpBackColor, r, 3) = 1
            If .Cell(flexcpBackColor, r, 2) = 0 Then .Cell(flexcpBackColor, r, 2) = 1
        
            .Cell(flexcpForeColor, r, 0) = .Cell(flexcpBackColor, r, 3)
            .Cell(flexcpBackColor, r, 0) = .Cell(flexcpBackColor, r, 2)
            .Cell(flexcpFontBold, r, 0) = InStr(1, .Cell(flexcpText, r, 1), "Bold")
            .Cell(flexcpFontItalic, r, 0) = InStr(1, .Cell(flexcpText, r, 1), "Italic")
        Next
    End With
End Sub


