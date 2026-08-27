VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FPriceList 
   Caption         =   "Vendor Item Prices"
   ClientHeight    =   5295
   ClientLeft      =   2850
   ClientTop       =   2730
   ClientWidth     =   11100
   Icon            =   "FPriceList.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   11100
   Begin HFEst.Slider Slider 
      Height          =   3855
      Left            =   2730
      Top             =   540
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   6800
      Max             =   7275
   End
   Begin VSFlex8Ctl.VSFlexGrid gAssemblies 
      Height          =   2655
      Left            =   60
      TabIndex        =   1
      Top             =   585
      Width           =   2595
      _cx             =   1982796257
      _cy             =   1982796363
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FPriceList.frx":000C
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   1058
      ButtonWidth     =   1799
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "NewPricelist"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Key             =   "View"
            Style           =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "More"
            Key             =   "IncreaseDecimals"
            Object.ToolTipText     =   "Increase decimal places"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Less"
            Key             =   "DecreaseDecimals"
            Object.ToolTipText     =   "Decrease decimal places"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Key             =   "Preview"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Key             =   "PricelistExport"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            Key             =   "PricelistImport"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forecasting"
            Key             =   "Forecast"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   9420
         Top             =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2655
      Left            =   2820
      TabIndex        =   2
      Top             =   600
      Width           =   6975
      _cx             =   1982803983
      _cy             =   1982796363
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
      BackColorSel    =   14336431
      ForeColorSel    =   -2147483640
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   45
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FPriceList.frx":005E
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
      FrozenRows      =   1
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483624
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "FPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FPriceList::"

Private Type ViewDefs
    Name         As String
    KeyFlds      As String
    DisplayFlds  As String
End Type
Private mDirty                  As Boolean
Private mViews()                As ViewDefs

Private mViewIndex As Long
Private mDecimals  As Long

Private mTimerTask As String 'stupid menus

Private Const mcITEM_RENUMBER = 0
Private Const mcITEM_NEWITEM = 1
Private Const mcITEM_DUPLICATE = 2
Private Const mcITEM_VIEWFILES = 4
Private Const mcITEM_DELETE = 8
Private Const mcITEM_PRICEGROUPS = 10

'used purely for debugging
Private mViewQuery As String

Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    Dim i As Long
    Dim PricingGroup As String
    Dim r As Long
    Dim c As Long
    
    With gItems
        If Row = 1 Then
            .Redraw = flexRDNone
            For r = 2 To .Rows - 1
                If .RowData(r) <> "Deleted" Then
                    .RowHidden(r) = False
                End If
                'If .EditText <> "" Then
                    For c = 0 To .Cols - 1
                        If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                            
                            .RowHidden(r) = True
                            Exit For
                        End If
                    Next
                'End If
            Next
            .Redraw = flexRDBuffered
            Exit Sub
        End If
    End With
    
    
    With gItems
    
        'UPDATE Others in this pricegroup that are visible - this is cosmetic only
        'trigger TIU_tblVendorCost maintains data
        
        Select Case .ColKey(Col)
            Case "Expiry3", "Expiry2", "Expiry1", "Effective1", "Effective2", "Last3", "Last2", "Last1", "Current", "Next1", "Next2", "Forecast1", "Forecast2", "Forecast3", "Forecast4", "Forecast5", "Forecast6", "Forecast7", "Forecast8", "Forecast9", "Forecast10", "Forecast11", "Forecast12"
                
                If Val(.Cell(flexcpData, Row, .ColIndex("PriceLink"))) <> 0 Then
                
                    PricingGroup = Val(.Cell(flexcpData, Row, .ColIndex("PriceLink"))) & Chr(4) & _
                                   .Cell(flexcpText, Row, .ColIndex("Vendor")) & Chr(4) & _
                                   .Cell(flexcpText, Row, .ColIndex("Assembly")) & Chr(4) & _
                                   .Cell(flexcpText, Row, .ColIndex("Model")) & Chr(4) & _
                                   .Cell(flexcpText, Row, .ColIndex("Community")) & Chr(4) & _
                                   .Cell(flexcpText, Row, .ColIndex("CommunityPhase"))
                                   
                    For i = 1 To .Rows - 1
                    
                        If PricingGroup = Val(.Cell(flexcpData, i, .ColIndex("PriceLink"))) & Chr(4) & _
                                          .Cell(flexcpText, i, .ColIndex("Vendor")) & Chr(4) & _
                                          .Cell(flexcpText, i, .ColIndex("Assembly")) & Chr(4) & _
                                          .Cell(flexcpText, i, .ColIndex("Model")) & Chr(4) & _
                                          .Cell(flexcpText, i, .ColIndex("Community")) & Chr(4) & _
                                          .Cell(flexcpText, i, .ColIndex("CommunityPhase")) Then

                            .TextMatrix(i, .ColIndex("Current")) = .TextMatrix(Row, .ColIndex("Current"))
                            .TextMatrix(i, .ColIndex("Effective1")) = .TextMatrix(Row, .ColIndex("Effective1"))
                            .TextMatrix(i, .ColIndex("Effective2")) = .TextMatrix(Row, .ColIndex("Effective2"))
                            .TextMatrix(i, .ColIndex("Expiry1")) = .TextMatrix(Row, .ColIndex("Expiry1"))
                            .TextMatrix(i, .ColIndex("Expiry2")) = .TextMatrix(Row, .ColIndex("Expiry2"))
                            .TextMatrix(i, .ColIndex("Expiry3")) = .TextMatrix(Row, .ColIndex("Expiry3"))
                            .TextMatrix(i, .ColIndex("Next1")) = .TextMatrix(Row, .ColIndex("Next1"))
                            .TextMatrix(i, .ColIndex("Next2")) = .TextMatrix(Row, .ColIndex("Next2"))
                            .TextMatrix(i, .ColIndex("Last1")) = .TextMatrix(Row, .ColIndex("Last1"))
                            .TextMatrix(i, .ColIndex("Last2")) = .TextMatrix(Row, .ColIndex("Last2"))
                            .TextMatrix(i, .ColIndex("Last3")) = .TextMatrix(Row, .ColIndex("Last3"))
                            .TextMatrix(i, .ColIndex("Forecast1")) = .TextMatrix(Row, .ColIndex("Forecast1"))
                            .TextMatrix(i, .ColIndex("Forecast2")) = .TextMatrix(Row, .ColIndex("Forecast2"))
                            .TextMatrix(i, .ColIndex("Forecast3")) = .TextMatrix(Row, .ColIndex("Forecast3"))
                            .TextMatrix(i, .ColIndex("Forecast4")) = .TextMatrix(Row, .ColIndex("Forecast4"))
                            .TextMatrix(i, .ColIndex("Forecast5")) = .TextMatrix(Row, .ColIndex("Forecast5"))
                            .TextMatrix(i, .ColIndex("Forecast6")) = .TextMatrix(Row, .ColIndex("Forecast6"))
                            .TextMatrix(i, .ColIndex("Forecast7")) = .TextMatrix(Row, .ColIndex("Forecast7"))
                            .TextMatrix(i, .ColIndex("Forecast8")) = .TextMatrix(Row, .ColIndex("Forecast8"))
                            .TextMatrix(i, .ColIndex("Forecast9")) = .TextMatrix(Row, .ColIndex("Forecast9"))
                            .TextMatrix(i, .ColIndex("Forecast10")) = .TextMatrix(Row, .ColIndex("Forecast10"))
                            .TextMatrix(i, .ColIndex("Forecast11")) = .TextMatrix(Row, .ColIndex("Forecast11"))
                            .TextMatrix(i, .ColIndex("Forecast12")) = .TextMatrix(Row, .ColIndex("Forecast12"))
                        End If
                    Next
                End If
        End Select
    End With
End Sub


Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
'On Error Resume Next
'change selection to only single column
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    gItems.ColSel = gItems.Col
    bInHere = False
End Sub

Private Sub gItems_AfterSort(ByVal Col As Long, Order As Integer)
gItems.FixedRows = 1
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   gItems.AutoSearch = flexSearchNone
    If Row = 1 Then
        gItems.ComboList = ""
        Cancel = gItems.ColDataType(Col) = flexDTBoolean
    Else
        With gItems
            .ComboList = ""
            .AutoSearch = flexSearchNone
            Select Case .ColKey(Col)
            
                Case "Effective1", "Effective2"
                    .ComboList = "|..."
                
                Case "Current", "Next1", "Next2", "Forecast1", "Forecast2", "Forecast3", "Forecast4", "Forecast5", "Forecast6", "Forecast7", "Forecast8", "Forecast9", "Forecast10", "Forecast11", "Forecast12"
                
                Case "TaxGroup"
                    .ComboList = "|..."
                
                Case "SKU"
                    .EditMaxLength = 50
                            
                Case Else
                    Cancel = True
                    .AutoSearch = flexSearchFromCursor
    
            End Select
        End With
    End If
End Sub

Private Sub gItems_BeforeSort(ByVal Col As Long, Order As Integer)
gItems.FixedRows = 2
End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim i As Long
    With gItems
        Select Case .ColKey(Col)
            Case "Effective1", "Effective2"
                Call DCalendar.Popup(gItems, .RowPos(.Row) + .RowHeight(.Row), .colPos(.Col))
                mDirty = True
                For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    If .RowHidden(i) = False Then
                        .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                        .RowData(i) = "dirty"
                    End If
                Next
                
            Case "TaxGroup"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Group", "SELECT TaxGroup,Description FROM TaxGroups where DivisionID = " & HFApp.DivisionID, s) Then
                    .Cell(flexcpText, .Row, .ColIndex("TaxGroup"), .RowSel, .ColIndex("TaxGroup")) = FPickList.SelectedItem("TaxGroup")
                    mDirty = True
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        If .RowHidden(i) = False Then
                            .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                            .RowData(i) = "dirty"
                        End If
                    Next
                End If
        End Select
    End With
End Sub


Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r As Long
    Dim b As Boolean
    
    Select Case True
        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.ShowForm(gItems)
            
        Case KeyCode = vbKeyDelete And Shift = vbCtrlMask
            With gItems
                If .Row < 1 Then Exit Sub
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    If Not .RowHidden(r) = True Then
                        .RowHidden(r) = True
                        .RowData(r) = "Deleted"
                        mDirty = True
                    End If
                Next
                
                On Error GoTo ExitSub
                For r = Min(.Row, .RowSel) To 1 Step -1
                    If Not .RowHidden(r) Then
                        .Row = r
                        Exit Sub
                    End If
                Next
                For r = Max(.Row, .RowSel) To .Rows - 1
                    If Not .RowHidden(r) Then
                        .Row = r
                        Exit Sub
                    End If
                Next
            End With
        
        Case KeyCode = vbKeyDelete
            With gItems
                If .Row > 0 Then
                    Call gItems_BeforeEdit(.Row, .Col, b)
                    If Not b Then
                        .TextMatrix(.Row, .Col) = ""
                        Call gItems_ValidateEdit(.Row, .Col, b)
                        Call gItems_AfterEdit(.Row, .Col)
                    End If
                End If
                
            End With
            
        
            
    End Select
ExitSub: Exit Sub
End Sub

Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    Dim PriceLink As Long
    Dim s As String
    If Row = 1 Then Exit Sub
    
    With gItems
    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
    If Not .RowHidden(i) Then
    
        Select Case .ColKey(Col)
            Case "Effective1", "Effective2"
                If Trim(.EditText) = "" Then
                Else
                    If IsDate(.EditText) Then
                        .EditText = format(.EditText, HFApp.Options(DateFormat))
                        .TextMatrix(i, Col) = .EditText
                    Else
                        Cancel = True
                    End If
                End If
                
            Case "Current"
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                Else
                    If .Cell(flexcpData, i, Col) <> "changed" Then
                        .Cell(flexcpData, i, Col) = "changed"
                        .TextMatrix(i, .ColIndex("Last3")) = .TextMatrix(i, .ColIndex("Last2"))
                        .TextMatrix(i, .ColIndex("Last2")) = .TextMatrix(i, .ColIndex("Last1"))
                        .TextMatrix(i, .ColIndex("Last1")) = .TextMatrix(i, .ColIndex("Current"))
                        .TextMatrix(i, .ColIndex("Expiry3")) = .TextMatrix(i, .ColIndex("Expiry2"))
                        .TextMatrix(i, .ColIndex("Expiry2")) = .TextMatrix(i, .ColIndex("Expiry1"))
                        .TextMatrix(i, .ColIndex("Expiry1")) = VBA.Date()
                    End If
                    .TextMatrix(i, Col) = .EditText
                End If
                
                    
            Case "Next1", "Next2", "Forecast1", "Forecast2", "Forecast3", "Forecast4", "Forecast5", "Forecast6", "Forecast7", "Forecast8", "Forecast9", "Forecast10", "Forecast11", "Forecast12"
                If IsNumeric(.EditText) Then
                    .TextMatrix(i, Col) = .EditText
                Else
                    Cancel = True
                End If
                        
            Case "TaxGroup"
                s = .EditText
                Cancel = Not ValidateField(gItems, s, "Tax group not found", "SELECT TaxGroup FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, .EditText))
                .EditText = s
                .TextMatrix(i, Col) = .EditText
                
            Case "SKU"
                .TextMatrix(i, Col) = .EditText
            
            Case Else
                Cancel = True
                
        End Select
    
    
                
    
    
        If Not Cancel Then
            mDirty = True
            If .RowHidden(i) = False Then .RowData(i) = "dirty"
        End If
        
    End If
    Next
    End With
    
End Sub



Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Select Case mTimerTask
        Case "newitem"
            Call Toolbar_ButtonClick(Toolbar.Buttons("NewPricelist"))
        Case "attachments"
            With Me.gItems
            Call HFApp.RunTask("EditAttachments" & _
                               "|PLST~" & .TextMatrix(.Row, .ColIndex("Community")) & "~" & .TextMatrix(.Row, .ColIndex("CommunityPhase")) & "~" & .TextMatrix(.Row, .ColIndex("Assembly")) & "~" & .TextMatrix(.Row, .ColIndex("Model")) & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Phase Quote" & _
                               "|PLST~" & .TextMatrix(.Row, .ColIndex("Community")) & "~~" & .TextMatrix(.Row, .ColIndex("Assembly")) & "~" & .TextMatrix(.Row, .ColIndex("Model")) & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Community Quote" & _
                               "|PLST~~~" & .TextMatrix(.Row, .ColIndex("Assembly")) & "~" & .TextMatrix(.Row, .ColIndex("Model")) & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Global Quote" & _
                               "|PLST~" & .TextMatrix(.Row, .ColIndex("Community")) & "~" & .TextMatrix(.Row, .ColIndex("CommunityPhase")) & "~~~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Phase Rate" & _
                               "|PLST~" & .TextMatrix(.Row, .ColIndex("Community")) & "~~~~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Community Rate" & _
                               "|PLST~~~~~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Global Rate" & _
                               "|ASM~" & .TextMatrix(.Row, .ColIndex("Community")) & "~" & .TextMatrix(.Row, .ColIndex("Model")) & "~~" & .TextMatrix(.Row, .ColIndex("Assembly")) & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Assembly Master" & _
                               "|ASM~~" & .TextMatrix(.Row, .ColIndex("Model")) & "~~" & .TextMatrix(.Row, .ColIndex("Assembly")) & "~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Assembly Master" & _
                               "|ITM~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Item Database")
            End With
    End Select
    mTimerTask = ""
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    
    Dim s As String
    Dim f As Form
    
    Dim Vendor    As String
    Dim community As String
    Dim CommunityPhase As String
    Dim Assembly  As String
    
    Select Case Button.key
        Case "DecreaseDecimals": Call SetDecimalPlaces(mDecimals - 1)
        Case "IncreaseDecimals": Call SetDecimalPlaces(mDecimals + 1)
    
        Case "NewPricelist"
            'get defaults from items list if available
            If ActiveControl Is gItems Then
                Vendor = gItems.TextMatrix(gItems.Row, gItems.ColIndex("Vendor"))
                community = gItems.TextMatrix(gItems.Row, gItems.ColIndex("Community"))
                If community = "" Then community = gAssemblies.Cell(flexcpText, gAssemblies.Row, 0)
                CommunityPhase = gItems.TextMatrix(gItems.Row, gItems.ColIndex("CommunityPhase"))
                Assembly = gItems.TextMatrix(gItems.Row, gItems.ColIndex("Assembly"))
            End If
            If ActiveControl Is gAssemblies Then
                'get vendor from view tree if available
                Vendor = Parse(Parse(gAssemblies.RowData(gAssemblies.Row), 2, "Vendor = "), 2, "'")
                community = Parse(Parse(gAssemblies.RowData(gAssemblies.Row), 2, "Community = "), 2, "'")
                If community = "" Then community = gAssemblies.Cell(flexcpText, gAssemblies.Row, 0)
                CommunityPhase = Parse(Parse(gAssemblies.RowData(gAssemblies.Row), 2, "CommunityPhase = "), 2, "'")
                Assembly = Parse(Parse(gAssemblies.RowData(gAssemblies.Row), 2, "Assembly = "), 2, "'")
            End If
            
            If SaveData(True) Then Call FAddPricelist.ShowForm(Vendor, community, CommunityPhase)
            
            
        Case "Save"
            Call SaveData(False)
        
        Case "View"
            Call Toolbar_ButtonDropDown(Button)
        
        Case "Preview"
            If mDirty Then
                If vbOK = MsgBox("You must save your changes before opening the print preview window.", vbOKCancel + vbQuestion, App.ProductName) Then
                     If Not SaveData(False) Then Exit Sub
                Else
                    Exit Sub
                End If
            End If
            s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\PriceList.rpt")
            'Call FRptViewer.ShowReport(s, True, True)
            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            'On Error GoTo eh
            Call c.PrintPreview("Print Preview")
            
        
        Case "PricelistExport"
            If SaveData(True) Then
                On Error Resume Next
                Call FExportPricelists.ShowForm
            End If
        
        Case "PricelistImport"
            If SaveData(True) Then
                On Error Resume Next
                Call FImportPricelists.ShowForm
            End If
            
        Case "Forecast"
            Call FCostForecast.Show(vbModal)
            
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonClick")
End Sub

Private Sub Toolbar_ButtonDropDown(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim ParentMenu As Long
    Dim i As Long
    
    Select Case Button.key
        Case "View"
            With FMain.PopMenu
                ParentMenu = .MenuIndex("mnuPriceListViews")
                Call .ClearSubMenusOfItem(ParentMenu)
                For i = 0 To UBound(mViews)
                    .AddItem mViews(i).Name, "PriceListView" & i, , i, ParentMenu, , i = mViewIndex
                Next
            End With
            PopupMenu FMain.mnuPriceListViews, , Button.Left, Button.Top + Button.Height
    End Select
    Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonDropDown")
End Sub
Public Sub mnuPriceListViewsSub_Click(Index As Integer)
    If Not SaveData(True) Then Exit Sub
    
    Call IniPutGrid(Me, gItems, , mViewIndex)
    mViewIndex = Index
    Call IniGetGrid(Me, gItems, , , mViewIndex)
    Call LoadAssemblies(True)
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    
    mViewIndex = IniGet(AppIni, Me.Name, "mViewIndex", 0)
    mDecimals = IniGet(AppIni, Me.Name, "Decimals", -1)
    Call SetDecimalPlaces(mDecimals)
    
    Call IniGetGrid(Me, gItems, , , mViewIndex)
    Call IniGetForm(Me)
    Call LoadCustomDescriptions
    
    Call LoadCostTypes(, gItems)
    Call LoadViews
    Me.Show

End Sub


Private Sub LoadViews()
    Dim i As Long
    ReDim mViews(3) As ViewDefs
    
    mViews(i).Name = "Assembly"
    mViews(i).KeyFlds = "CommunityDesc,CommunityPhase,Assembly,CostCode"
    mViews(i).DisplayFlds = "Community + ' - ' + CommunityDesc,CommunityPhase + ' - ' + CommunityPhaseDesc,Assembly + ' - ' + AssemblyDesc,CostCode + ' - ' + CostCodeDesc"
    i = i + 1
            
    mViews(i).Name = "Vendor"
    mViews(i).KeyFlds = "Vendor,CommunityDesc,CommunityPhase,Assembly"
    mViews(i).DisplayFlds = "VendorName,Community + ' - ' + CommunityDesc,CommunityPhase + ' - ' + CommunityPhaseDesc,Assembly + ' - ' + AssemblyDesc"
    i = i + 1
    
    mViews(i).Name = "Purchase Order"
    mViews(i).KeyFlds = "POIndex,CommunityDesc,CommunityPhase,Assembly"
    mViews(i).DisplayFlds = "POIndex + ' ' + case when poindex=POIndexDescription then '' else POIndexDescription end,Community + ' - ' + CommunityDesc,CommunityPhase + ' - ' + CommunityPhaseDesc,Assembly + ' - ' + AssemblyDesc"
    i = i + 1
    
    mViews(i).Name = "Cost Code"
    mViews(i).KeyFlds = "CostCode,CommunityDesc,CommunityPhase,Assembly"
    mViews(i).DisplayFlds = "CostCode + ' - ' + CostCodeDesc,Community + ' - ' + CommunityDesc,CommunityPhase + ' - ' + CommunityPhaseDesc,Assembly + ' - ' + AssemblyDesc"
    i = i + 1
    
    Call IniGet(AppIni, Me.Name, "mViewIndex", mViewIndex)

    Call LoadAssemblies(True)

End Sub



Private Sub LoadAssemblies(ClearTree As Boolean)
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
    Dim KeyValue   As String
    Dim DisplayValue   As String

'PROPERTY                  CONTAINS
'.rowdata()                Where clause
'.cell(flexcpText, r, 0)   DisplayFld Value
'.cell(flexcpData, r, 0)   DisplayFld Field Name (also used as tooltip)
'.Cell(flexcpText, r, 1)   KeyFld Value
'.Cell(flexcpData, r, 1)   KeyFld Field Name

    
    
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
            On Error GoTo 0
            
            KeyFld = Parse(mViews(mViewIndex).KeyFlds, Level + 2)
            DisplayFld = Parse(mViews(mViewIndex).DisplayFlds, Level + 2)
            
            s = ""
            s = s & "SELECT DISTINCT " & KeyFld & "," & DisplayFld & vbCrLf
            s = s & "  FROM VendorPrices" & vbCrLf
            s = s & " WHERE divisionid in(0," & HFApp.DivisionID & ")" & vbCrLf
            On Error Resume Next
            WhereClause = .RowData(.Row)
            On Error GoTo 0
            If WhereClause <> "" Then s = s & WhereClause & vbCrLf
            s = s & " ORDER BY 2"
            
            mViewQuery = s
            Set rs = HFApp.SqlExec(s)
            
            While Not rs.EOF
                KeyValue = "" & rs(0)
                DisplayValue = Trim("" & rs(1))
                
                If DisplayValue = "- global -" Then
                    DisplayValue = "Global (any " & FMain.CD_Community & ")"
                End If
                If DisplayValue = "- corporate -" Then
                    DisplayValue = "Corporate (any Division)"
                End If
                
                If Level = -1 Then
                    r = .Rows
                    Call .AddItem(DisplayValue, r)
                    .Cell(flexcpData, r, 0) = DisplayFld
                    .Cell(flexcpText, r, 1) = KeyValue
                    .Cell(flexcpData, r, 1) = KeyFld
                    
                      
                    If IsIn(DisplayValue, "- any " & FMain.CD_Community & " -", "- any phase -", "- any assembly -", "Global (any " & FMain.CD_Community & ")", "Corporate (any Division)") Then
                        .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbHighlight
                    End If
                    
            
                    .IsSubtotal(r) = True
                    .RowOutlineLevel(r) = Level + 1
                    If Level + 2 <> levels Then
                        Call .GetNode(r).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                    End If
                    Set n = .GetNode(r)
                    n.Expanded = False
                    
                    .RowData(r) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Str, KeyValue)
                    
                Else
                    r = .Row
                    Set n = .GetNode(r).AddNode(flexNTLastChild, DisplayValue)
                    .Cell(flexcpData, n.Row, 0) = DisplayFld
                    .Cell(flexcpText, n.Row, 1) = KeyValue
                    .Cell(flexcpData, n.Row, 1) = KeyFld
                    If Level + 2 < levels Then
                        Call .GetNode(n.Row).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                    End If
                    n.Expanded = False
                    .RowData(n.Row) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Str, KeyValue)
                    
                    If IsIn(DisplayValue, "- any " & FMain.CD_Community & " -", "- any phase -", "- any assembly -", "Global (any " & FMain.CD_Community & ")", "Corporate (any Division)") Then
                        .Cell(flexcpForeColor, n.Row, 0, n.Row, .Cols - 1) = vbHighlight
                    End If
                    
                End If
                
                rs.MoveNext
            Wend
        End If
        
        Call .AutoSize(0)
        .Redraw = flexRDBuffered
    End With
End Sub



Private Sub gAssemblies_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
Static bInHere As Boolean
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


Private Sub gAssemblies_DblClick()
    If gAssemblies.GetNode.Expanded Then
        Call gAssemblies_KeyDown(vbKeyReturn, 0)
    Else
        gAssemblies.IsCollapsed(gAssemblies.Row) = flexOutlineExpanded
    End If
End Sub

Private Sub gAssemblies_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim s As String
    Dim i As Long
    If InIde And gAssemblies.MouseRow = 0 Then
        s = ""
        s = s & "View Definition:" & vbCrLf
        s = s & "------------------------------------" & vbCrLf
        For i = 1 To Parse(mViews(mViewIndex).KeyFlds)
            s = s & String(4 * (i - 1), " ") & Parse(mViews(mViewIndex).KeyFlds, i) & vbCrLf
        Next
        s = s & vbCrLf & vbCrLf
        s = s & "Last Query:" & vbCrLf
        s = s & "------------------------------------" & vbCrLf
        s = s & mViewQuery
        MsgBox s, vbInformation, "View Definition"
    End If
End Sub

Private Sub gAssemblies_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim i As Long
    i = gAssemblies.MouseRow
    If i = -1 Then
        gAssemblies.ToolTipText = ""
    Else
        gAssemblies.ToolTipText = PrettyName(gAssemblies.Cell(flexcpData, gAssemblies.MouseRow, 0))
    End If
End Sub

Private Sub Slider_Move()
    Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Slider.Min = 960
    Slider.Max = Me.ScaleWidth - 960
    Slider.Move Slider.Left, Toolbar.Height, Slider.Width, Me.ScaleHeight - Toolbar.Height

    gAssemblies.Move 0, Slider.Top, Slider.Left, Slider.Height
    gItems.Move Slider.Left + Slider.Width, Slider.Top, Me.ScaleWidth - Slider.Left - Slider.Width, Slider.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gItems, , mViewIndex)
    Call IniPut(AppIni, Me.Name, "mViewIndex", mViewIndex)
    Call IniPut(AppIni, Me.Name, "Decimals", mDecimals)
    Unload FComments
End Sub

Private Sub gAssemblies_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim Level As Long
    Dim levels As Long
    
    With gAssemblies
        If .Row < 0 Then Exit Sub
        Select Case KeyCode

            Case vbKeyReturn
            
                levels = Parse(mViews(mViewIndex).DisplayFlds)
                If .Rows > 1 Then
                    Level = .RowOutlineLevel(.Row)
                End If
                .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                .Cell(flexcpFontBold, .Row, 0) = True
                Call .AutoSize(0, .Cols - 1)
                Call LoadItems
                KeyCode = 0

            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If

            Case vbKeyRight
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

Private Sub LoadItems()
    Dim s As String
    Dim WhereClause As String
    Dim rs As Recordset
    Dim r As Long

    If Not SaveData(True) Then Exit Sub

    With gItems
        Screen.MousePointer = vbHourglass
        .Redraw = flexRDNone
        .Rows = 2

        s = ""
        s = s & "SELECT * FROM VendorPrices v" & vbCrLf
        s = s & "LEFT OUTER JOIN PriceGroups g ON (v.PriceLink=g.PriceGroup and g.DivisionID=" & HFApp.DivisionID & ")" & vbCrLf
        On Error Resume Next
        WhereClause = gAssemblies.RowData(gAssemblies.Row)
        On Error GoTo 0
        If WhereClause <> "" Then
            s = s & " WHERE ISNULL(Vendor,'')<>'' AND " & Mid(WhereClause, 6) & " AND v.DivisionID in(0," & HFApp.DivisionID & ")" & vbCrLf
        Else
            s = s & " WHERE v.DivisionID in(0," & HFApp.DivisionID & ")" & vbCrLf
        End If
        s = s & "ORDER BY " & mViews(mViewIndex).DisplayFlds
        
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1


            .Cell(flexcpData, r, .ColIndex("PriceLink")) = "" & rs("PriceLink")
            .Cell(flexcpText, r, .ColIndex("PriceLink")) = "" & rs("PriceGroupDesc")
            If Val(.Cell(flexcpData, r, .ColIndex("PriceLink"))) <> 0 Then
                .Cell(flexcpPicture, r, .ColIndex("PriceLink")) = FMain.SmallIcons.ListImages.Item("links").Picture
            End If

            .TextMatrix(r, .ColIndex("ItemType")) = "" & rs("ItemType")
            
            .TextMatrix(r, .ColIndex("PriceLevel")) = Replace("" & rs("PriceLevel"), "Community", FMain.CD_Community)
            
            
            
            .TextMatrix(r, .ColIndex("Vendor")) = "" & rs("Vendor")
            .TextMatrix(r, .ColIndex("VendorName")) = "" & rs("VendorName")
            .TextMatrix(r, .ColIndex("Community")) = "" & rs("Community")
            .TextMatrix(r, .ColIndex("CommunityPhase")) = "" & rs("CommunityPhase")
            .TextMatrix(r, .ColIndex("CommunityDesc")) = "" & rs("CommunityDesc")
            .TextMatrix(r, .ColIndex("TaxGroup")) = "" & rs("TaxGroup")
            .TextMatrix(r, .ColIndex("SKU")) = "" & rs("PartNumber")
            .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
            .TextMatrix(r, .ColIndex("ItemNumber")) = "" & rs("ItemNumber")
            .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(r, .ColIndex("POIndexDescription")) = "" & rs("POIndexDescription")
            .TextMatrix(r, .ColIndex("CostCode")) = "" & rs("CostCode")
            .TextMatrix(r, .ColIndex("CostCodeDesc")) = "" & rs("CostCodeDesc")
            .TextMatrix(r, .ColIndex("ItemDesc")) = "" & rs("ItemDesc")
            .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
            .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
            .TextMatrix(r, .ColIndex("AssemblyDesc")) = "" & rs("AssemblyDesc")
            .TextMatrix(r, .ColIndex("Last3")) = "" & rs("Last3")
            .TextMatrix(r, .ColIndex("Expiry3")) = format("" & rs("Expiry3"), HFApp.Options(DateFormat))
            .TextMatrix(r, .ColIndex("Last2")) = "" & rs("Last2")
            .TextMatrix(r, .ColIndex("Expiry2")) = format("" & rs("Expiry2"), HFApp.Options(DateFormat))
            .TextMatrix(r, .ColIndex("Last1")) = "" & rs("Last1")
            .TextMatrix(r, .ColIndex("Expiry1")) = format("" & rs("Expiry1"), HFApp.Options(DateFormat))
            .TextMatrix(r, .ColIndex("Current")) = "" & rs("CurrentCost")
            .TextMatrix(r, .ColIndex("Next1")) = "" & rs("Next1")
            .TextMatrix(r, .ColIndex("Effective1")) = format("" & rs("Effective1"), HFApp.Options(DateFormat))
            .TextMatrix(r, .ColIndex("Next2")) = "" & rs("Next2")
            .TextMatrix(r, .ColIndex("Effective2")) = format("" & rs("Effective2"), HFApp.Options(DateFormat))
            .TextMatrix(r, .ColIndex("Forecast1")) = "" & rs("Forecast1")
            .TextMatrix(r, .ColIndex("Forecast2")) = "" & rs("Forecast2")
            .TextMatrix(r, .ColIndex("Forecast3")) = "" & rs("Forecast3")
            .TextMatrix(r, .ColIndex("Forecast4")) = "" & rs("Forecast4")
            .TextMatrix(r, .ColIndex("Forecast5")) = "" & rs("Forecast5")
            .TextMatrix(r, .ColIndex("Forecast6")) = "" & rs("Forecast6")
            .TextMatrix(r, .ColIndex("Forecast7")) = "" & rs("Forecast7")
            .TextMatrix(r, .ColIndex("Forecast8")) = "" & rs("Forecast8")
            .TextMatrix(r, .ColIndex("Forecast9")) = "" & rs("Forecast9")
            .TextMatrix(r, .ColIndex("Forecast10")) = "" & rs("Forecast10")
            .TextMatrix(r, .ColIndex("Forecast11")) = "" & rs("Forecast11")
            .TextMatrix(r, .ColIndex("Forecast12")) = "" & rs("Forecast12")

            rs.MoveNext
        Wend

        .Redraw = flexRDBuffered
        Screen.MousePointer = vbDefault
    End With
    mDirty = False
End Sub

Public Function SaveData(prompt As Boolean) As Boolean

    Dim i As Long
    Dim s As String

    If Not mDirty Then
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

    Screen.MousePointer = vbHourglass

    With gItems
    For i = .Rows - 1 To 2 Step -1
        If .RowHidden(i) And .RowData(i) = "Deleted" Then
            s = ""
            s = s & "DELETE FROM tblVendorCost" & vbCrLf
            s = s & "WHERE ISNULL(Community,'')=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Community"))) & vbCrLf
            s = s & "  AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, .TextMatrix(i, .ColIndex("CommunityPhase"))) & vbCrLf
            s = s & "  AND ISNULL(Assembly,'')=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf
            s = s & "  AND ISNULL(Model,'')=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Model"))) & vbCrLf
            s = s & "  AND Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
            s = s & "  AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
            s = s & "  AND Vendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Vendor"))) & vbCrLf
            If .TextMatrix(i, .ColIndex("PriceLevel")) = "Corporate (any Division)" Then
                s = s & "  AND DivisionID = 0" & vbCrLf
            Else
                s = s & "  AND DivisionID = " & HFApp.DivisionID & vbCrLf
            End If
            Call HFApp.SqlExec(s)
            .RemoveItem i
        End If
    Next
    
    For i = 2 To .Rows - 1
        If .RowData(i) = "DIRTY" Then
            s = ""
            s = s & "UPDATE tblVendorCost" & vbCrLf
            s = s & "SET Current_Cost=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Current"))) & vbCrLf
            s = s & "   ,Next_Cost1=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Next1"))) & vbCrLf
            s = s & "   ,Next_Effective1=" & DbQuote(Date, .TextMatrix(i, .ColIndex("Effective1"))) & vbCrLf
            s = s & "   ,Next_Cost2=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Next2"))) & vbCrLf
            s = s & "   ,Next_Effective2=" & DbQuote(Date, .TextMatrix(i, .ColIndex("Effective2"))) & vbCrLf
            s = s & "   ,Last_Cost1=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Last1"))) & vbCrLf
            s = s & "   ,Last1_Expiry=" & DbQuote(Date, .TextMatrix(i, .ColIndex("Expiry1"))) & vbCrLf
            s = s & "   ,Last_Cost2=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Last2"))) & vbCrLf
            s = s & "   ,Last2_Expiry=" & DbQuote(Date, .TextMatrix(i, .ColIndex("Expiry2"))) & vbCrLf
            s = s & "   ,Last_Cost3=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Last3"))) & vbCrLf
            s = s & "   ,Last3_Expiry=" & DbQuote(Date, .TextMatrix(i, .ColIndex("Expiry3"))) & vbCrLf
            s = s & "   ,Forecast1=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast1"))) & vbCrLf
            s = s & "   ,Forecast2=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast2"))) & vbCrLf
            s = s & "   ,Forecast3=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast3"))) & vbCrLf
            s = s & "   ,Forecast4=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast4"))) & vbCrLf
            s = s & "   ,Forecast5=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast5"))) & vbCrLf
            s = s & "   ,Forecast6=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast6"))) & vbCrLf
            s = s & "   ,Forecast7=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast7"))) & vbCrLf
            s = s & "   ,Forecast8=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast8"))) & vbCrLf
            s = s & "   ,Forecast9=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast9"))) & vbCrLf
            s = s & "   ,Forecast10=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast10"))) & vbCrLf
            s = s & "   ,Forecast11=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast11"))) & vbCrLf
            s = s & "   ,Forecast12=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Forecast12"))) & vbCrLf
            s = s & "   ,PartNumber=" & DbQuote(Str, .TextMatrix(i, .ColIndex("SKU"))) & vbCrLf
            s = s & "   ,TaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "WHERE ISNULL(Community,'')=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Community"))) & vbCrLf
            s = s & "  AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, .TextMatrix(i, .ColIndex("CommunityPhase"))) & vbCrLf
            s = s & "  AND ISNULL(Assembly,'')=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf
            s = s & "  AND ISNULL(Model,'')=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Model"))) & vbCrLf
            s = s & "  AND Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
            s = s & "  AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
            s = s & "  AND Vendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Vendor"))) & vbCrLf
            If .TextMatrix(i, .ColIndex("PriceLevel")) = "Corporate (any Division)" Then
                s = s & "  AND DivisionID = 0" & vbCrLf
            Else
                s = s & "  AND DivisionID = " & HFApp.DivisionID & vbCrLf
            End If
            Call HFApp.SqlExec(s)
            
            
            .RowData(i) = ""
        End If
    Next
    End With
    SaveData = True
    mDirty = False
    Screen.MousePointer = vbDefault

End Function

Private Function PrettyName(fieldname As String) As String
    Dim s As String
    s = fieldname
    Select Case s
        Case "CommunityPhase + ' - ' + CommunityPhaseDesc":     s = "Community Phase"
        Case "Community + ' - ' + CommunityDesc":     s = "Community"
        Case "CostCode + ' - ' + CostCodeDesc":       s = "Cost Code"
        Case "Assembly + ' - ' + AssemblyDesc":       s = "Assembly"
        Case "CostCodeDesc":     s = "Cost Code"
        Case "CommunityDesc":    s = "Community"
        Case "CommunityPhase":   s = "Phase"
        Case "VendorName":       s = "Vendor"
        Case "AssemblyDesc":     s = "Assembly"
        Case "POIndex + ' ' + case when poindex=POIndexDescription then '' else POIndexDescription end":          s = "Purchase Order"
        Case "CostCode":         s = "Cost Code"
        Case "":                 s = ""
        Case Else
            Debug.Print "Case """ & s & """:       s = """ & s & """"
            MsgBox "STOP!!!" & vbCrLf & vbCrLf & "YOU NEED TO TRANSLATE THIS", vbExclamation
    End Select
    PrettyName = s
End Function

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Dim r As Long
    If Button <> vbRightButton Then Exit Sub
    
    Cancel = True
    If gItems.MouseRow = 0 Then
        Call FMain.ShowColumnMenu(gItems, , , Not gItems.ColKey(gItems.MouseCol) Like "Forecast*")
    Else
        FMain.mnuFItemsItemsSub(mcITEM_RENUMBER).Enabled = False
        FMain.mnuFItemsItemsSub(mcITEM_NEWITEM).Enabled = True
        FMain.mnuFItemsItemsSub(mcITEM_DUPLICATE).Enabled = False
        FMain.mnuFItemsItemsSub(mcITEM_DELETE).Enabled = True
        FMain.mnuFItemsItemsSub(mcITEM_PRICEGROUPS).Enabled = False
        FMain.mnuFItemsItemsSub(mcITEM_VIEWFILES).Enabled = True
        PopupMenu FMain.mnuFItemsItems
    End If

End Sub


Public Sub mnuFItemsItemsSub_Click(Index As Integer)

    Select Case Index
        Case mcITEM_DELETE
            Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)
        Case mcITEM_NEWITEM
            mTimerTask = "NewItem"
            Timer1.Enabled = True
        Case mcITEM_VIEWFILES
            mTimerTask = "Attachments"
            Timer1.Enabled = True
    End Select
End Sub



Public Sub mnuFItemsPriceGroupSub_Click(Index As Integer)
    Dim i As Long
    Dim s As String
    Dim PriceGroup As Long
    Dim PriceGroupDesc As String
    Dim rs As Recordset
    
    If Not SaveData(True) Then Exit Sub
    
    With gItems
        s = ""
        For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            s = s & " or (Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & " and Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & ")"
        Next
        s = "(" & Mid(s, 5) & ")"
    End With
    
    
    Select Case Index
        Case 0 'make group
            s = "update tblPhaseItem SET PriceLink=(SELECT ISNULL(MAX(PriceLink),0)+1 FROM tblPhaseItem) WHERE " & s & " and DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
            With gItems
                Set rs = HFApp.SqlExec("select PriceGroup,PriceGroupDesc from tblphaseitem left outer join PriceGroups on(PriceLink=PriceGroup) where DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Phase"))) & " and Item=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Item"))))
                PriceGroup = rs("PriceGroup")
                PriceGroupDesc = rs("PriceGroupDesc")
                For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    .Cell(flexcpData, i, .ColIndex("PriceLink")) = PriceGroup
                    .Cell(flexcpText, i, .ColIndex("PriceLink")) = PriceGroupDesc
                    If Val(.Cell(flexcpData, i, .ColIndex("PriceLink"))) <> 0 Then
                        .Cell(flexcpPicture, i, .ColIndex("PriceLink")) = FMain.SmallIcons.ListImages.Item("links").Picture
                    End If
                Next
            End With
        
        Case 1 'join group
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Pricing Group", "select PriceGroup, PriceGroupDesc from PriceGroups", , , , , "PriceGroup") Then
                PriceGroup = FPickList.SelectedItem("PriceGroup")
                PriceGroupDesc = FPickList.SelectedItem("PriceGroupDesc")
                s = "update tblPhaseItem SET PriceLink=" & DbQuote(Num, PriceGroup) & " WHERE " & s & " and DivisionID = " & HFApp.DivisionID
                Call HFApp.SqlExec(s)
                With gItems
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        .Cell(flexcpData, i, .ColIndex("PriceLink")) = PriceGroup
                        .Cell(flexcpText, i, .ColIndex("PriceLink")) = PriceGroupDesc
                        If Val(.Cell(flexcpData, i, .ColIndex("PriceLink"))) <> 0 Then
                            .Cell(flexcpPicture, i, .ColIndex("PriceLink")) = FMain.SmallIcons.ListImages.Item("links").Picture
                        End If
                    Next
                End With
            End If
        
        Case 2 'leave group
            s = "update tblPhaseItem SET PriceLink=0 WHERE " & s & " and DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
            With gItems
                For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    .Cell(flexcpData, i, .ColIndex("PriceLink")) = 0
                    .Cell(flexcpText, i, .ColIndex("PriceLink")) = ""
                    .Cell(flexcpPicture, i, .ColIndex("PriceLink")) = Nothing
                Next
            End With
        
    End Select
    
End Sub


Private Sub SetDecimalPlaces(Decimals As Long)
    Dim format  As String
        
    mDecimals = Decimals
    
    If mDecimals < 0 Then mDecimals = -1
    If mDecimals = -1 Then
        format = ""
    Else
        format = String(mDecimals, "0")
        If format = "" Then
            format = "0"
        Else
            format = "0." & format
        End If
    End If
    
    With gItems
        .ColFormat(.ColIndex("Last1")) = format
        .ColFormat(.ColIndex("Last2")) = format
        .ColFormat(.ColIndex("Last3")) = format
        .ColFormat(.ColIndex("Next1")) = format
        .ColFormat(.ColIndex("Next2")) = format
        .ColFormat(.ColIndex("Current")) = format
        .ColFormat(.ColIndex("Forecast1")) = format
        .ColFormat(.ColIndex("Forecast2")) = format
        .ColFormat(.ColIndex("Forecast3")) = format
        .ColFormat(.ColIndex("Forecast4")) = format
        .ColFormat(.ColIndex("Forecast5")) = format
        .ColFormat(.ColIndex("Forecast6")) = format
        .ColFormat(.ColIndex("Forecast7")) = format
        .ColFormat(.ColIndex("Forecast8")) = format
        .ColFormat(.ColIndex("Forecast9")) = format
        .ColFormat(.ColIndex("Forecast10")) = format
        .ColFormat(.ColIndex("Forecast11")) = format
        .ColFormat(.ColIndex("Forecast12")) = format
    End With
    
End Sub

Private Sub LoadCustomDescriptions()
    With gItems
    .TextMatrix(0, .ColIndex("Community")) = FMain.CD_Community
    .TextMatrix(0, .ColIndex("CommunityDesc")) = FMain.CD_Community & " Description"
    .TextMatrix(0, .ColIndex("CommunityPhase")) = FMain.CD_Community & " Phase"
    End With
    
End Sub


