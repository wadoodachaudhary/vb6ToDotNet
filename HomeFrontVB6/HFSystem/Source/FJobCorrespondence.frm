VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FJobCorrespondence 
   Caption         =   "Correspondence Log"
   ClientHeight    =   10665
   ClientLeft      =   1200
   ClientTop       =   1035
   ClientWidth     =   15720
   Icon            =   "FJobCorrespondence.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   15720
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   1058
      ButtonWidth     =   953
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   13080
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   62
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":000C
               Key             =   "SaveAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":08E6
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":11C0
               Key             =   "High"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1A9A
               Key             =   "Low"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":2374
               Key             =   "AssemblyCosts"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":2C4E
               Key             =   "MassChange"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":3528
               Key             =   "FieldPOs"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":3E02
               Key             =   "NewRFQ"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":46DC
               Key             =   "CreateJob"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":4FB6
               Key             =   "takeoffsettings"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":5890
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":616A
               Key             =   "DecreaseDecimals"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":6A44
               Key             =   "IncreaseDecimals"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":731E
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":7BF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":84D2
               Key             =   "UpdatePrices"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":8DAC
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":9686
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":9F60
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":A83A
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":B114
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":B9EE
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":C2C8
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":CBA2
               Key             =   "SendPOs"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":D47C
               Key             =   "SendRFQs"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":DD56
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":E630
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":EF0A
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":F7E4
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":100BE
               Key             =   "TakeoffItemChart"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":10998
               Key             =   "New"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":11272
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":11B4C
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":12426
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":12D00
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":135DA
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":13EB4
               Key             =   "AddPricelist"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1478E
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":15068
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":15942
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1621C
               Key             =   "PricelistExport"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":16AF6
               Key             =   "PricelistImport"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":173D0
               Key             =   "NewPricelist"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":17CAA
               Key             =   "View"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":18584
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":18E5E
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":19738
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1A012
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1A8EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1B1C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1BAA0
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1C37A
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1CC54
               Key             =   "Models"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1D52E
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1DE08
               Key             =   "Generate"
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1E6E2
               Key             =   "Timberline"
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1EFBC
               Key             =   "MBImport"
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":1F896
               Key             =   "EEEstimating"
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":20170
               Key             =   "EEExport"
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":20A4A
               Key             =   "EEImport"
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":21324
               Key             =   "MBExport"
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":21BFE
               Key             =   "MasterBuilder"
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer1 
         Left            =   14190
         Top             =   60
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gMessages 
      Height          =   5355
      Left            =   510
      TabIndex        =   1
      Top             =   1200
      Width           =   10905
      _cx             =   19235
      _cy             =   9446
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FJobCorrespondence.frx":228D8
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
      ExplorerBar     =   3
      PicturesOver    =   0   'False
      FillStyle       =   1
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
      Begin MSComctlLib.ImageList SmallIcons 
         Left            =   690
         Top             =   1530
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":22AB5
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":2304F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":235E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":23B83
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":2411D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJobCorrespondence.frx":246B7
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FJobCorrespondence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FJobCorrespondence::"

Private mDirty As Boolean

Private mJob As String


Public Sub ShowForm(Job As String)
    mJob = Job
    Call Me.Show(vbModal)
End Sub

Private Function PickContacts() As String
    Dim i As Long
    Dim s As String
    
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Contact", "select * from addressbook", , , , , , True) Then
        s = ""
        For i = 1 To FPickList.SelectedItems
            s = s & "; " & FPickList.SelectedItem("name", i)
        Next
        PickContacts = Mid(s, 3)
    End If

End Function

Private Sub Form_Load()
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gMessages)
    
    
    gMessages.ColImageList(gMessages.ColIndex("Priority")) = Me.SmallIcons.hImageList
    gMessages.ColImageList(gMessages.ColIndex("Original")) = Me.SmallIcons.hImageList
    gMessages.Cell(flexcpPicture, 0, gMessages.ColIndex("priority")) = SmallIcons.ListImages(6).Picture
    
    
    Call LoadJob(mJob)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gMessages)
End Sub

Private Sub gMessages_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gMessages
        If .TextMatrix(Row, .ColIndex("MessageType")) = "Email" Then
            Cancel = True
            Exit Sub
        End If
        
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "Attachments", "Size", "ID", "Original", "Priority"
                Cancel = True
                
            Case "MessageType"
                .ComboList = "Mail|Fax|Phone|Other"
                
            Case "SentDate", "ReceivedDate", "Sender", "Recipients", "RecipientsCC", "RecipientsBCC"
                .ComboList = "|..."
        End Select
        
    End With
End Sub

Private Sub gMessages_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gMessages
        Select Case .ColKey(Col)
            Case "SentDate", "ReceivedDate"
                Call DCalendar.Popup(gMessages, .RowPos(.Row) + .RowHeight(.Row), .colPos(.Col))
                .RowData(Row) = "DIRTY"
                Dirty = True
            
            Case "Sender", "Recipients", "RecipientsCC", "RecipientsBCC"
                s = PickContacts()
                If s <> "" Then
                    .Text = s
                    .RowData(Row) = "DIRTY"
                    Dirty = True
                End If
            
        End Select
    End With
            
End Sub

Private Sub gMessages_Click()
    With gMessages
        If .ColKey(.Col) = "Priority" And .TextMatrix(.Row, .ColIndex("MessageType")) <> "Email" Then
            Select Case .ValueMatrix(.Row, .Col)
                Case 0:  .TextMatrix(.Row, .Col) = "1"
                Case 1:  .TextMatrix(.Row, .Col) = "2"
                Case 2:  .TextMatrix(.Row, .Col) = "0"
            End Select
            .RowData(.Row) = "DIRTY"
            Dirty = True
        End If
    End With
End Sub

Private Sub gMessages_DblClick()
    Dim s As String
    Dim ID As String
    
    With gMessages
        If .ColKey(.MouseCol) = "Original" And .TextMatrix(.Row, .ColIndex("MessageType")) = "Email" Then
            ID = .TextMatrix(.Row, .ColIndex("id"))
            's = TempFile("msg")
            'Call DBGetFile(s, , HFApp.Databases(dbHomeFront), "messages WHERE id=" & DbQuote(Str, ID), "Original")
            'Call ShellFile(Me.hwnd, s)
            
            
            s = "c:\demo\files\" & ID & ".msg"
            Call ShellFile(Me.hwnd, s)
        End If
    End With
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next
        
    gMessages.Move 0, Toolbar.Height - 15, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height + 15

End Sub


Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    
    Dim i As Long
    Dim s As String
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    
    Screen.MousePointer = vbHourglass
    With gMessages
        
        For i = .Rows - 2 To 1 Step -1
            If .RowHidden(i) Then
                s = "delete from messages where id=" & DbQuote(Str, .TextMatrix(i, .ColIndex("id")))
                Call HFApp.SqlExec(s)
                Call .RemoveItem(i)
            End If
        Next
        
        
        For i = 1 To .Rows - 2
            If .RowData(i) = "DIRTY" Then
                If .TextMatrix(i, .ColIndex("id")) = "" Then
                    .TextMatrix(i, .ColIndex("id")) = CreateGUID()
                    s = ""
                    s = s & "insert into messages(job,rfiid,id,priority,messagetype,sender,recipients,recipientscc,recipientsbcc,subject,body,sentdate,receiveddate)" & vbCrLf
                    s = s & "values(" & DbQuote(Str, mJob) & vbCrLf
                    s = s & "      ," & DbQuote(Num, 0) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("ID"))) & vbCrLf
                    s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Priority"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("MessageType"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Sender"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Recipients"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsCC"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsBCC"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Subject"))) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Body"))) & vbCrLf
                    s = s & "      ," & DbQuote(Date, .TextMatrix(i, .ColIndex("SentDate"))) & vbCrLf
                    s = s & "      ," & DbQuote(Date, .TextMatrix(i, .ColIndex("ReceivedDate"))) & vbCrLf
                    s = s & ")"
                Else
                    s = ""
                    s = s & "update messages" & vbCrLf
                    s = s & "set priority=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Priority"))) & vbCrLf
                    s = s & "   ,messagetype=" & DbQuote(Str, .TextMatrix(i, .ColIndex("MessageType"))) & vbCrLf
                    s = s & "   ,sender=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Sender"))) & vbCrLf
                    s = s & "   ,recipients=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Recipients"))) & vbCrLf
                    s = s & "   ,recipientscc=" & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsCC"))) & vbCrLf
                    s = s & "   ,recipientsbcc=" & DbQuote(Str, .TextMatrix(i, .ColIndex("RecipientsBCC"))) & vbCrLf
                    s = s & "   ,subject=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Subject"))) & vbCrLf
                    s = s & "   ,body=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Body"))) & vbCrLf
                    s = s & "   ,sentdate=" & DbQuote(Date, .TextMatrix(i, .ColIndex("SentDate"))) & vbCrLf
                    s = s & "   ,receiveddate=" & DbQuote(Date, .TextMatrix(i, .ColIndex("ReceivedDate"))) & vbCrLf
                    s = s & "where id=" & DbQuote(Str, .TextMatrix(i, .ColIndex("id")))
                End If
                Call HFApp.SqlExec(s)
                .RowData(i) = ""
            End If
        Next
        
    End With
    
    
    mDirty = False
    Screen.MousePointer = vbDefault

Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function

Private Sub LoadJob(Job As String)
On Error GoTo eh
    
    Dim i  As Long
    Dim s  As String
    Dim rs As Recordset
    
    
    
    If Job = "" Then
        s = ""
        s = s & "SELECT DISTINCT j.Job_No Job" & vbCrLf
        s = s & "      ,j.Description" & vbCrLf
        s = s & "      ,j.Municipal_Address Address" & vbCrLf
        s = s & "      ,j.PM,Purchaser" & vbCrLf
        s = s & "      ,j.Estimator" & vbCrLf
        s = s & "FROM tblJobs j" & vbCrLf
        s = s & "WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "AND ISNULL(j.Inactive,0)=0 " & vbCrLf
        
        If FPickList.Choose(HFApp.Databases(dbHomefront), "Jobs", s, mJob) Then
            mJob = FPickList.SelectedItem("Job")
        Else
            Exit Sub
        End If
    Else
        mJob = Job
    End If
    
    
    'load messages
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "  from messages" & vbCrLf
    s = s & " where job=" & DbQuote(Str, mJob) & vbCrLf
    s = s & "   and isnull(rfiid,0)=0"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gMessages
        .Rows = 1
        i = 0
        While Not rs.EOF
            .AddItem ""
            i = i + 1
            .TextMatrix(i, .ColIndex("ID")) = "" & rs("ID")
            .TextMatrix(i, .ColIndex("Original")) = IIf("" & rs("MessageType") = "Email", "4", "")
            .TextMatrix(i, .ColIndex("Priority")) = "" & rs("Priority")
            .TextMatrix(i, .ColIndex("Attachments")) = IIf(Val("" & rs("Attachments")) = 0, "", "3")
            .TextMatrix(i, .ColIndex("MessageType")) = "" & rs("MessageType")
            .TextMatrix(i, .ColIndex("Sender")) = "" & rs("Sender")
            .TextMatrix(i, .ColIndex("Recipients")) = "" & rs("Recipients")
            .TextMatrix(i, .ColIndex("RecipientsCC")) = "" & rs("RecipientsCC")
            .TextMatrix(i, .ColIndex("RecipientsBCC")) = "" & rs("RecipientsBCC")
            .TextMatrix(i, .ColIndex("Subject")) = "" & rs("Subject")
            .TextMatrix(i, .ColIndex("Body")) = "" & rs("Body")
            .TextMatrix(i, .ColIndex("Size")) = IIf(Val("" & rs("Size")) = 0, "", "" & rs("Size") & " KB")
            .TextMatrix(i, .ColIndex("SentDate")) = "" & rs("SentDate")
            .TextMatrix(i, .ColIndex("ReceivedDate")) = "" & rs("ReceivedDate")
            rs.MoveNext
        Wend
        .AddItem ""
    End With
    
    Dirty = False
Exit Sub
eh: Call errHandler(SRCFILE & "LoadJob")
End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Open":    If Not SaveData(True) Then Exit Sub
                        Call LoadJob("")
        Case "Save":    Call SaveData(False)
    End Select
End Sub


Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
End Property



Private Sub gMessages_BeforeSort(ByVal Col As Long, Order As Integer)
On Error Resume Next
    gMessages.RemoveItem gMessages.Rows - 1
End Sub

Private Sub gMessages_AfterSort(ByVal Col As Long, Order As Integer)
On Error Resume Next
    gMessages.AddItem ""
End Sub


Private Sub gMessages_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    Dim i As Long
    With gMessages
        If Row <> .Rows - 1 Then Exit Sub
        
        i = Row
        .TextMatrix(i, .ColIndex("Priority")) = "1"
        .TextMatrix(i, .ColIndex("MessageType")) = "Other"
        .TextMatrix(i, .ColIndex("ReceivedDate")) = Format(Now(), "medium date")
    
        'add new last row
        .AddItem ""
    
    
    End With
End Sub


Private Sub gMessages_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh

    Dim s  As String
    Dim q  As String
    
    With gMessages
        s = .EditText
        Select Case .ColKey(Col)
                
            Case "SentDate", "ReceivedDate"
                If IsDate(s) Then
                    s = Format(s, "Medium Date")
                Else
                    Cancel = True
                End If
         End Select
        .EditText = s
        
        If Not Cancel Then
            Dirty = True
            .RowData(Row) = "DIRTY"
        End If
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "gMessages_ValidateEdit")
End Sub








