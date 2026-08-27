VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FVisaImporter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payables Desk"
   ClientHeight    =   5220
   ClientLeft      =   5805
   ClientTop       =   2040
   ClientWidth     =   6510
   Icon            =   "FVisaImporter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2085
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   4665
      _cx             =   8229
      _cy             =   3678
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6510
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4635
      Width           =   6510
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5340
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4140
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   0
         X2              =   26540
         Y1              =   0
         Y2              =   0
      End
   End
End
Attribute VB_Name = "FVisaImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Const SRCFILE = "FVisaImporter::"

Private mSessionID As String
Private mFilename  As String

Public Function ShowForm()
    mWMS = False
    If VBGetOpenFileName(mFilename, , , , , True, "Comma Separated Values Files (*.csv)|*.csv", , , , , FMain.hwnd) Then
        CurrentFrame = 0
    Else
        Unload Me
    End If
    Me.Show vbModal
End Function

Private Function TextMatrix(r As Long, ColKey As String) As String
On Error Resume Next
    TextMatrix = gData.TextMatrix(r, gData.ColIndex(ColKey))
End Function

Private Function ValidateTable() As Boolean
'    Dim s As String
'    Dim rs As Recordset
'
'    s = ""
'    s = s & "UPDATE ImportedCosts" & vbCrLf
'    s = s & "   SET UnknownVendor= case when v.vendor_id is null then 1 else 0 end" & vbCrLf
'    s = s & "      ,UnknownCommunity= case when c.community<>'' and l.area is null then 1 else 0 end" & vbCrLf
'    s = s & "      ,UnknownAssembly= case when c.Assembly<>'' AND a.Assembly IS NULL then 1 else 0 end" & vbCrLf
'    s = s & "      ,UnknownItem= case when i.Phase is null then 1 else 0 end" & vbCrLf
'    s = s & "      ,UnknownUOM= case when upper(rtrim(ltrim(c.OrderUOM)))<>upper(rtrim(ltrim(i.OrderUOM))) then 1 else 0 end" & vbCrLf
'    s = s & "      ,ExpectedUOM= case when upper(rtrim(ltrim(c.OrderUOM)))<>upper(rtrim(ltrim(i.OrderUOM))) then i.OrderUOM else '' end" & vbCrLf
'    s = s & "      ,UnknownExpiryDate = case when ((Next_Cost1<>0 AND Next_Effective1 IS NULL) OR (Next_Cost2<>0 AND Next_Effective2 IS NULL)) then 1 else 0 end" & vbCrLf
'    s = s & "  FROM ImportedCosts c" & vbCrLf
'    s = s & "       LEFT OUTER JOIN tblVendors v ON (c.Vendor=v.Vendor_ID)" & vbCrLf
'    s = s & "       LEFT OUTER JOIN tblLocality l ON (c.Community=l.Area)" & vbCrLf
'    s = s & "       LEFT OUTER JOIN DistinctAssemblies a ON (c.Assembly=a.Assembly)" & vbCrLf
'    s = s & "       LEFT OUTER JOIN tblPhaseItem i ON (c.Phase=i.Phase AND c.Item=i.Item)" & vbCrLf
'    s = s & " WHERE SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
'    Set rs = HFApp.SqlExec(s)
'
'
'    s = ""
'    s = s & "SELECT DISTINCT 'The file contains unrecognized vendor(s)'" & vbCrLf
'    s = s & "  FROM ImportedCosts c" & vbCrLf
'    s = s & " WHERE UnknownVendor=1" & vbCrLf
'    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
'    s = s & "UNION ALL" & vbCrLf
'    s = s & "SELECT DISTINCT 'The file contains unrecognized community(s)'" & vbCrLf
'    s = s & "  FROM ImportedCosts c" & vbCrLf
'    s = s & " WHERE UnknownCommunity=1" & vbCrLf
'    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
'    s = s & "UNION ALL" & vbCrLf
'    s = s & "SELECT DISTINCT 'The file contains unrecognized assembly(s)'" & vbCrLf
'    s = s & "  FROM ImportedCosts c" & vbCrLf
'    s = s & " WHERE UnknownAssembly=1" & vbCrLf
'    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
'    s = s & "UNION ALL" & vbCrLf
'    s = s & "SELECT DISTINCT 'The file contains unrecognized item(s)'" & vbCrLf
'    s = s & "  FROM ImportedCosts c" & vbCrLf
'    s = s & " WHERE UnknownItem=1" & vbCrLf
'    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
'    s = s & "UNION ALL" & vbCrLf
'    s = s & "SELECT DISTINCT 'The file contains conflicting units'" & vbCrLf
'    s = s & "  FROM ImportedCosts c" & vbCrLf
'    s = s & " WHERE UnknownUOM=1" & vbCrLf
'    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
'    s = s & "UNION ALL" & vbCrLf
'    s = s & "SELECT DISTINCT 'The file contains future pricing with no effective date(s)'" & vbCrLf
'    s = s & "  FROM ImportedCosts c" & vbCrLf
'    s = s & " WHERE UnknownExpiryDate=1" & vbCrLf
'    s = s & "   AND SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
'    Set rs = HFApp.SqlExec(s)
'    s = ""
'    While Not rs.EOF
'        s = s & "" & rs(0) & vbCrLf
'        rs.MoveNext
'    Wend
'
'    If s <> "" Then
'        cmdNav(3).Enabled = False
'        imgError.Visible = True
'        imgOK.Visible = False
'        lblDescription.Caption = "The pricelist import wizard is not able to import this data."
'        lblErrors.Caption = s
'        ValidateTable = False
'
'        frmOptions(0).Visible = True
'        frmOptions(1).Top = frmOptions(0).Top + frmOptions(0).Height
'    Else
'        imgError.Visible = False
'        imgOK.Visible = True
'        ValidateTable = True
'        lblDescription.Caption = "The pricelist import wizard is ready to process the file."
'        lblErrors.Caption = "View the reports listed below to see how this will change your vendor pricing" & vbCrLf & _
'                            "database. If you are satisfied with the changes and would like to commit them" & vbCrLf & _
'                            "to the database click ""Finish"". If you are not satisfied click ""Cancel"" to" & vbCrLf & _
'                            "exit the wizard without updating your database."
'
'        frmOptions(1).Visible = True
'        frmOptions(1).Top = frmOptions(0).Top
'
'    End If
'
End Function

Private Property Get CurrentFrame() As Long
    Dim i As Long
    For i = 0 To WizFrame.UBound
        If WizFrame(i).Visible = True Then
            CurrentFrame = i
            Exit Property
        End If
    Next
    CurrentFrame = WizFrame.LBound
End Property

Private Property Let CurrentFrame(RHS As Long)
On Error GoTo eh
    Dim i As Long
    Dim fileinfo As New ClsFileInfo
    Dim s As String
    Dim r As Long
    Dim Vendor As String
    Dim VendorDesc As String
    
    Screen.MousePointer = vbHourglass
    
    
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
    
    
    With gData
    Select Case RHS
    
        Case 0 ' show file
            If WizFrame(RHS).Tag = "" Then
                WizFrame(RHS).Tag = "LOADED"
                If mWMS Then
                    mFilename = HFApp.Options(WMSDatabase)
                    lblFilename.Caption = "File name: " & mFilename
                    fileinfo.FullPathName = mFilename
                    lblFileDate.Caption = "Last Modified: " & Format(fileinfo.ModifyTime, "long date")
                    Call LoadWMSPrices
                Else
                    lblFilename.Caption = "File name: " & mFilename
                    fileinfo.FullPathName = mFilename
                    lblFileDate.Caption = "Last Modified: " & Format(fileinfo.ModifyTime, "long date")
                    Call LoadExcelSheet(GetObject(mFilename), "", gData)
                    If gData.Cols > 16 Then gData.Cols = 16
                End If
            End If
            
        Case 1 ' write data then show options
            If WizFrame(RHS).Tag = "" Then
                WizFrame(RHS).Tag = "LOADED"
                If ValidateGrid() Then
                    r = 1
                    WizFrame(0).Visible = True
                    ProgressBar.Visible = True
                    WizFrame(1).Visible = False
                    While r < .Rows
                        ProgressBar.Value = (r / (.Rows - 1)) * 100
                        If .TextMatrix(r, 1) = "Vendor" Then
                            Vendor = .TextMatrix(r, 2)
                            If r < .Rows Then VendorDesc = .TextMatrix(r + 1, 2)
                            r = r + 1
                        End If
                        
                        If r < .Rows Then
                            If (Trim(.TextMatrix(r, .ColIndex("Group"))) <> "") And Trim(.TextMatrix(r, .ColIndex("Group"))) <> "Group" Then
                                s = ""
                                s = s & "INSERT INTO ImportedCosts(SessionID,FileName,Community,CommunityDesc,CommunityPhase,Vendor,VendorDesc,Assembly,AssemblyDesc,Phase,Item,ItemDesc,PartNumber,OrderUOM,Current_Cost,Next_Cost1,Next_Effective1,Next_Cost2,Next_Effective2)" & vbCrLf
                                s = s & "VALUES(" & DbQuote(Str, mSessionID, , , 50) & vbCrLf
                                s = s & "      ," & DbQuote(Str, FileTitle(mFilename), , True, 100) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Community")), , True, 10) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Community Name")), , True, 100) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase")), , True, 10) & vbCrLf
                                s = s & "      ," & DbQuote(Str, Vendor, , True, 20) & vbCrLf
                                s = s & "      ," & DbQuote(Str, VendorDesc, , True, 100) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly")), , True, 20) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly Description")), , True, 100) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Group")), , True, 20) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item")), , True, 10) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item Description")), , True, 100) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Part Number")), , True, 50) & vbCrLf
                                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("UOM")), , True, 4) & vbCrLf
                                s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Price"))) & vbCrLf
                                s = s & "      ," & DbQuote(Num, MYTextMatrix(r, .ColIndex("Next Price 1"))) & vbCrLf
                                s = s & "      ," & DbQuote(Date, MYTextMatrix(r, .ColIndex("Effective 1"))) & vbCrLf
                                s = s & "      ," & DbQuote(Num, MYTextMatrix(r, .ColIndex("Next Price 2"))) & vbCrLf
                                s = s & "      ," & DbQuote(Date, MYTextMatrix(r, .ColIndex("Effective 2"))) & vbCrLf
                                s = s & ")"
                                HFApp.SqlExec s
                            End If
                        End If
                        r = r + 1
                        
                    Wend
                    If s <> "" Then HFApp.SqlExec s
                    WizFrame(0).Visible = False
                    ProgressBar.Visible = False
                    WizFrame(1).Visible = True
                    Call ValidateTable
                    Screen.MousePointer = vbDefault
                End If
            End If
                
    End Select
    End With
    
    
    
    
    Screen.MousePointer = vbDefault
    Exit Property
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call ErrHandler(SRCFILE & "CurrentFrame", s)
    End If
End Property


Private Sub chkIgnoreErrors_Click()
    frmOptions(1).Visible = chkIgnoreErrors.Value = vbChecked
    cmdNav(3).Enabled = chkIgnoreErrors.Value = vbChecked And Trim(txtRejectFile.Text) <> ""
End Sub

Private Sub cmdChooseFolder_Click(Index As Integer)
    Dim s As String
    txtRejectFile.SetFocus
    s = txtRejectFile.Text
    If VBGetOpenFileName(s, , False, , , True, "Excel Files (*.xls)|*.xls", , , , "xls", Me.hwnd) Then
        txtRejectFile.Text = s
        cmdNav(3).Enabled = chkIgnoreErrors.Value = vbChecked And Trim(txtRejectFile.Text) <> ""
    End If
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3: If SaveData Then Unload Me
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).BorderStyle = 0
        WizFrame(i).Move 0, WizHead1.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead1.Height - WizFoot.Height
    Next
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
    gData.Move margin, gData.Top, WizFrame(0).Width - 2 * margin, WizFrame(0).Height - gData.Top - margin
    ProgressBar.Move gData.left, gData.Top + gData.Height - ProgressBar.Height, gData.Width
End Sub

Private Sub Form_Load()
    mSessionID = MachineName & App.hInstance
    Call IniGetForm(Me)
    
    Call LoadCostTypes(cboImportColumn, , True)
    Call cboImportColumn.InsertItem("Current", 0)
    Call cboImportColumn.InsertItem("Next 1", 1)
    Call cboImportColumn.InsertItem("Next 2", 2)
    cboImportColumn.ListIndex = HFApp.Options(PricelistImportDestination)
    chkIgnoreZeroValues.Value = IIf(HFApp.Options(PricelistImportSkipZeros), vbChecked, vbUnchecked)
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HFApp.Databases(dbHomefront).CommandTimeout = 0
    Call HFApp.SqlExec("DELETE FROM ImportedCosts WHERE SessionID=" & DbQuote(Str, mSessionID))
    HFApp.Options.Value(PricelistImportSkipZeros) = chkIgnoreZeroValues.Value = vbChecked
    HFApp.Options.Value(PricelistImportDestination) = cboImportColumn.ListIndex

    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub


Private Sub lblShowReport_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Call FRptViewer.ShowReport(HFApp.SystemFolder & "System\Reports\Estimating\PriceImportSummary.rpt", True, False, True, "SessionID", mSessionID, "ImportDestination", cboImportColumn.ListIndex)
        Case 1: Call FRptViewer.ShowReport(HFApp.SystemFolder & "System\Reports\Estimating\PriceImportComparison.rpt", True, False, True, "SessionID", mSessionID, "ImportDestination", cboImportColumn.ListIndex)
    End Select
End Sub

Private Function SaveData() As Boolean
On Error GoTo eh
    Dim s As String
    
    
    
    s = ""
    s = s & "UPDATE tblVendorCost " & vbCrLf
    Select Case cboImportColumn.ListIndex
        Case 0
            If chkIgnoreZeroValues.Value = vbChecked Then
                s = s & "   SET Current_Cost=   Case when ic.Current_Cost <> 0                                  then ic.Current_Cost    else vc.Current_Cost    end" & vbCrLf
                s = s & "      ,Next_Cost1=     Case when ic.Next_Cost1 <> 0 And Not ic.Next_Effective1 Is Null then ic.Next_Cost1      else vc.Next_Cost1      end" & vbCrLf
                s = s & "      ,Next_Effective1=Case when ic.Next_Cost1 <> 0 And Not ic.Next_Effective1 Is Null then ic.Next_Effective1 else vc.Next_Effective1 end" & vbCrLf
                s = s & "      ,Next_Cost2=     Case when ic.Next_Cost2 <> 0 And Not ic.Next_Effective2 Is Null then ic.Next_Cost2      else vc.Next_Cost2      end" & vbCrLf
                s = s & "      ,Next_Effective2=Case when ic.Next_Cost2 <> 0 And Not ic.Next_Effective2 Is Null then ic.Next_Effective2 else vc.Next_Effective2 end" & vbCrLf
            Else
                s = s & "   SET Current_Cost=ic.Current_Cost" & vbCrLf
                s = s & "      ,Next_Cost1=     Case when Not ic.Next_Effective1 Is Null then ic.Next_Cost1      else vc.Next_Cost1      end" & vbCrLf
                s = s & "      ,Next_Effective1=Case when Not ic.Next_Effective1 Is Null then ic.Next_Effective1 else vc.Next_Effective1 end" & vbCrLf
                s = s & "      ,Next_Cost2=     Case when Not ic.Next_Effective2 Is Null then ic.Next_Cost2      else vc.Next_Cost2      end" & vbCrLf
                s = s & "      ,Next_Effective2=Case when Not ic.Next_Effective2 Is Null then ic.Next_Effective2 else vc.Next_Effective2 end" & vbCrLf
            End If
            
        Case 1
            If Not IsDate(txtEffectiveDate) Then
                MsgBox "You must specify an effective date", vbExclamation
                Exit Function
            End If
            If chkIgnoreZeroValues.Value = vbChecked Then
                s = s & "   SET Next_Cost1=Case when ic.Current_Cost <> 0 then ic.Current_Cost else vc.Next_Cost1 end" & vbCrLf
            Else
                s = s & "   SET Next_Cost1=ic.Current_Cost" & vbCrLf
            End If
            s = s & "      ,Next_Effective1=" & DbQuote(Date, txtEffectiveDate) & vbCrLf
            
        Case 2
            If Not IsDate(txtEffectiveDate) Then
                MsgBox "You must specify an effective date", vbExclamation
                Exit Function
            End If
            If chkIgnoreZeroValues.Value = vbChecked Then
                s = s & "   SET Next_Cost2=Case when ic.Current_Cost <> 0 then ic.Current_Cost else vc.Next_Cost2 end" & vbCrLf
            Else
                s = s & "   SET Next_Cost2=ic.Current_Cost" & vbCrLf
            End If
            s = s & "      ,Next_Effective2=" & DbQuote(Date, txtEffectiveDate) & vbCrLf
            
        Case Else
            If chkIgnoreZeroValues.Value = vbChecked Then
                s = s & "   SET Forecast" & cboImportColumn.ListIndex - 2 & "=case when ic.Current_Cost<>0 then ic.Current_Cost else vc.Current_Cost end" & vbCrLf
            Else
                s = s & "   SET Forecast" & cboImportColumn.ListIndex - 2 & "=ic.Current_Cost" & vbCrLf
            End If
            
    End Select
    s = s & "      ,PartNumber=ic.PartNumber" & vbCrLf
    s = s & "  FROM ImportedCosts ic " & vbCrLf
    s = s & "  JOIN tblVendorCost vc " & vbCrLf
    s = s & "    ON(ic.Community=vc.Community AND" & vbCrLf
    s = s & "       ic.CommunityPhase=vc.CommunityPhase AND" & vbCrLf
    s = s & "       ic.Assembly=vc.Assembly AND" & vbCrLf
    s = s & "       ic.Phase=vc.Phase AND" & vbCrLf
    s = s & "       ic.Item=vc.Item AND" & vbCrLf
    s = s & "       ic.Vendor=vc.Vendor AND" & vbCrLf
    s = s & "       ic.UnknownVendor=0 AND" & vbCrLf
    s = s & "       ic.UnknownCommunity=0 AND" & vbCrLf
    s = s & "       ic.UnknownAssembly=0 AND" & vbCrLf
    s = s & "       ic.UnknownItem=0 AND" & vbCrLf
    s = s & "       ic.UnknownUOM=0 AND" & vbCrLf
    s = s & "       ic.UnknownExpiryDate=0 AND" & vbCrLf
    s = s & "       ic.SessionID=" & DbQuote(Str, mSessionID) & ")" & vbCrLf
    Call HFApp.SqlExec(s)
    
    
    
    
    s = ""
    s = s & "INSERT INTO tblVendorCost(Community,CommunityPhase,Assembly,Phase,Item,PriceLink,Vendor" & vbCrLf
    s = s & "      ,Current_Cost" & vbCrLf
    s = s & "      ,Next_Cost1,Next_Effective1" & vbCrLf
    s = s & "      ,Next_Cost2,Next_Effective2" & vbCrLf
    s = s & "      ,Forecast1,Forecast2,Forecast3,Forecast4,Forecast5,Forecast6,Forecast7,Forecast8,Forecast9,Forecast10,Forecast11,Forecast12)" & vbCrLf
    s = s & "SELECT ic.Community,ic.CommunityPhase,ic.Assembly,ic.Phase,ic.Item,case when ISNULL(ic.Assembly,'')='' then 0 else pi.PriceLink end,ic.Vendor" & vbCrLf
    s = s & "      ,MAX(ic.Current_Cost)" & vbCrLf
    s = s & "      ,case when MAX(ic.Next_Cost1)<>0 then MAX(ic.Next_Cost1) else MAX(ic.Current_Cost) end" & vbCrLf
    s = s & "      ,case when MAX(ic.Next_Cost1)<>0 then MIN(ic.Next_Effective1) else NULL end" & vbCrLf
    s = s & "      ,case when MAX(ic.Next_Cost2)<>0 then MAX(ic.Next_Cost2) else MAX(ic.Current_Cost) end" & vbCrLf
    s = s & "      ,case when MAX(ic.Next_Cost2)<>0 then MIN(ic.Next_Effective2) else NULL end" & vbCrLf
    s = s & "      ,MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost),MAX(ic.Current_Cost)" & vbCrLf
    s = s & "  FROM ImportedCosts ic " & vbCrLf
    s = s & "  LEFT OUTER JOIN tblPhaseItem pi on(ic.phase=pi.phase and ic.item=pi.item)" & vbCrLf
    s = s & "  LEFT OUTER JOIN tblVendorCost vc " & vbCrLf
    s = s & "    ON(ic.Community=vc.Community AND" & vbCrLf
    s = s & "       ic.CommunityPhase=vc.CommunityPhase AND" & vbCrLf
    s = s & "       ic.Assembly=vc.Assembly AND" & vbCrLf
    s = s & "       ic.Phase=vc.Phase AND" & vbCrLf
    s = s & "       ic.Item=vc.Item AND" & vbCrLf
    s = s & "       ic.Vendor=vc.Vendor AND" & vbCrLf
    s = s & "       ic.SessionID=" & DbQuote(Str, mSessionID) & ")" & vbCrLf
    s = s & " WHERE vc.Vendor IS NULL" & vbCrLf
    s = s & "   AND ic.SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "   AND ic.UnknownVendor=0" & vbCrLf
    s = s & "   AND ic.UnknownCommunity=0" & vbCrLf
    s = s & "   AND ic.UnknownAssembly=0" & vbCrLf
    s = s & "   AND ic.UnknownItem=0" & vbCrLf
    s = s & "   AND ic.UnknownUOM=0" & vbCrLf
    s = s & "   AND ic.UnknownExpiryDate=0" & vbCrLf
    s = s & "GROUP BY ic.Community,ic.CommunityPhase,ic.Assembly,pi.PriceLink,ic.Phase,ic.Item,ic.Vendor" & vbCrLf
    Call HFApp.SqlExec(s)
    
    
    SaveData = True
    If Me.chkIgnoreErrors.Value = vbChecked Then Call WriteRejectFile
    Exit Function
    
eh: Call ErrHandler(SRCFILE & "SaveData", s)

End Function

Private Sub cmdEffectiveDate_Click()
On Error Resume Next
    Call txtEffectiveDate_KeyDown(vbKeyF4, 0)
End Sub
Private Sub txtEffectiveDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo eh
    If KeyCode = vbKeyF4 And Shift = 0 Then Call DCalendar.Popup(txtEffectiveDate)
Exit Sub
eh: Call ErrHandler(SRCFILE & "txtEffectiveDate_KeyDown")
End Sub
Private Sub txtEffectiveDate_Validate(Cancel As Boolean)
On Error GoTo eh
    If IsDate(txtEffectiveDate) Or txtEffectiveDate = "" Then
        txtEffectiveDate.Text = Format(txtEffectiveDate, HFApp.Options(DateFormat))
    Else
        MsgBox "Not a valid date", vbExclamation
        Call SelectAll(txtEffectiveDate)
        Cancel = True
    End If
    Exit Sub
eh: Call ErrHandler(SRCFILE & "txtEffectiveDate_Validate")
End Sub
Private Sub cboImportColumn_Click()
    cmdEffectiveDate.Visible = cboImportColumn.ListIndex = 1 Or cboImportColumn.ListIndex = 2
    lblEffectiveDate.Visible = cmdEffectiveDate.Visible
    txtEffectiveDate.Visible = cmdEffectiveDate.Visible
End Sub

Private Sub LoadWMSPrices()
    Dim WMS As New Connection
    Dim rs As Recordset
    Dim s As String
    Dim i As Long
    Dim Vendor As String
    
    
    s = "DRIVER=Firebird/InterBase(r) driver;UID=" & HFApp.Options(WMSUid) & ";PWD=" & HFApp.Options(WMSPWD) & ";DBNAME=" & HFApp.Options(WMSDatabase)
    WMS.Open s
    
    s = ""
    s = s & "SELECT " & vbCrLf
    s = s & "  SUPPLIER.SUB_AP_NUMBER  Vendor" & vbCrLf
    s = s & ", SUPPLIER.SUB_NAME       CompanyName" & vbCrLf
    s = s & ", AREA.SALES_AREA         Community" & vbCrLf
    s = s & ", AREA.AREA_COUNTY        Phase" & vbCrLf
    s = s & ", AREA.AREA_DESC          CommunityName" & vbCrLf
    s = s & ", ''                      Assembly" & vbCrLf
    s = s & ", ''                      AssemblyDescription" & vbCrLf
    s = s & ", TLPE_PHASES.PHASE_CODE  gGroup" & vbCrLf
    s = s & ", TLPE_ITEMS.ITEM_NUMBER  Item" & vbCrLf
    s = s & ", TLPE_ITEMS.ITEM_DESC    ItemDescription" & vbCrLf
    s = s & ", ''                      PartNumber" & vbCrLf
    s = s & ",TLPE_ITEMS.ORDER_UNIT    UOM" & vbCrLf
    s = s & ", COSTS.UNIT_COST         Price" & vbCrLf
    s = s & ",COSTS.NEXT_COST          NextPrice1" & vbCrLf
    s = s & ", COSTS.NEXT_COST_DUE     Effective1" & vbCrLf
    s = s & ", COSTS.NEXT_COST_2       NextPrice2" & vbCrLf
    s = s & ", COSTS.NEXT_COST_2_DUE   Effective2" & vbCrLf
    s = s & "FROM COSTS LEFT JOIN AREA ON COSTS.AREA_ID = AREA.AREA_ID" & vbCrLf
    s = s & "INNER JOIN SUPPLIER ON COSTS.SUB_NUMBER = SUPPLIER.SUB_NUMBER" & vbCrLf
    s = s & "INNER JOIN TLPE_ITEMS ON COSTS.TLPE_ITEMS_ID = TLPE_ITEMS.TLPE_ITEMS_ID" & vbCrLf
    s = s & "INNER JOIN TLPE_PHASES ON TLPE_ITEMS.TLPE_PHASES_ID = TLPE_PHASES.TLPE_PHASES_ID" & vbCrLf
    s = s & "order by 1,3,4,8,9" & vbCrLf
    
    Set rs = WMS.Execute(s)
    
    With gData
        .Rows = 1
        .Cols = 16
        .FixedCols = 1
        .FixedRows = 1
        While Not rs.EOF
            If Vendor <> "" & rs("vendor") Then
                Vendor = "" & rs("vendor")
                .AddItem ""
                .AddItem ""
                .AddItem vbTab & "Vendor" & vbTab & Vendor
                .AddItem vbTab & "Company Name" & vbTab & rs("CompanyName")
                .AddItem vbTab & "Community" & vbTab & "Phase" & vbTab & "Community Name" & vbTab & "Assembly" & vbTab & "Assembly Description" & vbTab & "Group" & vbTab & "Item" & vbTab & "Item Description" & vbTab & "Part Number" & vbTab & "UOM" & vbTab & "Price" & vbTab & "Next Price 1" & vbTab & "Effective 1" & vbTab & "Next Price 2" & vbTab & "Effective 2"
                .Refresh
                .Row = .Rows - 1
            End If
            .AddItem vbTab & rs("Community") & vbTab & _
                     rs("Phase") & vbTab & _
                     rs("CommunityName") & vbTab & _
                     rs("Assembly") & vbTab & _
                     rs("AssemblyDescription") & vbTab & _
                     rs("gGroup") & vbTab & _
                     rs("Item") & vbTab & _
                     rs("ItemDescription") & vbTab & _
                     rs("PartNumber") & vbTab & _
                     rs("UOM") & vbTab & _
                     rs("Price") & vbTab & _
                     rs("NextPrice1") & vbTab & _
                     rs("Effective1") & vbTab & _
                     rs("NextPrice2") & vbTab & _
                     rs("Effective2")
            rs.MoveNext
        Wend
        
        
        'write row and column numbers into grid.headings
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
        For i = 1 To .Cols - 1
            .TextMatrix(0, i) = FormatExcelColRef(i)
        Next
        .ColWidth(0) = 360
        .ColAlignment(0) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        Call .AutoSize(0, .Cols - 1)
    End With

End Sub

Private Function MYTextMatrix(Row As Long, Col As Long) As String
On Error Resume Next
    MYTextMatrix = gData.TextMatrix(Row, Col)
End Function


Private Sub WriteRejectFile()
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    Dim i As Long
    Dim Vendor As String
    
    
    s = ""
    s = s & "SELECT Vendor" & vbCrLf
    s = s & "      ,VendorDesc Companyname" & vbCrLf
    s = s & "      ,Community" & vbCrLf
    s = s & "      ,communityPhase Phase" & vbCrLf
    s = s & "      ,CommunityDesc CommunityName" & vbCrLf
    s = s & "      ,Assembly" & vbCrLf
    s = s & "      ,AssemblyDesc AssemblyDescription" & vbCrLf
    s = s & "      ,Phase gGroup" & vbCrLf
    s = s & "      ,Item" & vbCrLf
    s = s & "      ,ItemDesc ItemDescription" & vbCrLf
    s = s & "      ,PartNumber" & vbCrLf
    s = s & "      ,OrderUOM UOM" & vbCrLf
    s = s & "      ,Current_Cost Price" & vbCrLf
    s = s & "      ,Next_Cost1 NextPrice1" & vbCrLf
    s = s & "      ,Next_Effective1 Effective1" & vbCrLf
    s = s & "      ,Next_Cost2 NextPrice2" & vbCrLf
    s = s & "      ,Next_Effective2 Effective2" & vbCrLf
    s = s & "      ,substring(case when UnknownVendor=1 then '; unknown vendor' else '' end" & vbCrLf
    s = s & "        + case when UnknownCommunity=1 then '; unknown community' else '' end" & vbCrLf
    s = s & "        + case when UnknownAssembly=1 then '; unknown assembly' else '' end" & vbCrLf
    s = s & "        + case when UnknownItem=1 then '; unknown phase/item' else '' end" & vbCrLf
    s = s & "        + case when UnknownUOM=1 then '; HF expects UOM to be ""' + ISNULL(ExpectedUOM,'') + '"" ' else '' end" & vbCrLf
    s = s & "        + case when UnknownExpiryDate=1 then '; effective date is required' else '' end,3,9999) Comments" & vbCrLf
    s = s & "  FROM ImportedCosts " & vbCrLf
    s = s & " WHERE SessionID=" & DbQuote(Str, mSessionID) & vbCrLf
    s = s & "   AND (UnknownVendor=1 OR" & vbCrLf
    s = s & "        UnknownCommunity=1 OR" & vbCrLf
    s = s & "        UnknownAssembly=1 OR" & vbCrLf
    s = s & "        UnknownItem=1 OR" & vbCrLf
    s = s & "        UnknownUOM=1 OR" & vbCrLf
    s = s & "        UnknownExpiryDate=1)" & vbCrLf
    s = s & "ORDER BY 1,3,4,8,9" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    With gRejects
        .Rows = 0
        .Cols = 16
        .FixedCols = 0
        .FixedRows = 0
        While Not rs.EOF
            If Vendor <> "" & rs("vendor") Then
                Vendor = "" & rs("vendor")
                .AddItem ""
                .AddItem ""
                .AddItem "Vendor" & vbTab & Vendor
                .AddItem "Company Name" & vbTab & rs("CompanyName")
                .Cell(flexcpFontBold, .Rows - 2, 0, .Rows - 1, 0) = True
                .AddItem "Community" & vbTab & "Phase" & vbTab & "Community Name" & vbTab & "Assembly" & vbTab & "Assembly Description" & vbTab & "Group" & vbTab & "Item" & vbTab & "Item Description" & vbTab & "Part Number" & vbTab & "UOM" & vbTab & "Price" & vbTab & "Next Price 1" & vbTab & "Effective 1" & vbTab & "Next Price 2" & vbTab & "Effective 2"
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 15) = True
                
                .Refresh
                .Row = .Rows - 1
            End If
            .AddItem "" & rs("Community") & vbTab & rs("Phase") & vbTab & rs("CommunityName") & vbTab & rs("Assembly") & vbTab & rs("AssemblyDescription") & vbTab & rs("gGroup") & vbTab & rs("Item") & vbTab & rs("ItemDescription") & vbTab & rs("PartNumber") & vbTab & rs("UOM") & vbTab & rs("Price") & vbTab & rs("NextPrice1") & vbTab & rs("Effective1") & vbTab & rs("NextPrice2") & vbTab & rs("Effective2") & vbTab & rs("Comments")
            rs.MoveNext
        Wend
        .Cell(flexcpForeColor, 0, 15, .Rows - 1, 15) = vbRed
        
        
        Call .SaveGrid(txtRejectFile.Text, flexFileExcel)
        
    End With
Exit Sub
eh: Call ErrHandler(SRCFILE & "WriteRejectFile")
End Sub



