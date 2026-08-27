VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FSelectInvoicesToPay 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Select To Pay"
   ClientHeight    =   5325
   ClientLeft      =   6945
   ClientTop       =   4005
   ClientWidth     =   8085
   Icon            =   "FSelectInvoicesToPay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin HFPayables.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   1588
      Caption         =   "Which invoices will be paid via Wallet?"
      Description     =   "Set the bank account to pay an invoice through your Hyphen Wallet."
      Icon            =   "FSelectInvoicesToPay.frx":000C
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   6765
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Save"
      Height          =   375
      Index           =   0
      Left            =   5445
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3645
      Left            =   135
      TabIndex        =   0
      Top             =   1035
      Width           =   7830
      _cx             =   13811
      _cy             =   6429
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FSelectInvoicesToPay.frx":045E
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
Attribute VB_Name = "FSelectInvoicesToPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FSelectInvoicesToPay::"
Private mDirty As Boolean





Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then If Not SaveData(False) Then Exit Sub
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
End Sub

Private Sub Form_Resize()
On Error Resume Next
Const margin = 120

    gData.Move margin, WizHead1.Height + margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - 3 * margin - cmdNav(0).Height - WizHead1.Height
        
    cmdNav(0).Move Me.ScaleWidth - 2 * cmdNav(0).Width - 2 * gData.left, Me.ScaleHeight - cmdNav(0).Height - gData.left
    cmdNav(1).Move Me.ScaleWidth - cmdNav(0).Width - gData.left, Me.ScaleHeight - cmdNav(0).Height - gData.left
End Sub

Private Sub LoadData()
On Error GoTo eh
    Dim s As String
    Dim banks As String
    Dim rs As Recordset
    
    'load bank accounts
    s = "select Bank_Account,Description from cmm_master__bank_account"
    Set rs = HFApp.SqlExec(s, dbAccounting)
    banks = ""
    While Not rs.EOF
        banks = banks & "|" & rs(0) & vbTab & rs(1)
        rs.MoveNext
    Wend
    banks = Mid(banks, 2)
    If banks = "" Then err.Raise vbObjectError, "Unable to read bank accounts from Sage"
            
    'load selected invoices
    s = ""
    s = s & "select i.Vendor,v.Name,i.Invoice,i.Invoice_Date InvoiceDate,i.Description,i.pmt_amount NetToPay,min(d.bank_account) BankAccount" & vbCrLf
    s = s & "from apm_master__invoice i " & vbCrLf
    s = s & "left join apm_master__distribution d on i.vendor=d.vendor and i.invoice=d.invoice" & vbCrLf
    s = s & "left join apm_master__vendor v on i.vendor=v.vendor " & vbCrLf
    s = s & "where i.included_for_pmt=1 and i.status<>'Fully paid'" & vbCrLf
    s = s & "group by i.Vendor,v.Name,i.Invoice,i.Invoice_Date,i.Description,i.pmt_amount" & vbCrLf
    s = s & "order by 1, 2" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbAccounting)
    With gData
        .Rows = 1
        While Not rs.EOF
            .AddItem "" & rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & rs(3) & vbTab & rs(4) & vbTab & rs(5) & vbTab & rs(6) & vbTab & rs(6)
            rs.MoveNext
        Wend
        .ColComboList(.ColIndex("BankAccount")) = banks
        
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
    End With

Exit Sub
eh: Call ErrHandler(SRCFILE & "LoadData")
End Sub


Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim FileName As String
    Dim stream   As ADODB.stream
    Dim s As String
    Dim i As Long
    
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
    
    
    Set stream = New ADODB.stream
    stream.Open
    stream.Position = 0
    stream.Charset = "UTF-16"
    stream.WriteText "<?xml version='1.0' encoding=""UTF-16"" ?>" & vbCrLf
    stream.WriteText "<SetInvBank>" & vbCrLf
    
    With gData
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("OrigBankAccount")) <> .TextMatrix(i, .ColIndex("BankAccount")) Then
            
                stream.WriteText "<Inv Vendor=" & AddQuotes(fixxml(.TextMatrix(i, .ColIndex("Vendor")))) & _
                                     " Invoice=" & AddQuotes(fixxml(.TextMatrix(i, .ColIndex("Invoice")))) & _
                                     " Bank= " & AddQuotes(fixxml(.TextMatrix(i, .ColIndex("BankAccount")))) & _
                                     " />" & vbCrLf

            End If
        Next
    End With
    
    stream.WriteText "</SetInvBank>" & vbCrLf
    On Error Resume Next
    FileName = TempFile("xml")
    Kill FileName
    On Error GoTo eh
    Call stream.SaveToFile(FileName, adSaveCreateOverWrite)
    
    'remove trailing slash
    s = Trim(HFApp.Options(Timberline_Data_Path))
    While Right(s, 1) = "\"
        s = Mid(s, 1, Len(s) - 1)
    Wend
    s = AddQuotes(PathAppend(App.path, "\S3W\Sage300Wrapper.exe")) & " " & _
        AddQuotes(s) & " " & _
        AddQuotes(HFApp.Options(Timberline_UID)) & " " & _
        AddQuotes(HFApp.Options(Timberline_PWD)) & " " & _
        AddQuotes(FileName)
    Call ShellAndLoop(s, vbHide)
    
    
    
    
    SaveData = True
    mDirty = False
    cmdNav(0).Enabled = False
    
    
Exit Function
eh: Call ErrHandler(SRCFILE & "SaveData")
End Function

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = gData.ColKey(Col) <> "BankAccount"
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mDirty = True
End Sub

Private Function fixxml(s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    fixxml = s
End Function

