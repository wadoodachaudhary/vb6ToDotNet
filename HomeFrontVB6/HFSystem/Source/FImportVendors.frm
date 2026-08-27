VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FImportVendors 
   Caption         =   "Pricelist Import Wizard"
   ClientHeight    =   6165
   ClientLeft      =   2190
   ClientTop       =   1920
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FImportVendors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7305
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7305
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5580
      Width           =   7305
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   -60
         X2              =   26420
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
   End
   Begin HFSystem.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1588
      Caption         =   "Import Vendors"
      Description     =   "The vendor import wizard will help you update your database"
      Icon            =   "FImportVendors.frx":000C
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   900
      Width           =   7515
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   3060
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   556
         Picture         =   "FImportVendors.frx":08E6
         ForeColor       =   0
         BarPicture      =   "FImportVendors.frx":0902
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   2655
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   7155
         _cx             =   1986408781
         _cy             =   1986400843
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FImportVendors.frx":091E
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
      Begin VB.Label lblFileDate 
         AutoSize        =   -1  'True
         Caption         =   "Last Modified April 18, 2007"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label lblFilename 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   45
      End
   End
End
Attribute VB_Name = "FImportVendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FImportVendors::"
Private mFilename  As String

Public Function ShowForm()
On Error GoTo eh
    Dim i As Long
    i = IniGet(AppIni, "Options", "VendorImportFileExt", 0)
    
    If VBGetOpenFileName(mFilename, , , , , True, "Excel Files (*.xls)|*.xls|List Files (*.iif)|*.iif|Text Files (*.txt;*.csv)|*.txt;*.csv|All Files (*.*)|*.*", i, , , , FVendor.hwnd) Then
        Call IniPut(AppIni, "Options", "VendorImportFileExt", i)
        
        CurrentFrame = 0
        Me.Show vbModal
    Else
        Unload Me
    End If
        
Exit Function
eh: Call errHandler(SRCFILE & "ShowForm")
End Function


Private Function TextMatrix(r As Long, ColKey As String) As String
On Error Resume Next
    TextMatrix = gData.TextMatrix(r, gData.ColIndex(ColKey))
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
            If WizFrame(RHS).tag = "" Then
                WizFrame(RHS).tag = "LOADED"
                lblFilename.Caption = "File name: " & mFilename
                fileinfo.FullPathName = mFilename
                lblFileDate.Caption = "Last Modified: " & Format(fileinfo.ModifyTime, "long date")
                
                Select Case FileExt(mFilename)
                    Case "xls", "xlsx"
                        mFilename = SaveToCSV(mFilename)
                        Call ReadText(mFilename)
                        On Error Resume Next
                        Kill mFilename
                        On Error GoTo eh
                    Case Else
                        Call ReadText(mFilename)
                End Select

            End If
                            
    End Select
    End With
    Screen.MousePointer = vbDefault
    Exit Property
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "CurrentFrame", s)
    End If
End Property

Private Function SaveToCSV(Filename As String) As String
On Error Resume Next
    Dim s As String
    Dim xlSheet As Object 'Excel.Worksheet
    Set xlSheet = GetObject(Filename).Sheets(1)
    s = TempFile("csv")
    Kill s
    Call xlSheet.SaveAs(s, 6, , , , , False)
    Set xlSheet = Nothing
    SaveToCSV = s
End Function

Private Sub ReadText(Filename As String)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    Select Case FileExt(Filename)
        Case "csv": Call gData.LoadGrid(Filename, flexFileCommaText)
        
        Case "iif": Call gData.LoadGrid(Filename, flexFileTabText)
                    With gData
                        'remove leading rows
                        While .Rows > 1 And .TextMatrix(1, 1) <> "!VEND"
                            Call .RemoveItem(1)
                        Wend
                        'remove trailing rows
                        While .Rows > 1 And .TextMatrix(.Rows - 1, 1) <> "VEND"
                            Call .RemoveItem(.Rows - 1)
                        Wend
                    End With
        
        Case Else:  Call gData.LoadGrid(Filename, flexFileTabText)
    End Select
    
    
    
    With gData
    
    'set column indexs
    For i = 1 To .Cols - 1
        .ColKey(i) = Replace(.TextMatrix(1, i), " ", "")
    Next
    Call .RemoveItem(0)
    .FixedRows = 1
    
    'translate column headings
    On Error Resume Next
    If .ColIndex("refnum") <> -1 Then .ColKey(.ColIndex("refnum")) = "Vendor_ID"
    If .ColIndex("vendor") <> -1 Then .ColKey(.ColIndex("vendor")) = "Vendor_ID"
    If .ColIndex("vendorcode") <> -1 Then .ColKey(.ColIndex("vendorcode")) = "Vendor_ID"
    If .ColIndex("id") <> -1 Then .ColKey(.ColIndex("id")) = "Vendor_ID"
    If .ColIndex("code") <> -1 Then .ColKey(.ColIndex("code")) = "Vendor_ID"
    If .ColIndex("name") <> -1 Then .ColKey(.ColIndex("name")) = "Vendor_Name"
    If .ColIndex("name") <> -1 Then .ColKey(.ColIndex("name")) = "Vendor_Name"
    If .ColIndex("prov") <> -1 Then .ColKey(.ColIndex("prov")) = "state"
    If .ColIndex("province") <> -1 Then .ColKey(.ColIndex("province")) = "state"
    If .ColIndex("postalcode") <> -1 Then .ColKey(.ColIndex("postalcode")) = "zip"
    If .ColIndex("zipcode") <> -1 Then .ColKey(.ColIndex("zipcode")) = "zip"
    If .ColIndex("emailaddress") <> -1 Then .ColKey(.ColIndex("emailaddress")) = "email"
    If .ColIndex("phone1") <> -1 Then .ColKey(.ColIndex("phone1")) = "phone"
    If .ColIndex("faxnum") <> -1 Then .ColKey(.ColIndex("faxnum")) = "fax"
    If .ColIndex("cont1") <> -1 Then .ColKey(.ColIndex("cont1")) = "contact"
    If .ColIndex("vtype") <> -1 Then .ColKey(.ColIndex("vtype")) = "tradetype"
    On Error GoTo eh
    For i = 1 To .Cols - 1
        .TextMatrix(0, i) = .ColKey(i)
    Next
    
    
    'check that required cols are available
    If .ColIndex("Vendor_ID") = -1 Then: Err.Raise 5, , "Required column ""Description"" is missing"
    
    'remove rows that have no vendor id
    For i = .Rows - 1 To 1 Step -1
        If Trim(.TextMatrix(i, .ColIndex("Vendor_ID")) = "") Then Call .RemoveItem(i)
    Next
    
    End With
Exit Sub
eh: MsgBox Err.Description, vbExclamation, App.ProductName
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
On Error GoTo eh
    Call IniGetForm(Me)
Exit Sub
eh: Call errHandler(SRCFILE & "SaveData")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub



Private Function SaveData() As Boolean
On Error GoTo eh
    Dim tabdef As Recordset
    Dim bIgnoreDupKey As Boolean
    Dim r As Long
    Dim c As Long
    Dim X As Long
    Dim ctype As String
    Dim d As adodb.field
    Dim s As String
    
    Screen.MousePointer = vbHourglass
    ProgressBar.Visible = True
    
    Set tabdef = HFApp.SqlExec("select * from tblvendors where 1=2")
    With gData
    For r = 1 To .Rows - 1
        'Me.Refresh
        '.Refresh
        'DoEvents
                
        ProgressBar.Value = (r / (.Rows - 1)) * 100
        
        
        s = "INSERT INTO tblVendors(DivisionID,webupdated,nodatachanged,Vendor_ID) VALUES (" & HFApp.DivisionID & ",0,0," & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor_ID")), , True) & ")"
        bIgnoreDupKey = True
        Call HFApp.SqlExec(s, dbHomefront)
        bIgnoreDupKey = False
        
        s = ""
        For c = 1 To .Cols - 1
            On Error Resume Next
            Set d = Nothing
            Set d = tabdef.fields(.ColKey(c))
            On Error GoTo eh
            If Not IsIn(.ColKey(c), "vendor_id", "divisionid", "paymentterms", "SchedContactID", "ServiceContactID", "FPOContactID", "PurchContactID", "createddate", "createdby", "modifieddate", "modifiedby", "DateSentToBuildPro") And Not d Is Nothing Then
                s = s & "," & d.Name & "=" & DbQuote(DbQuoteType(d.Type), .TextMatrix(r, c), False, True, d.DefinedSize)
            End If
        Next
        s = "UPDATE tblVendors SET " & Mid(s, 2) & " WHERE DivisionID = " & HFApp.DivisionID & " and Vendor_ID=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor_ID")))
        Call HFApp.SqlExec(s, dbHomefront)
    
    
    
        'Update each contact field individually to prevent us from clearing data that wasnt provided in the import file. Only update fields that were provided
        s = ""
        For X = 1 To 4
        
            ctype = Choose(X, "FPO", "Purch", "Service", "Sched")
            
            s = ""
            c = .ColIndex(ctype & "contact"):    If c > -1 Then s = s & " ,firstname = " & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
            c = .ColIndex(ctype & "phone"):      If c > -1 Then s = s & " ,workphone = " & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
            c = .ColIndex(ctype & "email"):      If c > -1 Then s = s & " ,email = " & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
            c = .ColIndex(ctype & "fax"):        If c > -1 Then s = s & " ,fax = " & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
            c = .ColIndex(ctype & "smsaddress"): If c > -1 Then s = s & " ,smsaddress = " & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
            c = .ColIndex(ctype & "cell"):       If c > -1 Then s = s & " ,cellphone = " & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
            
            'DelMeth=CommMode
            'print  0=1
            'email  1=3
            'fax    2=4
            c = .ColIndex(ctype & "delmethod"):
            If c > -1 Then
                If IsIn(.TextMatrix(r, c), "0", "1", "2") Then
                    s = s & " ,CommModeID = " & DbQuote(Num, Choose(.ValueMatrix(r, c) + 1, "1", "3", "4")) & vbCrLf
                End If
            End If
            
            If s <> "" Then
                s = "update contacts set" & vbCrLf & "  " & Mid(s, 3)
                s = s & "from contacts c" & vbCrLf
                s = s & "join tblvendors v on c.contactid=" & ctype & "contactid" & vbCrLf
                s = s & "where v.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "and v.Vendor_ID=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor_ID")))
                Call HFApp.SqlExec(s, dbHomefront)
            End If
            
        Next
    
    Next
    End With
    
    ProgressBar.Visible = False
    Screen.MousePointer = vbDefault
    
    SaveData = True
Exit Function
eh:
    If bIgnoreDupKey And InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "SaveData", s)
        ProgressBar.Visible = False
    End If
End Function

