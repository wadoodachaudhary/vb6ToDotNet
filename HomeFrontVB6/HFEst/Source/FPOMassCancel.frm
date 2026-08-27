VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FPOMassCancel 
   Caption         =   "Purchase Order Cancel Wizard"
   ClientHeight    =   10770
   ClientLeft      =   11580
   ClientTop       =   2235
   ClientWidth     =   7650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FPOMassCancel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   7650
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   1588
      Caption         =   "Cancel Purchase Orders"
      Description     =   "The PO cancel wizard will help you cancel purchase orders"
      Icon            =   "FPOMassCancel.frx":000C
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   7395
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   2895
         Left            =   1020
         TabIndex        =   21
         Top             =   1590
         Width           =   6165
         _cx             =   10874
         _cy             =   5106
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FPOMassCancel.frx":08E6
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
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
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   2
         Left            =   420
         Picture         =   "FPOMassCancel.frx":09BB
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lblDescription 
         Caption         =   $"FPOMassCancel.frx":1285
         Height          =   1125
         Index           =   3
         Left            =   1050
         TabIndex        =   23
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5805
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select POs to be cancelled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   22
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   2340
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   1
         Left            =   240
         Picture         =   "FPOMassCancel.frx":131E
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7650
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   10185
      Width           =   7650
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Commi&t"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   9
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
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   870
      Visible         =   0   'False
      Width           =   7395
      Begin VB.TextBox txtPONumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   5
         Top             =   3030
         Width           =   3915
      End
      Begin VB.TextBox txtPOIndex 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   4
         Top             =   2790
         Width           =   3915
      End
      Begin VB.TextBox txtPOGroup 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   3
         Top             =   2550
         Width           =   3915
      End
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   2
         Top             =   2310
         Width           =   3915
      End
      Begin VB.TextBox txtJob 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   1
         Top             =   2070
         Width           =   3915
      End
      Begin VB.TextBox txtCommunity 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2400
         TabIndex        =   0
         Top             =   1830
         Width           =   3915
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FPOMassCancel.frx":1BE8
         Height          =   1125
         Index           =   1
         Left            =   1050
         TabIndex        =   20
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5865
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PO Number"
         Height          =   195
         Index           =   5
         Left            =   1470
         TabIndex        =   19
         Top             =   3060
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PO Index"
         Height          =   195
         Index           =   4
         Left            =   1635
         TabIndex        =   18
         Top             =   2820
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PO Group"
         Height          =   195
         Index           =   3
         Left            =   1590
         TabIndex        =   17
         Top             =   2580
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendor"
         Height          =   195
         Index           =   2
         Left            =   1785
         TabIndex        =   16
         Top             =   2340
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Job"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   15
         Top             =   2100
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Community"
         Height          =   195
         Index           =   0
         Left            =   1530
         TabIndex        =   14
         Top             =   1860
         Width           =   765
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Purchase Orders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   13
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   1905
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FPOMassCancel.frx":1DB5
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Remove"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FPOMassCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FPOMassCancel::"

Private mRecordset As Recordset

Public Function ShowForm()
    cmdNav(2).Enabled = True
    Me.Show vbModal
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
    Dim c As Long
    Dim r As Long
    Dim s As String
    Dim rs As Recordset
    
    Screen.MousePointer = vbHourglass
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
        
        
    With gData
    Select Case RHS
        Case 0 'filter pos
            Set gData.DataSource = Nothing
            gData.Rows = 1
            
        Case 1 'cancel list
            gData.Rows = 1
            s = ""
            s = s & "select Community,CommunityDesc,Job,JobDesc,Vendor,VendorDesc,POGroup,POIndex,POIndexDescription,PONumber,cast(PODate as date) PODate" & vbCrLf
            s = s & "from purchaseorders" & vbCrLf
            s = s & "where Cancelled=0" & vbCrLf
            s = s & "and DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & WhereClause()
            s = s & "order by 1,2,3,4,5,6,7" & vbCrLf
            Set rs = HFApp.SqlExec(s, dbHomefront)
            
            'load headings
            .Cols = rs.fields.Count + 1
            .TextMatrix(0, 0) = "Selected"
            .ColKey(0) = "Selected"
            .ColDataType(0) = flexDTBoolean
            For c = 1 To gData.Cols - 1
                .TextMatrix(0, c) = rs.fields(c - 1).Name
                .ColKey(c) = rs.fields(c - 1).Name
                .ColDataType(c) = flexDTString
            Next
            .TextMatrix(0, .ColIndex("Community")) = FMain.CD_Community
            .TextMatrix(0, .ColIndex("CommunityDesc")) = FMain.CD_Community & " Desc"
            .ColDataType(.ColIndex("PODate")) = flexDTDate
            .ColFormat(.ColIndex("PODate")) = HFApp.Options(DateFormat)
            
            'load data
            r = 0
            While Not rs.EOF
                .AddItem ""
                r = r + 1
                For c = 1 To gData.Cols - 1
                    .TextMatrix(r, c) = "" & rs(c - 1)
                Next
                rs.MoveNext
            Wend
            
            'size columns
            Call .AutoSize(0, .Cols - 1)
            
            
    End Select
    End With
    
    Screen.MousePointer = vbDefault
    
Exit Property
eh: Call errHandler(SRCFILE & "CurrentFrame")
End Property





Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3: Call ApplyChange
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
    
    gData.Move gData.Left, gData.Top, WizFrame(0).Width - gData.Left - 180, WizFrame(0).Height - gData.Top - 180
    
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadCustomDescriptions
    
    CurrentFrame = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = gData.ColKey(Col) <> "Selected"
End Sub

Private Function WhereClause() As String
    Dim s As String
    
    If txtCommunity.Text <> "" Then s = s & "       and (community like " & DbQuote(Str, txtCommunity.Text) & " or communitydesc like " & DbQuote(Str, txtCommunity.Text) & ")" & vbCrLf
    If txtJob.Text <> "" Then s = s & "       and (job like " & DbQuote(Str, txtJob.Text) & " or jobdesc like " & DbQuote(Str, txtJob.Text) & ")" & vbCrLf
    If txtVendor.Text <> "" Then s = s & "       and (vendor like " & DbQuote(Str, txtVendor.Text) & " or vendordesc like " & DbQuote(Str, txtVendor.Text) & ")" & vbCrLf
    If txtPOIndex.Text <> "" Then s = s & "       and (poindex like " & DbQuote(Str, txtPOIndex.Text) & " or poindexdescription like " & DbQuote(Str, txtPOIndex.Text) & ")" & vbCrLf
    If txtPOGroup.Text <> "" Then s = s & "       and pogroup like " & DbQuote(Str, txtPOGroup.Text) & vbCrLf
    If txtPONumber.Text <> "" Then s = s & "       and ponumber like " & DbQuote(Str, txtPONumber.Text) & vbCrLf
    
    
    WhereClause = Trim(s)
End Function

Private Sub ApplyChange()
On Error GoTo eh
    
    Dim rc As Long
    Dim r As Long
    Dim Count As Long
    Dim email As Long
    Dim reason As String
    
    With gData
    
        Count = 0
        For r = .Rows - 1 To 1 Step -1
        If .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexChecked Then Count = Count + 1
        Next
    
        
        If Count = 0 Then
            MsgBox "No purchase orders are selected", vbInformation, Me.Caption
            Exit Sub
        End If
    
        rc = MsgBox(Count & " purchase orders will be cancelled. Would you like to send emails to the vendors?", vbQuestion + vbYesNoCancel, Me.Caption)
        If rc = vbCancel Then Exit Sub
        If rc = vbYes Then
            email = 1
        Else
            email = 2
        End If
        
        
        If Not FCancelPO.CancelPO(reason) Then
            Exit Sub
        End If
        
        For r = .Rows - 1 To 1 Step -1
        If .Cell(flexcpChecked, r, .ColIndex("Selected")) = flexChecked Then
            If CancelPO(.TextMatrix(r, .ColIndex("PONumber")), False, email, reason) Then Call .RemoveItem(r)
        End If
        Next
    
    End With
    
    
    
Exit Sub
eh: Call errHandler(SRCFILE & "ApplyChange")
End Sub

Private Sub LoadCustomDescriptions()
    Label2(0).Caption = FMain.CD_Community
End Sub
