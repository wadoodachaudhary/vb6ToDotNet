VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FItemPriceUpdateWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Pricing Wizard"
   ClientHeight    =   4770
   ClientLeft      =   2490
   ClientTop       =   480
   ClientWidth     =   6330
   ControlBox      =   0   'False
   Icon            =   "FItemPriceUpdateWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6330
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4185
      Width           =   6330
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5160
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2820
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1620
         TabIndex        =   6
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
         X1              =   -540
         X2              =   25940
         Y1              =   0
         Y2              =   0
      End
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1588
      Caption         =   "Update Item Pricing"
      Description     =   "The item pricing wizard will update the prices of your item database."
      Icon            =   "FItemPriceUpdateWizard.frx":000C
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3600
      Index           =   2
      Left            =   6870
      TabIndex        =   2
      Top             =   4740
      Visible         =   0   'False
      Width           =   6315
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   315
         Left            =   1260
         TabIndex        =   17
         Top             =   540
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         Picture         =   "FItemPriceUpdateWizard.frx":08E6
         ForeColor       =   0
         BarPicture      =   "FItemPriceUpdateWizard.frx":0902
         ShowText        =   -1  'True
         TextAlignX      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Segments        =   -1  'True
         XpStyle         =   -1  'True
      End
      Begin VB.Label lblSaving 
         AutoSize        =   -1  'True
         Caption         =   "The pricing wizard is ready to proceed."
         Height          =   195
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   2730
      End
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         Caption         =   "Pricelist 4 of 14"
         Height          =   195
         Left            =   1500
         TabIndex        =   18
         Top             =   900
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   2
         Left            =   300
         Picture         =   "FItemPriceUpdateWizard.frx":091E
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3600
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   4680
      Width           =   6315
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1155
         Left            =   300
         TabIndex        =   21
         Top             =   2010
         Width           =   7815
         Begin VB.OptionButton optEstimating 
            Caption         =   "Yes, make Estimatings price match HomeFronts."
            Height          =   255
            Index           =   1
            Left            =   900
            TabIndex        =   23
            Top             =   720
            Width           =   4635
         End
         Begin VB.OptionButton optEstimating 
            Caption         =   "No, do not update Estimating."
            Height          =   255
            Index           =   0
            Left            =   900
            TabIndex        =   22
            Top             =   450
            Value           =   -1  'True
            Width           =   4635
         End
         Begin VB.Label lblEstimating 
            AutoSize        =   -1  'True
            Caption         =   "Do you want to update Timberline Estimating?"
            Height          =   195
            Left            =   780
            TabIndex        =   24
            Top             =   120
            Width           =   3240
         End
         Begin VB.Image imgEstimating 
            Height          =   480
            Left            =   0
            Picture         =   "FItemPriceUpdateWizard.frx":11E8
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.OptionButton optUpdate 
         Caption         =   "Use the highest price found in any community."
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   20
         Top             =   570
         Value           =   -1  'True
         Width           =   4635
      End
      Begin VB.ComboBox cboCommunity 
         Enabled         =   0   'False
         Height          =   240
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1620
         Width           =   4035
      End
      Begin VB.OptionButton optUpdate 
         Caption         =   "Use the highest price found in the selected community."
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   12
         Top             =   1110
         Width           =   4635
      End
      Begin VB.OptionButton optUpdate 
         Caption         =   "Use the default vendors price in the selected community."
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   840
         Width           =   4635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Community"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   13
         Top             =   1380
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "What prices do you want to use?"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   2340
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   1
         Left            =   300
         Picture         =   "FItemPriceUpdateWizard.frx":1AB2
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "3"
      Height          =   3600
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gPhases 
         Height          =   2580
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   4695
         _cx             =   8281
         _cy             =   4551
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FItemPriceUpdateWizard.frx":237C
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
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
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
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkAllPhases 
         Caption         =   "All Phases"
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Which phases do you want to update?"
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   2745
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   300
         Picture         =   "FItemPriceUpdateWizard.frx":23C0
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "FItemPriceUpdateWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FItemPriceUpdateWizard::"

Private bInHere As Boolean 'stupid flag see chkAllCommunities_Click() and gCommunities_AfterEdit()

Private Sub Form_Load()
    Dim i As Long
    gPhases.Rows = 0
    Call IniGetForm(Me)
    For i = WizFrame.LBound To WizFrame.UBound
        WizFrame(i).Move 0, 900
    Next
    CurrentFrame = 0
End Sub

Private Sub cmdNav_Click(Index As Integer)
'On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3: If SaveData Then Unload Me
    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

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
    Dim i As Long
    
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
    
    Dim rs As Recordset
    Dim r As Long
    Dim c As Long
    Dim s As String
    
    Select Case RHS
        Case 0 'phases
            With gPhases
                If .Rows = 0 Then
                    .Redraw = flexRDNone
                    r = -1
                    Set rs = HFApp.SqlExec("SELECT Phase,Description FROM tblEstPhases WHERE DivisionID = " & HFApp.DivisionID & " and GroupPhase=0 ORDER BY Phase")
                    While Not rs.EOF
                        r = r + 1
                        .AddItem rs(0) & vbTab & rs(1)
                        rs.MoveNext
                    Wend
                    .Cell(flexcpChecked, 0, 0, .Rows - 1, 0) = flexUnchecked
                    chkAllPhases.value = vbChecked
                    Call .AutoSize(0, .cols - 1)
                    .Redraw = flexRDBuffered
                End If
            End With
            
        Case 1
            s = ""
            s = s & "SELECT ISNULL(l.Area,'')+' -- '+ISNULL(l.Description,'')" & vbCrLf
            s = s & "      ,ISNULL(l.Area,'')" & vbCrLf
            s = s & "      ,0 " & vbCrLf
            s = s & "  FROM tblLocality l" & vbCrLf
            s = s & " WHERE l.Inactive<>1" & vbCrLf
            s = s & "UNION ALL" & vbCrLf
            s = s & "SELECT ISNULL(l.Area,'')+'.'+ISNULL(p.CommunityPhase,'') + ' -- ' + ISNULL(l.description,'') + ISNULL(' ' + p.description,'')" & vbCrLf
            s = s & "      ,ISNULL(l.Area,'')+CHAR(2)+ISNULL(p.CommunityPhase,'')" & vbCrLf
            s = s & "      ,0" & vbCrLf
            s = s & "  FROM tblLocality l" & vbCrLf
            s = s & "       JOIN CommunityPhase p ON (l.Area=p.Community)" & vbCrLf
            s = s & "WHERE l.Inactive<>1" & vbCrLf
            s = s & "ORDER BY 1" & vbCrLf
            Call LoadComboBox(cboCommunity, HFApp.Databases(dbHomeFront), s)
            cboCommunity.AddItem "Global Pricing", 0
            cboCommunity.tag = Chr(1) & cboCommunity.tag
            
            
            imgEstimating.Visible = HFApp.Databases(dbEstimating).State = adStateOpen
            lblEstimating.Visible = HFApp.Databases(dbEstimating).State = adStateOpen
            optEstimating(0).Visible = HFApp.Databases(dbEstimating).State = adStateOpen
            optEstimating(1).Visible = HFApp.Databases(dbEstimating).State = adStateOpen
            

            On Error Resume Next
            cboCommunity.ListIndex = 0
            
        Case 2
        
    End Select
    
End Property

Private Sub chkAllPhases_Click()
    If Not bInHere Then
        gPhases.Cell(flexcpChecked, 0, 0, gPhases.Rows - 1, 0) = IIf(chkAllPhases.value = vbChecked, flexChecked, flexUnchecked)
    End If
End Sub
Private Sub gPhases_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    bInHere = True
    With gPhases
    For i = 0 To .Rows - 1
        If .Cell(flexcpChecked, 0, 0) <> .Cell(flexcpChecked, i, 0) Then
            chkAllPhases.value = vbGrayed
            bInHere = False
            Exit Sub
        End If
    Next
    chkAllPhases.value = IIf(.Cell(flexcpChecked, 0, 0) = flexChecked, vbChecked, vbUnchecked)
    End With
    bInHere = False
End Sub


Private Function SaveData() As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    Dim Count As Long
    Dim i     As Long
    Dim phases As String
    Dim community As String
    Dim CommunityPhase As String
    Dim WhereClause As String

    cmdNav(0).Enabled = False
    cmdNav(1).Enabled = False
    cmdNav(2).Enabled = False
    cmdNav(3).Enabled = False
        
    
    phases = SelectedPhases
    community = GetComboBoxListKey(cboCommunity)
    CommunityPhase = Parse(community, 2, Chr(2))
    community = Parse(community, 1, Chr(2))
        
        
    'update list
    lblSaving.Visible = True
    lblSaving.Caption = "Querying vendor pricelists..."
    Select Case True
    
        Case optUpdate(0):  'highest price in db
            s = ""
            s = s & "update tblphaseitem " & vbCrLf
            s = s & "set price = (select max(current_cost)" & vbCrLf
            s = s & "               from tblvendorcost " & vbCrLf
            s = s & "              where tblvendorcost.phase=tblphaseitem.phase " & vbCrLf
            s = s & "                and tblvendorcost.item=tblphaseitem.item and tblvendorcost.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
            s = s & "where 0<> (select isnull(max(current_cost),0) " & vbCrLf
            s = s & "               from tblvendorcost " & vbCrLf
            s = s & "              where tblvendorcost.phase=tblphaseitem.phase " & vbCrLf
            s = s & "                and tblvendorcost.item=tblphaseitem.item and tblvendorcost.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
            If phases <> "" Then
                s = s & "   and Phase IN(" & phases & ")" & vbCrLf
            End If
            s = s & " and DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
         
        
        Case optUpdate(1):  'use default vendor
            s = ""
            s = s & "UPDATE tblPhaseItem" & vbCrLf
            s = s & "   SET Price=Current_Cost" & vbCrLf
            s = s & "  FROM tblVendorCost c JOIN tblPhaseitem i ON(c.DivisionID = i.DivisionID and c.Phase=i.Phase AND c.Item=i.Item)" & vbCrLf
            s = s & " WHERE c.Vendor=dbo.Purch_GetCommunityVendor(c.Community,i.POIndex," & HFApp.DivisionID & ")" & vbCrLf
            If phases <> "" Then
                s = s & "   AND i.Phase IN(" & phases & ")" & vbCrLf
            End If
            s = s & "   AND c.Community=" & DbQuote(str, community) & vbCrLf
            s = s & "   AND c.CommunityPhase=" & DbQuote(str, CommunityPhase) & vbCrLf
            s = s & "   AND c.Assembly=''" & vbCrLf
            s = s & "   AND c.Model=''" & vbCrLf
            s = s & "   AND c.DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
            
            
        Case optUpdate(2):  'use greatest cost
            s = ""
            s = s & "UPDATE tblPhaseItem" & vbCrLf
            s = s & "   SET Price=(SELECT MAX(Current_Cost)" & vbCrLf
            s = s & "                FROM tblVendorCost" & vbCrLf
            s = s & "               WHERE ISNULL(Community,'')=" & DbQuote(str, community) & vbCrLf
            s = s & "                 AND ISNULL(CommunityPhase,'')=" & DbQuote(str, CommunityPhase) & vbCrLf
            s = s & "                 AND ISNULL(Assembly,'')=''" & vbCrLf
            s = s & "                 AND ISNULL(Model,'')=''" & vbCrLf
            s = s & "                 AND tblVendorCost.Phase=tblPhaseItem.Phase" & vbCrLf
            s = s & "                 AND tblVendorCost.Item=tblPhaseItem.Item)" & vbCrLf
            s = s & "                 AND tblVendorCost.DivisionID = " & HFApp.DivisionID
            s = s & " WHERE 0<>(SELECT ISNULL(MAX(Current_Cost),0)" & vbCrLf
            s = s & "                FROM tblVendorCost" & vbCrLf
            s = s & "               WHERE ISNULL(Community,'')=" & DbQuote(str, community) & vbCrLf
            s = s & "                 AND ISNULL(CommunityPhase,'')=" & DbQuote(str, CommunityPhase) & vbCrLf
            s = s & "                 AND ISNULL(Assembly,'')=''" & vbCrLf
            s = s & "                 AND ISNULL(Model,'')=''" & vbCrLf
            s = s & "                 AND tblVendorCost.Phase=tblPhaseItem.Phase" & vbCrLf
            s = s & "                 AND tblVendorCost.Item=tblPhaseItem.Item)" & vbCrLf
            s = s & "                 AND tblVendorCost.DivisionID = " & HFApp.DivisionID
            If phases <> "" Then
                s = s & "   AND Phase IN(" & phases & ")" & vbCrLf
            End If
            s = s & " and DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
        
    End Select
    
    If HFApp.Databases(dbEstimating).State = adStateOpen And optEstimating(1).value Then
        s = ""
        s = s & "select count(*) from EstimatingItems "
        If phases <> "" Then
            s = s & " where DivisionID = " & HFApp.DivisionID & " and phase IN(" & phases & ")" & vbCrLf
        Else
            s = s & " where DivisionID = " & HFApp.DivisionID
        End If
        Count = HFApp.SqlExec(s)(0)
        
        
        s = ""
        s = s & "select * from EstimatingItems Where DivisionID = " & HFApp.DivisionID & " and "
        If phases <> "" Then s = s & " where phase IN(" & phases & ")" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        i = 0
        While Not rs.EOF
                        
            DoEvents
            
            'show progress
            i = i + 1
            ProgressBar.Visible = True
            lblProgress.Visible = True
            lblSaving.Caption = "Updating Estimating prices..."
            ProgressBar.value = i / Count * 100
            lblProgress.Caption = "item " & i & " of " & Count
            lblProgress.Refresh
    
    
    
            
            s = ""
            s = s & "UPDATE DAT_PEI__DB_ITEM" & vbCrLf
            Select Case "" & rs("CostCategory")
                Case "S":  s = s & "   SET sub_price=" & DbQuote(Num, Abs(Val("" & rs("Price")))) & vbCrLf
                Case "L":  s = s & "   SET labor_price=" & DbQuote(Num, Abs(Val("" & rs("Price")))) & vbCrLf
                Case "O":  s = s & "   SET other_price=" & DbQuote(Num, Abs(Val("" & rs("Price")))) & vbCrLf
                Case "E":  s = s & "   SET equipment_price=" & DbQuote(Num, Abs(Val("" & rs("Price")))) & vbCrLf
                Case Else: s = s & "   SET material_price=" & DbQuote(Num, Abs(Val("" & rs("Price")))) & vbCrLf
            End Select
            s = s & " WHERE Phase_Code=" & DbQuote(str, HFApp.FormatPhase("" & rs("Phase"))) & vbCrLf
            s = s & "   AND Item_Number=" & DbQuote(str, HFApp.FormatItem("" & rs("ItemNumber"))) & vbCrLf
            Call HFApp.SqlExec(s, dbEstimating)
            
            
            
            If Val("" & rs("PriceLink")) <> 0 Then
                Select Case "" & rs("CostCategory")
                    Case "M": s = "UPDATE DAT_PEI__DB_ITEM SET material_price =" & DbQuote(Num, Abs(Val("" & rs("Price")))) & " WHERE mat_price_link=" & DbQuote(Num, "" & rs("PriceLink")) & vbCrLf
                    Case "L": s = "UPDATE DAT_PEI__DB_ITEM SET labor_price    =" & DbQuote(Num, Abs(Val("" & rs("Price")))) & " WHERE lab_price_link=" & DbQuote(Num, "" & rs("PriceLink")) & vbCrLf
                    Case "E": s = "UPDATE DAT_PEI__DB_ITEM SET equipment_price=" & DbQuote(Num, Abs(Val("" & rs("Price")))) & " WHERE  eq_price_link=" & DbQuote(Num, "" & rs("PriceLink")) & vbCrLf
                End Select
                Call HFApp.SqlExec(s, dbEstimating)
            End If
            
            
            rs.MoveNext
        Wend
    End If
    
        
        
    SaveData = True
    Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "SaveData")
    End If
End Function


Private Function SelectedPhases() As String
    Dim r As Long
    Dim s As String

    s = ""
    If chkAllPhases.value <> vbChecked Then
        With gPhases
        For r = 0 To .Rows - 1
            If .Cell(flexcpChecked, r, 0) = flexChecked Then
                s = s & "," & DbQuote(str, .TextMatrix(r, 0))
            End If
        Next
        End With
    End If
    SelectedPhases = Mid(s, 2)
    
End Function


Private Sub gPhases_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> 0
End Sub

Private Sub optUpdate_Click(Index As Integer)
    cboCommunity.Enabled = SelectedOption(optUpdate) <> 0
    Label3(4).Enabled = cboCommunity.Enabled
End Sub

