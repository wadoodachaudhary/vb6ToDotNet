VERSION 5.00
Begin VB.Form FPriceListUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Vendor Pricelists"
   ClientHeight    =   3720
   ClientLeft      =   1080
   ClientTop       =   780
   ClientWidth     =   5655
   Icon            =   "FPriceListUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1155
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   14355
      Begin VB.OptionButton optCommunity 
         Caption         =   "Any phase in Aspen Ridge."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   14355
      End
      Begin VB.OptionButton optCommunity 
         Caption         =   "Aspen Ridge Phase 3 only."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   14355
      End
      Begin VB.OptionButton optCommunity 
         Caption         =   "Everywhere."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   14355
      End
      Begin VB.Label lblCommunity 
         AutoSize        =   -1  'True
         Caption         =   "What communities does this price apply to?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3705
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   14355
      Begin VB.OptionButton optType 
         Caption         =   "Unit price. The item is always this price."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   14355
      End
      Begin VB.OptionButton optType 
         Caption         =   "Quote. It applies to this model or option only."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   14355
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "What type of price is this?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   0
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1560
      Picture         =   "FPriceListUpdate.frx":000C
      TabIndex        =   2
      ToolTipText     =   "Login"
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2865
      Picture         =   "FPriceListUpdate.frx":0596
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "It will change City of Calgary's price for        Permits and Fees to $450.00 per EACH."
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   5880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This will update your pricing database."
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
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   3300
   End
   Begin VB.Image imgMsg 
      Height          =   240
      Index           =   1
      Left            =   540
      Picture         =   "FPriceListUpdate.frx":0B20
      Top             =   420
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   180
      Picture         =   "FPriceListUpdate.frx":10AA
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "FPriceListUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FPriceListUpdate::"

Private mPrefix As String
Private mCommunity As String
Private mCommunityPhase As String
Private mAssemblyType As Integer
Private gData   As VSFlexGrid


Public Sub ShowForm(ShowBudgets As Boolean, Grid As VSFlexGrid, community As String, CommunityPhase As String, CommunityDescription As String, AssemblyType As String)
On Error GoTo eh
    Dim r As Long
    Dim c As Long
    
    Load Me

    Set gData = Grid
    mCommunity = community
    mCommunityPhase = CommunityPhase
    mAssemblyType = Val("" & AssemblyType)
    mPrefix = IIf(ShowBudgets, "Budget", "PO")
    With gData
    If .RowSel = .Row Then
        lblPrice.Caption = "What type of price is this?"
        lblCommunity.Caption = "What " & FMain.CD_Community & " does this price apply to?"
        lblMessage.Caption = "It will change " & .TextMatrix(.Row, .ColIndex(mPrefix & "VendorName")) & "'s price for" & vbCrLf & _
                             .TextMatrix(.Row, .ColIndex("ItemDesc")) & " to " & format(.TextMatrix(.Row, .ColIndex(mPrefix & "Rate")), "$#,##0.00###") & " per " & .TextMatrix(.Row, .ColIndex("OrderUOM"))
        
    Else
    
        c = 0
        For r = .Row To .RowSel
            If Not .IsSubtotal(r) Then c = c + 1
        Next
    
    
        lblPrice.Caption = "What type of prices are these?"
        lblCommunity.Caption = "What " & FMain.CD_Community & " do these prices apply to?"
        lblMessage.Caption = "It will update the vendors pricing for these " & c & " items."
    End If
    End With
    
    optCommunity(1).Caption = FMain.CD_Community & " in this phase only."
    optCommunity(2).Caption = "Any phase in " & FMain.CD_Community & "."
    
    
    Me.Show vbModal
Exit Sub
eh: Call errHandler(SRCFILE & "ShowForm")
End Sub
                
                
Private Sub SaveData()
On Error GoTo eh
    Dim s As String
    Dim r As Long
    Dim sAssembly As String
    Dim sModel As String
    Dim rs As ADODB.Recordset
    Dim sEstAssemblyID As Long
    
    
    
    With gData
    For r = .Row To .RowSel
    If Not .IsSubtotal(r) Then
    
    sEstAssemblyID = .ValueMatrix(r, .ColIndex("EstAssemblyID"))
    
    
        s = ""
        s = s & "INSERT INTO tblVendorCost(DivisionID,Community,CommunityPhase,Model,Assembly,Phase,Item,Vendor,Current_Cost)" & vbCrLf
        s = s & "VALUES(" & HFApp.DivisionID & vbCrLf
        s = s & "   ," & DbQuote(Str, IIf(optCommunity(1).value Or optCommunity(2).value, mCommunity, "")) & vbCrLf
        s = s & "   ," & DbQuote(Str, IIf(optCommunity(1).value, mCommunityPhase, "")) & vbCrLf
        If mAssemblyType <> 3 Then
            If optType(0).value Then
                s = s & "   ," & DbQuote(Str, IIf(optType(0).value, .TextMatrix(r, .ColIndex("Model")), "")) & vbCrLf
            Else
                s = s & "   ,''" & vbCrLf
            End If
        Else
            s = s & "   ,''" & vbCrLf
        End If
        If optType(0).value Then
            s = s & "   ," & DbQuote(Str, IIf(optType(0).value, .TextMatrix(r, .ColIndex("Assembly")), "")) & vbCrLf
        Else
            s = s & "   ,''" & vbCrLf
        End If
        s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & vbCrLf
        s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & vbCrLf
        s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex(mPrefix & "Vendor"))) & vbCrLf
        s = s & "   ," & DbQuote(Num, .TextMatrix(r, .ColIndex(mPrefix & "Rate"))) & ")" & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        
        
        s = ""
        s = s & "UPDATE tblVendorCost" & vbCrLf
        s = s & "   SET Current_Cost=" & DbQuote(Num, .TextMatrix(r, .ColIndex(mPrefix & "Rate"))) & vbCrLf
        s = s & " WHERE Vendor=" & DbQuote(Str, .TextMatrix(r, .ColIndex(mPrefix & "Vendor"))) & vbCrLf
        s = s & "   AND Phase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & vbCrLf
        s = s & "   AND Item=" & DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & vbCrLf
        If optType(0).value Then
            s = s & "   AND Assembly=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & vbCrLf
            If mAssemblyType <> 3 Then
                s = s & "   AND Model=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
            Else
                s = s & "   AND Model=''" & vbCrLf
            End If
        End If
        If optCommunity(1).value Or optCommunity(2).value Then
            s = s & "   AND Community=" & DbQuote(Str, mCommunity) & vbCrLf
        End If
        If optCommunity(1).value Then
            s = s & "   AND CommunityPhase=" & DbQuote(Str, mCommunityPhase) & vbCrLf
        End If
        s = s & " and DivisionID = " & HFApp.DivisionID
        Call HFApp.SqlExec(s, dbHomefront)
        
    End If
    Next
    End With
    
    
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "SaveData", s)
    End If
End Sub
                


Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then SaveData
    Unload Me
End Sub


Private Sub LoadCustomDescriptions()
    lblCommunity.Caption = "What " & FMain.CD_Community & " does this price apply to?"
End Sub

Private Sub Form_Load()
    Call LoadCustomDescriptions
End Sub
