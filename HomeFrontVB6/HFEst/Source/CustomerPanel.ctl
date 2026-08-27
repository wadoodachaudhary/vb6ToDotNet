VERSION 5.00
Begin VB.UserControl CustomerPanel 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   ScaleHeight     =   3585
   ScaleWidth      =   5190
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin VB.Label lblLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2580
         TabIndex        =   18
         Top             =   2460
         Width           =   1275
      End
      Begin VB.Image imgLink 
         Height          =   360
         Index           =   1
         Left            =   2130
         Picture         =   "CustomerPanel.ctx":0000
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contacts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   810
         TabIndex        =   17
         Top             =   2460
         Width           =   915
      End
      Begin VB.Image imgLink 
         Height          =   360
         Index           =   0
         Left            =   360
         Picture         =   "CustomerPanel.ctx":06EA
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pending Changes:"
         Height          =   285
         Index           =   3
         Left            =   0
         TabIndex        =   16
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label lblJTDPercent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(42%)"
         Height          =   195
         Left            =   3300
         TabIndex        =   15
         Top             =   1710
         Width           =   390
      End
      Begin VB.Label lblCostToComplete 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "219,388.03"
         Height          =   285
         Left            =   1590
         TabIndex        =   14
         Top             =   1950
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cost to Complete:"
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   13
         Top             =   1950
         Width           =   1545
      End
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   3300
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblJTDCost 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "150,847.22"
         Height          =   285
         Left            =   1590
         TabIndex        =   12
         Top             =   1710
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Job to Date Cost:"
         Height          =   285
         Index           =   5
         Left            =   0
         TabIndex        =   11
         Top             =   1710
         Width           =   1545
      End
      Begin VB.Label lblBudget 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "370,253.25"
         Height          =   285
         Left            =   1590
         TabIndex        =   10
         Top             =   1350
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Budget:"
         Height          =   405
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   1350
         Width           =   1545
      End
      Begin VB.Label lblPendingChanges 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   285
         Left            =   2370
         TabIndex        =   8
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label lblApprovedChanges 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "14,252.00"
         Height          =   285
         Left            =   1590
         TabIndex        =   7
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Approved Changes:"
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   6
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label lblEstimate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "352,450.00"
         Height          =   285
         Left            =   1590
         TabIndex        =   5
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Original Budget:"
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblContract 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "424,800.00"
         Height          =   285
         Left            =   1590
         TabIndex        =   3
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stuart Olsen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4635
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current Contract:"
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   1545
      End
   End
End
Attribute VB_Name = "CustomerPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ClickAttachments()
Public Event ClickContacts()

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
   ByVal hLibModule As Long) As Long

Private m_hMod As Long
Public Sub SetValues(Description As String, Contract As Double, Estimate As Double, ApprovedChanges As Double, PendingChanges As Double, JTDCost As Double)
On Error Resume Next
Exit Sub
    Dim Budget As Double
    Dim CostToComplete As Double
    
    Budget = Estimate + ApprovedChanges '+ PendingChanges
    CostToComplete = Budget - JTDCost
    
    lblDescription = Description
    lblContract = format(Contract, "#,##0.00")
    lblEstimate = format(Estimate, "#,##0.00")
    lblApprovedChanges = format(ApprovedChanges, "#,##0.00")
    lblPendingChanges = format(PendingChanges, "#,##0.00")
    lblBudget = format(Budget, "#,##0.00")
    lblJTDCost = format(JTDCost, "#,##0.00")
    lblCostToComplete = format(CostToComplete, "#,##0.00")
    
    If (Budget) = 0 Then
        lblJTDPercent = ""
    Else
        lblJTDPercent = "(" & Round(Val(JTDCost / Budget * 100), 1) & "%)"
    End If
    
End Sub
Public Property Get hWnd() As Long
On Error Resume Next
    hWnd = UserControl.hWnd
End Property

Public Property Get BackColor() As OLE_COLOR
On Error Resume Next
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(RHS As OLE_COLOR)
On Error Resume Next
    UserControl.BackColor = RHS
    Frame1.BackColor = RHS
End Property


Private Sub imgLink_Click(Index As Integer)
On Error Resume Next
    Call RunLink(Index)
End Sub
Private Sub lblLink_Click(Index As Integer)
    On Error Resume Next
    Call RunLink(Index)
End Sub
Private Sub RunLink(Index As Integer)
On Error Resume Next
    If Index = 0 Then
        RaiseEvent ClickContacts
    Else
        RaiseEvent ClickAttachments
    End If
End Sub

Private Sub UserControl_Initialize()
 '   m_hMod = LoadLibrary("shell32.dll")
 '   InitCommonControls
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    'BackColor = PropBag.ReadProperty("BackColor", UserControl.BackColor)
End Sub

Private Sub UserControl_Terminate()
Call SetErrorMode(SEM_NOERRORS)
'FreeLibrary m_hMod
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    'Call PropBag.WriteProperty("BackColor", UserControl.BackColor)
End Sub
