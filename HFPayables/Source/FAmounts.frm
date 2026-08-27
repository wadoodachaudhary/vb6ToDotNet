VERSION 5.00
Begin VB.Form FAmounts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FAmounts"
   ClientHeight    =   4935
   ClientLeft      =   8235
   ClientTop       =   390
   ClientWidth     =   5670
   Icon            =   "FAmounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmUnitValues 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   0
      TabIndex        =   26
      Top             =   2460
      Width           =   5535
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Committed"
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
         Left            =   750
         TabIndex        =   36
         Top             =   660
         Width           =   2760
      End
      Begin VB.Label lblCommittedQuantity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   2610
         TabIndex        =   35
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblCommitedTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   2610
         TabIndex        =   34
         Top             =   1500
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   2580
         X2              =   3540
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblCommittedPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   2610
         TabIndex        =   33
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unit Price"
         Height          =   195
         Index           =   8
         Left            =   1575
         TabIndex        =   32
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label lblInvoicedPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   3780
         TabIndex        =   31
         Top             =   1200
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   3750
         X2              =   4710
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblInvoicedTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   3780
         TabIndex        =   30
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label lblInvoicedQuantity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   3780
         TabIndex        =   29
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   28
         Top             =   960
         Width           =   585
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   270
         Picture         =   "FAmounts.frx":000C
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Unit values on this distribution exceed the committed amounts."
         Height          =   255
         Left            =   930
         TabIndex        =   27
         Top             =   240
         Width           =   4485
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoiced"
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
         Left            =   720
         TabIndex        =   37
         Top             =   660
         Width           =   3960
      End
   End
   Begin VB.Frame frmBudget 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      Begin VB.Image Image3 
         Height          =   480
         Left            =   270
         Picture         =   "FAmounts.frx":08D6
         Top             =   180
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblScreenTitle 
         Caption         =   "This distribution will exceed the job's budgeted amount."
         Height          =   495
         Left            =   930
         TabIndex        =   23
         Top             =   240
         Width           =   5115
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   270
         Picture         =   "FAmounts.frx":11A0
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Budgeted Amount"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   22
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Invoiced to date"
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   21
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblPhaseAmount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   3180
         TabIndex        =   20
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblCategoryAmount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   4380
         TabIndex        =   19
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblPhaseInvoiced 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   3180
         TabIndex        =   18
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblCategoryInvoiced 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   4380
         TabIndex        =   17
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblPhaseRemaining 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   3180
         TabIndex        =   16
         Top             =   2100
         Width           =   900
      End
      Begin VB.Label lblCategoryRemaining 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   4380
         TabIndex        =   15
         Top             =   2100
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4350
         X2              =   5310
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3150
         X2              =   4110
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Remaining"
         Height          =   195
         Index           =   1
         Left            =   915
         TabIndex        =   14
         Top             =   2100
         Width           =   750
      End
      Begin VB.Label lblCategoryCurrent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   4380
         TabIndex        =   13
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblPhaseCurrent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   3180
         TabIndex        =   12
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Current Invoice"
         Height          =   195
         Index           =   2
         Left            =   585
         TabIndex        =   11
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label lblCommitmentCurrent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   2010
         TabIndex        =   10
         Top             =   1320
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1980
         X2              =   2940
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblCommitmentRemaining 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   2010
         TabIndex        =   9
         Top             =   2100
         Width           =   900
      End
      Begin VB.Label lblCommitmentInvoiced 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   2010
         TabIndex        =   8
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblCommitmentAmount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   2010
         TabIndex        =   7
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblCommitmentTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "<Commitment>"
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
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   2760
      End
      Begin VB.Label lblCommitmentHold 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   2010
         TabIndex        =   5
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lblCategoryHold 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   4380
         TabIndex        =   4
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lblPhaseHold 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$999,999.99"
         Height          =   195
         Left            =   3180
         TabIndex        =   3
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Amount on Hold"
         Height          =   195
         Index           =   4
         Left            =   525
         TabIndex        =   2
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label lblPhaseTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "<Phase>"
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
         Left            =   120
         TabIndex        =   24
         Top             =   780
         Width           =   3960
      End
      Begin VB.Label lblCategoryTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "<Category>"
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
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Width           =   5160
      End
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4380
      TabIndex        =   0
      Top             =   4500
      Width           =   1215
   End
End
Attribute VB_Name = "FAmounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNav_Click(Index As Integer)
    Unload Me
End Sub

Public Sub ShowAmounts(AlwaysShow As Boolean, PhaseDesc As String, CategoryDesc As String, _
                       CommitmentAmount As Double, CommitmentCurrent As Double, CommitmentHold As Double, CommitmentInvoiced As Double, _
                       PhaseAmount As Double, PhaseCurrent As Double, PhaseHold As Double, PhaseInvoiced As Double, _
                       CategoryAmount As Double, CategoryCurrent As Double, CategoryHold As Double, CategoryInvoiced As Double, _
                       CommittedQuantity As Double, CommittedPrice As Double, _
                       InvoicedQuantity As Double, InvoicedPrice As Double)
                       
    Dim bShowBudgets    As Boolean
    Dim bShowUnitValues As Boolean
    Dim Remaining As Double
    Me.Caption = App.ProductName
    
    If ((CommitmentAmount <> 0) And (CommitmentAmount - CommitmentCurrent - CommitmentInvoiced - CommitmentHold < 0)) _
        Or ((PhaseAmount <> 0) And (PhaseAmount - PhaseCurrent - PhaseInvoiced - PhaseHold < 0)) _
        Or ((CategoryAmount <> 0) And (CategoryAmount - CategoryCurrent - CategoryInvoiced - CategoryHold < 0)) _
        Or (App.Options(ShowQtyAndUnitPrice) And InvoicedQuantity > CommittedQuantity) _
        Or (App.Options(ShowQtyAndUnitPrice) And InvoicedPrice > CommittedPrice) Then
        
        Image1.Visible = True
        Image3.Visible = False
        lblScreenTitle = "This distribution will exceed the " & App.Options(Caption_Job) & "'s budgeted amount for" & vbCrLf & PhaseDesc & " - " & CategoryDesc
    Else
        Image1.Visible = False
        Image3.Visible = True
        lblScreenTitle = PhaseDesc & " - " & CategoryDesc
    End If
    
    lblCommitmentTitle = App.Options(Caption_Commitment)
    lblCommitmentAmount = Format(CommitmentAmount, "#,##0.00")
    lblCommitmentCurrent = Format(CommitmentCurrent, "#,##0.00")
    lblCommitmentHold = Format(CommitmentHold, "#,##0.00")
    lblCommitmentInvoiced = Format(CommitmentInvoiced, "#,##0.00")
    Remaining = CommitmentAmount - CommitmentCurrent - CommitmentInvoiced - CommitmentHold
    lblCommitmentRemaining = Format(Remaining, "#,##0.00")
    lblCommitmentRemaining.Font.Bold = Remaining < 0
    lblCommitmentRemaining.ForeColor = IIf(Remaining < 0, vbRed, vbWindowText)
    
    lblPhaseTitle = App.Options(Caption_Phase)
    lblPhaseAmount = Format(PhaseAmount, "#,##0.00")
    lblPhaseCurrent = Format(PhaseCurrent, "#,##0.00")
    lblPhaseHold = Format(PhaseHold, "#,##0.00")
    lblPhaseInvoiced = Format(PhaseInvoiced, "#,##0.00")
    Remaining = PhaseAmount - PhaseCurrent - PhaseInvoiced - PhaseHold
    lblPhaseRemaining = Format(Remaining, "#,##0.00")
    lblPhaseRemaining.Font.Bold = Remaining < 0 And PhaseAmount <> 0 And Remaining < PhaseAmount
    lblPhaseRemaining.ForeColor = IIf(Remaining < 0 And PhaseAmount <> 0, vbRed, vbWindowText)
    
    lblCategoryTitle = App.Options(Caption_Category)
    lblCategoryAmount = Format(CategoryAmount, "#,##0.00")
    lblCategoryCurrent = Format(CategoryCurrent, "#,##0.00")
    lblCategoryHold = Format(CategoryHold, "#,##0.00")
    lblCategoryInvoiced = Format(CategoryInvoiced, "#,##0.00")
    Remaining = CategoryAmount - CategoryCurrent - CategoryInvoiced - CategoryHold
    lblCategoryRemaining = Format(Remaining, "#,##0.00")
    lblCategoryRemaining.Font.Bold = Remaining < 0 And CategoryAmount <> 0 And Remaining < CategoryAmount
    lblCategoryRemaining.ForeColor = IIf(Remaining < 0 And CategoryAmount <> 0, vbRed, vbWindowText)
        
    
    
    
    lblCommittedQuantity = CommittedQuantity
    lblCommittedPrice = Format(CommittedPrice, "#,##0.00")
    lblCommitedTotal = Format(CommittedQuantity * CommittedPrice, "#,##0.00")
    
    lblInvoicedQuantity = InvoicedQuantity
    lblInvoicedQuantity.ForeColor = IIf(InvoicedQuantity > CommittedQuantity And CommittedQuantity <> 0, vbRed, vbWindowText)
    lblInvoicedQuantity.FontBold = InvoicedQuantity > CommittedQuantity And CommittedQuantity <> 0
    lblInvoicedPrice = Format(InvoicedPrice, "#,##0.00")
    lblInvoicedPrice.ForeColor = IIf(InvoicedPrice > CommittedPrice And CommittedPrice <> 0, vbRed, vbWindowText)
    lblInvoicedPrice.FontBold = InvoicedPrice > CommittedPrice And CommittedPrice <> 0
    lblInvoicedTotal = Format(InvoicedQuantity * InvoicedPrice, "#,##0.00")
        
        
        
    
    bShowUnitValues = (App.Options(ShowQtyAndUnitPrice) And AlwaysShow) Or InvoicedQuantity > CommittedQuantity Or InvoicedPrice > CommittedPrice
    
    bShowBudgets = AlwaysShow Or CommitmentAmount - CommitmentCurrent - CommitmentInvoiced - CommitmentHold < 0 _
                    Or PhaseAmount - PhaseCurrent - PhaseInvoiced - PhaseHold < 0 _
                    Or CategoryAmount - CategoryCurrent - CategoryInvoiced - CategoryHold < 0
        
        
    frmBudget.Visible = bShowBudgets
    frmUnitValues.Visible = bShowUnitValues
    Me.Height = 900 + IIf(bShowBudgets, frmBudget.Height, 0) + IIf(bShowUnitValues, frmUnitValues.Height, 0) + 240
    frmUnitValues.Top = IIf(bShowBudgets, frmBudget.Height, 0)
    cmdNav(0).Move Me.ScaleWidth - cmdNav(0).Width - 120, Me.ScaleHeight - cmdNav(0).Height - 120
        
    Screen.MousePointer = vbDefault
    Me.Show vbModal
    Screen.MousePointer = vbHourglass
    
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

