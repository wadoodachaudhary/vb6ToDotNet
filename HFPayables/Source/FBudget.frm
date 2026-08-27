VERSION 5.00
Begin VB.Form FBudget 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FBudget"
   ClientHeight    =   2685
   ClientLeft      =   1530
   ClientTop       =   3255
   ClientWidth     =   4890
   Icon            =   "FBudget.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1838
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Current Invoice"
      Height          =   195
      Index           =   2
      Left            =   1035
      TabIndex        =   15
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lblPhaseCurrent 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   2250
      TabIndex        =   14
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lblCategoryCurrent 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   3450
      TabIndex        =   13
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Budget Remaining"
      Height          =   195
      Index           =   1
      Left            =   810
      TabIndex        =   12
      Top             =   1740
      Width           =   1305
   End
   Begin VB.Label lblScreenTitle 
      Caption         =   "This distribution will exceed the job's budgeted amount."
      Height          =   435
      Left            =   840
      TabIndex        =   11
      Top             =   240
      Width           =   3975
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2220
      X2              =   3180
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3420
      X2              =   4380
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblCategoryVariance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   3450
      TabIndex        =   10
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label lblPhaseVariance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   2250
      TabIndex        =   9
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label lblCategoryActual 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   3450
      TabIndex        =   8
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label lblPhaseActual 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   2250
      TabIndex        =   7
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label lblCategoryBudget 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   3450
      TabIndex        =   6
      Top             =   960
      Width           =   900
   End
   Begin VB.Label lblPhaseBudget 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   2250
      TabIndex        =   5
      Top             =   960
      Width           =   900
   End
   Begin VB.Label lblCategoryTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3150
      TabIndex        =   4
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label lblPhaseTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Phase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   990
      TabIndex        =   3
      Top             =   720
      Width           =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Invoiced to date"
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Budget"
      Height          =   195
      Index           =   0
      Left            =   1605
      TabIndex        =   1
      Top             =   960
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "FBudget.frx":000C
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "FBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNav_Click(Index As Integer)
    Unload Me
End Sub

Public Sub ShowBudget(PActual As Double, pCurrent As Double, PBudget As Double, _
                      CActual As Double, CCurrent As Double, CBudget As Double)
    
    Me.caption = App.ProductName
    
    lblScreenTitle = "This distribution will exceed the " & App.Options(Caption_Job) & "'s budgeted amount."
    
    lblPhaseTitle = App.Options(Caption_Phase)
    lblPhaseActual = Format(PActual, "#,##0.00")
    lblPhaseCurrent = Format(pCurrent, "#,##0.00")
    lblPhaseBudget = Format(PBudget, "#,##0.00")
    lblPhaseVariance = Format(PActual + pCurrent - PBudget, "#,##0.00")
    lblPhaseVariance.Font.Bold = PActual + pCurrent - PBudget > 0
    
    lblCategoryTitle = App.Options(Caption_Category)
    lblCategoryActual = Format(CActual, "#,##0.00")
    lblCategoryCurrent = Format(CCurrent, "#,##0.00")
    lblCategoryBudget = Format(CBudget, "#,##0.00")
    lblCategoryVariance = Format(CActual + CCurrent - CBudget, "#,##0.00")
    lblCategoryVariance.Font.Bold = CActual + CCurrent - CBudget > 0
        
    Me.Show vbModal
    
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

