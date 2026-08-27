VERSION 5.00
Begin VB.Form FCommitment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FCommitment"
   ClientHeight    =   2685
   ClientLeft      =   7545
   ClientTop       =   3315
   ClientWidth     =   4155
   Icon            =   "FCommitment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1470
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Current Invoice"
      Height          =   195
      Index           =   0
      Left            =   1395
      TabIndex        =   9
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label lblCurrent 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   2610
      TabIndex        =   8
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label lbVarianceCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<CommitmentCaption> Remaining"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   1680
      Width           =   2370
   End
   Begin VB.Label lblScreenTitle 
      Caption         =   "This distribution will exceed the amount of the <CommitmentCaption> item."
      Height          =   435
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   3195
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2580
      X2              =   3540
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Label lblVariance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   2610
      TabIndex        =   5
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label lblActual 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   2610
      TabIndex        =   4
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label lblBudget 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "$999,999.99"
      Height          =   195
      Left            =   2610
      TabIndex        =   3
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Invoiced to date"
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   2
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Label lbBudgetCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<CommitmentCaption> amount"
      Height          =   195
      Left            =   330
      TabIndex        =   1
      Top             =   900
      Width           =   2145
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "FCommitment.frx":000C
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "FCommitment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNav_Click(Index As Integer)
    Unload Me
End Sub

Public Sub ShowCommitment(Actual As Double, Current As Double, Budget As Double)
    
    Me.caption = App.ProductName
    
    lblScreenTitle = "The invoiced amount will exceed the " & App.Options(Caption_Commitment) & " item amount."
                     
    lblBudget = Format(Budget, "#,##0.00")
    lblActual = Format(Actual, "#,##0.00")
    lblCurrent = Format(Current, "#,##0.00")
    
    lbBudgetCaption = App.Options(Caption_Commitment) & " amount"
    lbVarianceCaption = App.Options(Caption_Commitment) & " remaining"
    
    lblVariance = Format(Budget - Actual - Current, "#,##0.00")
    lblVariance.Font.Bold = Budget - Actual - Current < 0
    lblVariance.ForeColor = IIf(lblVariance.Font.Bold, vbRed, vbWindowText)
        
    Me.Show vbModal
    
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

