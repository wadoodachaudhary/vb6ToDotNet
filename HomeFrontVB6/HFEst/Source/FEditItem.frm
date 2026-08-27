VERSION 5.00
Begin VB.Form FEditItem 
   Caption         =   "Item"
   ClientHeight    =   5070
   ClientLeft      =   4785
   ClientTop       =   4485
   ClientWidth     =   4755
   Icon            =   "FEditItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   4755
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   3600
      TabIndex        =   21
      Top             =   4620
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   2520
      TabIndex        =   20
      Top             =   4620
      Width           =   1035
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3180
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3540
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   14
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   12
      Top             =   4200
      Width           =   435
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   10
      Top             =   3900
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtPOIndex 
      Height          =   285
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   3
      Top             =   420
      Width           =   2535
   End
   Begin VB.TextBox txtNotes 
      Height          =   1365
      Left            =   1380
      MaxLength       =   2000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1050
      Width           =   3075
   End
   Begin VB.ComboBox cboJCCostCode 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2460
      Width           =   3075
   End
   Begin VB.ComboBox cboJCCategory 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2820
      Width           =   3075
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "JC Cost Code"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   2490
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "JC Category"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   2850
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   15
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Rounding"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Conversion"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Phase"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   9
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Item"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   450
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Notes"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Order unit"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   3210
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Takeoff unit"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   3570
      Width           =   1215
   End
End
Attribute VB_Name = "FEditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

