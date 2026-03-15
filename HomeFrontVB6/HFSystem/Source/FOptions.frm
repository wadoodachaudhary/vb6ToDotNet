VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6C83CF2C-BE8D-4EE2-9B07-7FF5E27AB2FD}#1.0#0"; "zybCombo.ocx"
Begin VB.Form FOptions 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Settings"
   ClientHeight    =   7815
   ClientLeft      =   6870
   ClientTop       =   2700
   ClientWidth     =   9405
   ForeColor       =   &H80000006&
   Icon            =   "FOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   521
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   ShowInTaskbar   =   0   'False
   Begin VB.Frame TabFrame 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Accounting Integration"
      Height          =   12075
      Index           =   4
      Left            =   4770
      TabIndex        =   172
      Tag             =   "Security"
      Top             =   1110
      Visible         =   0   'False
      Width           =   23145
      Begin VB.Frame AccountingFrame 
         Caption         =   "Quickbooks Online"
         Height          =   1425
         Index           =   6
         Left            =   13590
         TabIndex        =   361
         Top             =   6180
         Width           =   6375
         Begin VB.OptionButton optQBOJobHeirarchy 
            Caption         =   "Hierarchical - create customer records and sub jobs"
            Height          =   225
            Index           =   1
            Left            =   2265
            TabIndex        =   363
            Top             =   660
            Value           =   -1  'True
            Width           =   4035
         End
         Begin VB.OptionButton optQBOJobHeirarchy 
            Caption         =   "Simple - create the customer record only"
            Height          =   225
            Index           =   0
            Left            =   2265
            TabIndex        =   362
            Top             =   420
            Width           =   4035
         End
         Begin HFSystem.VBCombo cboQuickBooksOnlineVersion 
            Height          =   240
            Left            =   2265
            TabIndex        =   364
            Top             =   45
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   423
            Style           =   2
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Edition"
            Height          =   195
            Index           =   121
            Left            =   1710
            TabIndex        =   366
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Job Creation"
            Height          =   195
            Index           =   120
            Left            =   1305
            TabIndex        =   365
            Top             =   420
            Width           =   885
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   15
            Left            =   390
            Picture         =   "FOptions.frx":000C
            Stretch         =   -1  'True
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "required"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4890
         TabIndex        =   37
         Top             =   6180
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox check2 
         Alignment       =   1  'Right Justify
         Caption         =   "required"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4890
         TabIndex        =   38
         Top             =   6405
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chkUseVarianceReporting 
         Alignment       =   1  'Right Justify
         Caption         =   "required"
         Height          =   195
         Left            =   4890
         TabIndex        =   39
         Top             =   6630
         Width           =   1020
      End
      Begin VB.Frame AccountingFrame 
         Caption         =   "Xero"
         Height          =   1425
         Index           =   7
         Left            =   13575
         TabIndex        =   302
         Top             =   315
         Width           =   6375
         Begin VB.Image Image1 
            Height          =   480
            Index           =   10
            Left            =   390
            Picture         =   "FOptions.frx":08D6
            Stretch         =   -1  'True
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.TextBox txtPOSegment 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   8
         Left            =   2325
         TabIndex        =   26
         Text            =   "2"
         Top             =   5310
         Width           =   255
      End
      Begin VB.TextBox txtPOSegment 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   2175
         TabIndex        =   25
         Text            =   "/"
         Top             =   5310
         Width           =   135
      End
      Begin VB.OptionButton optPOByJobPOIndex 
         Caption         =   "Use job and PO specific numbering."
         Height          =   225
         Left            =   375
         TabIndex        =   21
         Top             =   5100
         Width           =   2955
      End
      Begin VB.TextBox txtPOSegment 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Index           =   4
         Left            =   1485
         TabIndex        =   22
         Text            =   "4"
         Top             =   5310
         Width           =   255
      End
      Begin VB.TextBox txtPOSegment 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   1905
         TabIndex        =   24
         Text            =   "4"
         Top             =   5310
         Width           =   255
      End
      Begin VB.TextBox txtPOSegment 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   1755
         TabIndex        =   23
         Text            =   "/"
         Top             =   5310
         Width           =   135
      End
      Begin VB.Frame AccountingFrame 
         Caption         =   "QuickBooks"
         Height          =   1425
         Index           =   2
         Left            =   6990
         TabIndex        =   244
         Top             =   3270
         Width           =   6375
         Begin VB.CheckBox chkQBPostRevenueToJob 
            Caption         =   "Post revenue to job"
            Height          =   195
            Left            =   2580
            TabIndex        =   96
            Top             =   1080
            Width           =   3135
         End
         Begin VB.OptionButton optQBJobHeirarchy 
            Caption         =   "Simple - create the customer record only"
            Height          =   225
            Index           =   0
            Left            =   2265
            TabIndex        =   94
            Top             =   570
            Width           =   4035
         End
         Begin VB.OptionButton optQBJobHeirarchy 
            Caption         =   "Hierarchical - create customer records and sub jobs"
            Height          =   225
            Index           =   1
            Left            =   2265
            TabIndex        =   95
            Top             =   810
            Value           =   -1  'True
            Width           =   4035
         End
         Begin VB.TextBox txtQuickBooksFile 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2280
            TabIndex        =   92
            Top             =   0
            Width           =   3615
         End
         Begin HFSystem.VBCombo cboQuickBooksVersion 
            Height          =   240
            Left            =   2280
            TabIndex        =   93
            Top             =   255
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   423
            Style           =   2
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Job Creation"
            Height          =   195
            Index           =   65
            Left            =   1305
            TabIndex        =   288
            Top             =   570
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Edition"
            Height          =   195
            Index           =   25
            Left            =   1710
            TabIndex        =   259
            Top             =   270
            Width           =   480
         End
         Begin VB.Image cmdChooseFile 
            Height          =   240
            Index           =   0
            Left            =   5925
            Picture         =   "FOptions.frx":11A0
            Top             =   0
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   390
            Picture         =   "FOptions.frx":172A
            Stretch         =   -1  'True
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Company File"
            Height          =   195
            Index           =   54
            Left            =   1035
            TabIndex        =   245
            Top             =   0
            Width           =   1155
         End
      End
      Begin VB.Frame AccountingFrame 
         Caption         =   "Sage 50 Accounting"
         Height          =   1425
         Index           =   3
         Left            =   6990
         TabIndex        =   265
         Top             =   1800
         Width           =   6375
         Begin VB.TextBox txtSimplyPswd 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   2265
            PasswordChar    =   "*"
            TabIndex        =   90
            Top             =   510
            Width           =   1755
         End
         Begin VB.TextBox txtSimplyUser 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2265
            TabIndex        =   89
            Top             =   255
            Width           =   1755
         End
         Begin VB.TextBox txtSimplyFile 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2265
            TabIndex        =   88
            Top             =   0
            Width           =   3615
         End
         Begin HFSystem.VBCombo cboSimplyInternalCustomer 
            Height          =   240
            Left            =   2265
            TabIndex        =   91
            Top             =   765
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   423
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Internal Customer"
            Height          =   195
            Index           =   63
            Left            =   960
            TabIndex        =   275
            Top             =   795
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Password"
            Height          =   195
            Index           =   55
            Left            =   1035
            TabIndex        =   268
            Top             =   525
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Security Context"
            Height          =   195
            Index           =   33
            Left            =   1035
            TabIndex        =   267
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Company File"
            Height          =   195
            Index           =   44
            Left            =   1035
            TabIndex        =   266
            Top             =   0
            Width           =   1155
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   390
            Picture         =   "FOptions.frx":1FF4
            Stretch         =   -1  'True
            Top             =   240
            Width           =   480
         End
         Begin VB.Image cmdChooseFile 
            Height          =   240
            Index           =   2
            Left            =   5925
            Picture         =   "FOptions.frx":28BE
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.TextBox txtNextPO 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "8000"
         Top             =   5955
         Width           =   1215
      End
      Begin VB.TextBox txtPOSegment 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   1755
         TabIndex        =   19
         Text            =   "/"
         Top             =   4725
         Width           =   135
      End
      Begin VB.TextBox txtPOSegment 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   1905
         TabIndex        =   20
         Text            =   "3"
         Top             =   4725
         Width           =   255
      End
      Begin VB.TextBox txtPOSegment 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   1485
         TabIndex        =   18
         Text            =   "8"
         Top             =   4725
         Width           =   255
      End
      Begin VB.TextBox txtPOPrefix 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1485
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "8000"
         Top             =   6210
         Width           =   525
      End
      Begin VB.OptionButton optPOByJob 
         Caption         =   "Use job specific numbering."
         Height          =   225
         Left            =   375
         TabIndex        =   17
         Top             =   4455
         Value           =   -1  'True
         Width           =   2565
      End
      Begin VB.OptionButton optPOSequential 
         Caption         =   "Use sequential numbering"
         Height          =   255
         Left            =   345
         TabIndex        =   27
         Top             =   5640
         Width           =   2535
      End
      Begin VB.CheckBox chkUseTaxGroups 
         Alignment       =   1  'Right Justify
         Caption         =   "required"
         Height          =   195
         Left            =   4890
         TabIndex        =   36
         Top             =   5955
         Width           =   1020
      End
      Begin HFSystem.VBCombo cboAccountingSystem 
         Height          =   240
         Left            =   2265
         TabIndex        =   1
         Top             =   675
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   423
         Style           =   2
      End
      Begin HFSystem.VBCombo cboTaxGroup 
         Height          =   240
         Index           =   0
         Left            =   5340
         TabIndex        =   30
         Top             =   4365
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         Style           =   2
      End
      Begin HFSystem.VBCombo cboTaxGroup 
         Height          =   240
         Index           =   1
         Left            =   5340
         TabIndex        =   31
         Top             =   4620
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         Style           =   2
      End
      Begin HFSystem.VBCombo cboTaxGroup 
         Height          =   240
         Index           =   2
         Left            =   5340
         TabIndex        =   32
         Top             =   4875
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         Style           =   2
      End
      Begin HFSystem.VBCombo cboTaxGroup 
         Height          =   240
         Index           =   3
         Left            =   5340
         TabIndex        =   33
         Top             =   5130
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         Style           =   2
      End
      Begin HFSystem.VBCombo cboTaxGroup 
         Height          =   240
         Index           =   4
         Left            =   5340
         TabIndex        =   34
         Top             =   5385
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         Style           =   2
      End
      Begin HFSystem.VBCombo cboTaxGroup 
         Height          =   240
         Index           =   5
         Left            =   5340
         TabIndex        =   35
         Top             =   5640
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.Frame AccountingFrame 
         Caption         =   "Sage 100 Contractor"
         Height          =   1425
         Index           =   5
         Left            =   6990
         TabIndex        =   85
         Top             =   330
         Width           =   6375
         Begin VB.CheckBox chkmbUseSubAcct 
            Caption         =   "Use Job as Sub Account"
            Height          =   195
            Left            =   2925
            TabIndex        =   87
            Top             =   1080
            Width           =   2460
         End
         Begin VB.TextBox txtMBUser 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2265
            TabIndex        =   83
            Top             =   510
            Width           =   1755
         End
         Begin VB.TextBox txtMBPswd 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   2265
            PasswordChar    =   "*"
            TabIndex        =   84
            Top             =   765
            Width           =   1755
         End
         Begin HFSystem.VBCombo cboMBCompany 
            Height          =   240
            Left            =   2265
            TabIndex        =   82
            Top             =   255
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   423
         End
         Begin HFSystem.VBCombo cboMBAPILevel 
            Height          =   240
            Left            =   4995
            TabIndex        =   86
            Top             =   630
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   423
            Style           =   2
         End
         Begin HFSystem.VBCombo cboMBDrive 
            Height          =   240
            Left            =   2265
            TabIndex        =   80
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   423
            Style           =   2
         End
         Begin HFSystem.VBCombo cboMBInstance 
            Height          =   240
            Left            =   2265
            TabIndex        =   81
            Top             =   0
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   423
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "API Level"
            Height          =   255
            Index           =   81
            Left            =   4155
            TabIndex        =   306
            Top             =   630
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Company"
            Height          =   195
            Index           =   2
            Left            =   1035
            TabIndex        =   196
            Top             =   -30
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Security Context"
            Height          =   195
            Index           =   20
            Left            =   1035
            TabIndex        =   195
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Password"
            Height          =   195
            Index           =   19
            Left            =   1035
            TabIndex        =   194
            Top             =   720
            Width           =   1155
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   390
            Picture         =   "FOptions.frx":2E48
            Stretch         =   -1  'True
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame AccountingFrame 
         Caption         =   "Sage 300 Construction and Real Estate"
         Height          =   1425
         Index           =   1
         Left            =   13590
         TabIndex        =   189
         Top             =   4710
         Width           =   6375
         Begin VB.TextBox txtTLARFolder 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2265
            TabIndex        =   77
            Top             =   255
            Width           =   3615
         End
         Begin VB.TextBox txtTLPswd 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   2265
            PasswordChar    =   "*"
            TabIndex        =   79
            Top             =   765
            Width           =   1755
         End
         Begin VB.TextBox txtTLUser 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2265
            TabIndex        =   78
            Text            =   " "
            Top             =   504
            Width           =   1755
         End
         Begin VB.TextBox txtTLFolder 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2265
            TabIndex        =   76
            Top             =   0
            Width           =   3615
         End
         Begin VB.Image cmdChooseFolder 
            Height          =   240
            Index           =   1
            Left            =   5940
            Picture         =   "FOptions.frx":3C8A
            Top             =   255
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "AR Data Folder"
            Height          =   195
            Index           =   37
            Left            =   1035
            TabIndex        =   197
            Top             =   255
            Width           =   1155
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   390
            Picture         =   "FOptions.frx":4214
            Stretch         =   -1  'True
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Password"
            Height          =   192
            Index           =   12
            Left            =   1032
            TabIndex        =   192
            Top             =   756
            Width           =   1152
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Security Context"
            Height          =   192
            Index           =   13
            Left            =   1032
            TabIndex        =   191
            Top             =   504
            Width           =   1152
         End
         Begin VB.Image cmdChooseFolder 
            Height          =   240
            Index           =   0
            Left            =   5940
            Picture         =   "FOptions.frx":4ADE
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Folder"
            Height          =   195
            Index           =   11
            Left            =   1035
            TabIndex        =   190
            Top             =   0
            Width           =   1155
         End
      End
      Begin HFSystem.VBCombo cboVarianceCat 
         Height          =   240
         Left            =   4710
         TabIndex        =   40
         Top             =   6870
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.Frame AccountingFrame 
         Caption         =   "Sage Intacct"
         Height          =   2325
         Index           =   4
         Left            =   3510
         TabIndex        =   341
         Top             =   8820
         Width           =   6375
         Begin HFSystem.VBCombo cboIntacctPOCOType 
            Height          =   240
            Left            =   4530
            TabIndex        =   12
            Top             =   1035
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   423
            Style           =   2
            Text            =   "Combo1"
         End
         Begin VB.TextBox txtIntacctBudgetGL 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1575
            TabIndex        =   8
            Top             =   1530
            Width           =   1830
         End
         Begin VB.TextBox txtIntacctCompanyID 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2265
            TabIndex        =   2
            Top             =   0
            Width           =   3615
         End
         Begin VB.TextBox txtIntacctUID 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2265
            TabIndex        =   3
            Top             =   255
            Width           =   1755
         End
         Begin VB.TextBox txtIntacctPWD 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   2265
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   510
            Width           =   1755
         End
         Begin HFSystem.VBCombo cboIntacctPOType 
            Height          =   240
            Left            =   4530
            TabIndex        =   11
            Top             =   780
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboIntacctSubContractType 
            Height          =   240
            Left            =   4530
            TabIndex        =   13
            Top             =   1290
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboIntacctInvoiceType 
            Height          =   240
            Left            =   4530
            TabIndex        =   15
            Top             =   1800
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboIntacctItem 
            Height          =   240
            Index           =   0
            Left            =   1575
            TabIndex        =   6
            Top             =   1020
            Width           =   1830
            _ExtentX        =   1455
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboIntacctItem 
            Height          =   240
            Index           =   1
            Left            =   1575
            TabIndex        =   7
            Top             =   1275
            Width           =   1830
            _ExtentX        =   1455
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboIntacctEntity 
            Height          =   240
            Left            =   1575
            TabIndex        =   5
            Top             =   765
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboIntacctEstimateType 
            Height          =   240
            Left            =   4530
            TabIndex        =   16
            Top             =   2055
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboIntacctSubContractCOType 
            Height          =   240
            Left            =   4515
            TabIndex        =   14
            Top             =   1545
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboPOReferenceFld 
            Height          =   240
            Left            =   1575
            TabIndex        =   9
            Top             =   1785
            Width           =   1830
            _ExtentX        =   1455
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin HFSystem.VBCombo cboInvReferenceFld 
            Height          =   240
            Left            =   1575
            TabIndex        =   10
            Top             =   2040
            Width           =   1830
            _ExtentX        =   1455
            _ExtentY        =   423
            DropDownWidth   =   300
            Style           =   2
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inv Reference Fld"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   125
            Left            =   225
            TabIndex        =   391
            Top             =   2070
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PO Reference Fld"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   124
            Left            =   225
            TabIndex        =   390
            Top             =   1815
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Contract CO"
            Height          =   195
            Index           =   82
            Left            =   3585
            TabIndex        =   385
            Top             =   1560
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PO CO"
            Height          =   195
            Index           =   80
            Left            =   3975
            TabIndex        =   384
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GST ItemID"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   119
            Left            =   660
            TabIndex        =   359
            Top             =   1305
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Budget GL"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   114
            Left            =   735
            TabIndex        =   354
            Top             =   1575
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Estimate"
            Height          =   195
            Index           =   113
            Left            =   3855
            TabIndex        =   353
            Top             =   2085
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Entity"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   24
            Left            =   1125
            TabIndex        =   350
            Top             =   795
            Width           =   390
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "One Time ItemID"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   111
            Left            =   315
            TabIndex        =   349
            Top             =   1050
            Width           =   1200
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Invoice"
            Height          =   195
            Index           =   110
            Left            =   3915
            TabIndex        =   348
            Top             =   1830
            Width           =   525
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Contract"
            Height          =   195
            Index           =   109
            Left            =   3855
            TabIndex        =   347
            Top             =   1305
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PO"
            Height          =   195
            Index           =   108
            Left            =   4230
            TabIndex        =   346
            Top             =   795
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Transaction Types"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   1
            Left            =   4530
            TabIndex        =   345
            Top             =   555
            Width           =   1320
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   390
            Picture         =   "FOptions.frx":5068
            Stretch         =   -1  'True
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Company ID"
            Height          =   195
            Index           =   35
            Left            =   1035
            TabIndex        =   344
            Top             =   0
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "User ID"
            Height          =   195
            Index           =   34
            Left            =   1035
            TabIndex        =   343
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Index           =   4
            Left            =   1500
            TabIndex        =   342
            Top             =   525
            Width           =   690
         End
      End
      Begin HFSystem.VBCombo cboBookOfAccount 
         Height          =   240
         Left            =   2265
         TabIndex        =   0
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   423
         DropDownWidth   =   300
         Style           =   2
      End
      Begin VB.Label lblEditList 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Wallet Configuration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   11
         Left            =   750
         TabIndex        =   367
         Top             =   450
         Width           =   1425
      End
      Begin VB.Label lblEditList 
         AutoSize        =   -1  'True
         Caption         =   "Variance Categories"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   3345
         TabIndex        =   351
         Top             =   6630
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Default category"
         Height          =   195
         Index           =   77
         Left            =   3480
         TabIndex        =   290
         Top             =   6885
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Format"
         Height          =   195
         Index           =   73
         Left            =   900
         TabIndex        =   289
         Top             =   5310
         Width           =   480
      End
      Begin VB.Image cmdNextPO 
         Height          =   240
         Left            =   2700
         Picture         =   "FOptions.frx":5932
         Top             =   5970
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "        Cost Codes && Taxes "
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
         Index           =   16
         Left            =   3120
         TabIndex        =   218
         Top             =   4155
         Width           =   2250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PO Numbering "
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
         Index           =   8
         Left            =   225
         TabIndex        =   221
         Top             =   4155
         Width           =   1290
      End
      Begin VB.Label lblNextPO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Next Num"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   660
         TabIndex        =   223
         Top             =   5970
         Width           =   705
      End
      Begin VB.Label lblPONumberSample 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AAAAAAAAAAAAAAAA/999"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   930
         TabIndex        =   166
         Top             =   6750
         Width           =   2040
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Format"
         Height          =   195
         Index           =   17
         Left            =   900
         TabIndex        =   222
         Top             =   4725
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   16
         X1              =   705
         X2              =   6105
         Y1              =   4245
         Y2              =   4245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   17
         X1              =   705
         X2              =   6105
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Label lblPOPrefix 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prefix"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   975
         TabIndex        =   220
         Top             =   6225
         Width           =   390
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sample"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   195
         TabIndex        =   219
         Top             =   6750
         Width           =   555
      End
      Begin VB.Label lblEditList 
         AutoSize        =   -1  'True
         Caption         =   "Tax groups..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   3825
         TabIndex        =   217
         Top             =   5955
         Width           =   930
      End
      Begin VB.Label lblEditList 
         AutoSize        =   -1  'True
         Caption         =   "Categories..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   3825
         TabIndex        =   216
         Top             =   6405
         Width           =   885
      End
      Begin VB.Label lblEditList 
         AutoSize        =   -1  'True
         Caption         =   "Cost codes..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   215
         Top             =   6180
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Other Tax Group"
         Height          =   195
         Index           =   10
         Left            =   4065
         TabIndex        =   211
         Top             =   5685
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Overhead Tax Group"
         Height          =   195
         Index           =   8
         Left            =   3750
         TabIndex        =   210
         Top             =   5430
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Equipment Tax Group"
         Height          =   195
         Index           =   7
         Left            =   3705
         TabIndex        =   209
         Top             =   5175
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Subcontract Tax Group"
         Height          =   195
         Index           =   6
         Left            =   3585
         TabIndex        =   208
         Top             =   4920
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Material Tax Group"
         Height          =   195
         Index           =   5
         Left            =   3900
         TabIndex        =   207
         Top             =   4665
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Labour Tax Group"
         Height          =   195
         Index           =   9
         Left            =   3945
         TabIndex        =   206
         Top             =   4410
         Width           =   1290
      End
      Begin VB.Label lblEditList 
         AutoSize        =   -1  'True
         Caption         =   "Change taxes on current jobs..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   8430
         TabIndex        =   205
         Top             =   6855
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "System"
         Height          =   195
         Index           =   3
         Left            =   1035
         TabIndex        =   193
         Top             =   705
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Accounting System "
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
         Index           =   12
         Left            =   240
         TabIndex        =   188
         Top             =   165
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   25
         X1              =   720
         X2              =   6120
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   24
         X1              =   720
         X2              =   6120
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   870
         TabIndex        =   224
         Top             =   6690
         Width           =   2220
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Purchasing Options"
      Height          =   7185
      Index           =   6
      Left            =   585
      TabIndex        =   173
      Tag             =   "Security"
      Top             =   6600
      Visible         =   0   'False
      Width           =   6795
      Begin VB.CheckBox chkPostPOQtyToAccounting 
         Caption         =   "Post PO order quantity to accounting."
         Height          =   195
         Left            =   420
         TabIndex        =   400
         Top             =   2145
         Width           =   5715
      End
      Begin VB.OptionButton optUseAltCostCodesForCO 
         Caption         =   "Only on change orders"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   399
         Top             =   5100
         Width           =   2040
      End
      Begin VB.OptionButton optUseAltCostCodesForCO 
         Caption         =   "On all options"
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   398
         Top             =   4860
         Value           =   -1  'True
         Width           =   2190
      End
      Begin VB.CheckBox chkPostSummarizedBudgets 
         Caption         =   "Post budgets summarized by job, cost code and category."
         Height          =   195
         Left            =   420
         TabIndex        =   394
         Top             =   2385
         Width           =   5535
      End
      Begin VB.CheckBox chkPostAssembliesAsSalesInvoices 
         Caption         =   "Post Precision Builder job assemblies as AR estimates."
         Height          =   195
         Left            =   420
         TabIndex        =   393
         Top             =   2610
         Width           =   5535
      End
      Begin VB.CheckBox chkPostJCExtraAsSubJob 
         Caption         =   "Post JC Extra as sub job"
         Height          =   195
         Left            =   420
         TabIndex        =   392
         Top             =   2835
         Width           =   5535
      End
      Begin VB.CheckBox chkUseComponents 
         Caption         =   "Use components in assembly definitions."
         Height          =   195
         Left            =   480
         TabIndex        =   315
         Top             =   1665
         Width           =   3195
      End
      Begin VB.CheckBox chkPurchasingRequiresSalesApproval 
         Caption         =   "Jobs require Sales approval before purchasing can begin."
         Height          =   195
         Left            =   480
         TabIndex        =   139
         Top             =   495
         Width           =   5535
      End
      Begin VB.CheckBox chkLetPurchaserChangeSaleQty 
         Caption         =   "Let purchasers change the sale quantity."
         Height          =   195
         Left            =   480
         TabIndex        =   143
         Top             =   1425
         Width           =   3195
      End
      Begin VB.CheckBox chkUsePricingFromEstimating 
         Caption         =   "When looking up prices don't use the item db. Use vendor pricelists only."
         Height          =   195
         Left            =   480
         TabIndex        =   146
         Top             =   3795
         Width           =   5715
      End
      Begin VB.CheckBox chkCanAddContractItemsFromPurchasing 
         Caption         =   "Let me add and remove contract items (model and option assemblies) on jobs."
         Height          =   195
         Left            =   480
         TabIndex        =   142
         Top             =   1185
         Width           =   6075
      End
      Begin VB.CheckBox chkZeroRateOnChangeVendor 
         Caption         =   "When the vendor changes set the item rate to zero if no value can be found."
         Height          =   195
         Left            =   480
         TabIndex        =   145
         Top             =   3570
         Width           =   5715
      End
      Begin VB.CheckBox chkZeroRateOnRefeshCosts 
         Caption         =   "When costs are refreshed set the item rate to zero if no value can be found."
         Height          =   195
         Left            =   480
         TabIndex        =   144
         Top             =   3345
         Width           =   5895
      End
      Begin VB.CheckBox chkWarnForecast 
         Caption         =   "Warn me when I use a forecasted cost basis to refresh PO costs."
         Enabled         =   0   'False
         Height          =   195
         Left            =   720
         TabIndex        =   141
         Top             =   945
         Width           =   5535
      End
      Begin VB.CheckBox chkAllowForecast 
         Caption         =   "Let me use a forecasted cost basis to refresh PO costs."
         Height          =   195
         Left            =   480
         TabIndex        =   140
         Top             =   720
         Width           =   5535
      End
      Begin VSFlex8Ctl.VSFlexGrid gFormatting 
         Height          =   960
         Left            =   510
         TabIndex        =   147
         Top             =   5670
         Width           =   4065
         _cx             =   1981946882
         _cy             =   1981941405
         Appearance      =   1
         BorderStyle     =   0
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FOptions.frx":5A7C
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
      Begin HFSystem.VBCombo cboZeroQtyTakeoffMode 
         Height          =   240
         Left            =   3525
         TabIndex        =   396
         Top             =   4365
         Width           =   2205
         _ExtentX        =   3228
         _ExtentY        =   423
         DropDownWidth   =   300
         Style           =   2
      End
      Begin VB.Label Label5 
         Caption         =   "When are the alternate cost codes and categories used?"
         Height          =   210
         Left            =   480
         TabIndex        =   397
         Top             =   4605
         Width           =   4110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Posting Options "
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
         Index           =   1
         Left            =   195
         TabIndex        =   395
         Top             =   1920
         Width           =   1410
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   675
         X2              =   6075
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   675
         X2              =   6075
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Conditional Formating "
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
         Index           =   10
         Left            =   210
         TabIndex        =   225
         Top             =   5430
         Width           =   1905
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   21
         X1              =   690
         X2              =   6090
         Y1              =   5550
         Y2              =   5550
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   20
         X1              =   690
         X2              =   6090
         Y1              =   5565
         Y2              =   5565
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Takeoff Handling "
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
         Index           =   5
         Left            =   240
         TabIndex        =   204
         Top             =   4125
         Width           =   1545
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   10
         X1              =   690
         X2              =   6090
         Y1              =   4245
         Y2              =   4245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   11
         X1              =   690
         X2              =   6090
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Label Label1 
         Caption         =   "How do you want to handle zero quantity takeoff items?"
         Height          =   225
         Index           =   18
         Left            =   465
         TabIndex        =   203
         Top             =   4365
         Width           =   2970
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Rate Retrieval "
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
         Index           =   9
         Left            =   240
         TabIndex        =   182
         Top             =   3105
         Width           =   1965
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   19
         X1              =   720
         X2              =   6120
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   18
         X1              =   720
         X2              =   6120
         Y1              =   3225
         Y2              =   3225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Budgeting && Purchasing "
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
         Index           =   6
         Left            =   240
         TabIndex        =   180
         Top             =   240
         Width           =   2100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   13
         X1              =   720
         X2              =   6120
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   12
         X1              =   720
         X2              =   6120
         Y1              =   375
         Y2              =   375
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "BuildPro Integration"
      Height          =   7185
      Index           =   14
      Left            =   7470
      TabIndex        =   328
      Top             =   1155
      Visible         =   0   'False
      Width           =   6795
      Begin VB.CheckBox chkTarion 
         Caption         =   "Integrate with Tarion Warranty Services"
         Height          =   285
         Left            =   2055
         TabIndex        =   48
         Top             =   2430
         Width           =   4095
      End
      Begin VB.TextBox txtBuildProWarrantyCoOwnerType 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   3585
         TabIndex        =   50
         Text            =   " "
         Top             =   3015
         Width           =   1755
      End
      Begin VB.CheckBox chkBuildProSendPhases 
         Caption         =   "Send phases to BuildPro"
         Height          =   285
         Left            =   2055
         TabIndex        =   47
         ToolTipText     =   $"FOptions.frx":5B7A
         Top             =   2175
         Width           =   4095
      End
      Begin VB.CheckBox chkMultiFamily 
         Caption         =   "Send shell and unit schedules for Multi family construction"
         Height          =   285
         Left            =   2055
         TabIndex        =   46
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txtBuildProWarrantyOwnerType 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   3585
         TabIndex        =   49
         Text            =   " "
         Top             =   2760
         Width           =   1755
      End
      Begin VB.CheckBox chkBuildProSendPOsImmediately 
         Caption         =   "Publish POs to BuildPro immediately upon generation"
         ForeColor       =   &H0000011D&
         Height          =   285
         Left            =   2055
         TabIndex        =   45
         Top             =   1665
         Width           =   4350
      End
      Begin VB.TextBox txtBuildProCompany 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2010
         TabIndex        =   41
         Top             =   495
         Width           =   1230
      End
      Begin zybCombo.zybCombobox cboBuildProEnvironment 
         Height          =   240
         Left            =   2010
         TabIndex        =   42
         Top             =   750
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   476
         Style           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ExtendedUI      =   0   'False
         DropDownWidth   =   0
      End
      Begin VB.TextBox txtBuildProPwd 
         BorderStyle     =   0  'None
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   2010
         PasswordChar    =   "*"
         TabIndex        =   44
         Top             =   1260
         Width           =   1755
      End
      Begin VB.TextBox txtBuildProUID 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2010
         TabIndex        =   43
         Text            =   " "
         Top             =   1005
         Width           =   1755
      End
      Begin VSFlex8Ctl.VSFlexGrid gEPOReasons 
         Height          =   3075
         Left            =   315
         TabIndex        =   51
         Top             =   3930
         Width           =   3315
         _cx             =   5847
         _cy             =   5424
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FOptions.frx":5C01
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
      Begin VB.TextBox txtScheduleTemplates 
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   3705
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Text            =   "FOptions.frx":5C6C
         Top             =   3915
         Width           =   2925
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "co buyer type"
         Height          =   195
         Index           =   115
         Left            =   1740
         TabIndex        =   355
         Top             =   3030
         Width           =   1770
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Warranty buyer type"
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   340
         Top             =   2775
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "BuildPro Settings "
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
         Index           =   28
         Left            =   270
         TabIndex        =   335
         Top             =   195
         Width           =   1530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   42
         X1              =   765
         X2              =   6165
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   43
         X1              =   765
         X2              =   6165
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Password"
         Height          =   195
         Index           =   99
         Left            =   765
         TabIndex        =   334
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Security Context"
         Height          =   195
         Index           =   100
         Left            =   765
         TabIndex        =   333
         Top             =   1005
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Environment"
         Height          =   195
         Index           =   102
         Left            =   765
         TabIndex        =   332
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Company Mnemonic"
         Height          =   195
         Index           =   103
         Left            =   405
         TabIndex        =   331
         Top             =   495
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "EPO Reasons"
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
         Index           =   30
         Left            =   255
         TabIndex        =   330
         Top             =   3630
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Schedule Templates "
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
         Index           =   29
         Left            =   3645
         TabIndex        =   329
         Top             =   3630
         Width           =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   45
         X1              =   780
         X2              =   6180
         Y1              =   3765
         Y2              =   3765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   44
         X1              =   780
         X2              =   6180
         Y1              =   3750
         Y2              =   3750
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Estimating Options"
      Height          =   6690
      Index           =   2
      Left            =   24060
      TabIndex        =   170
      Top             =   8985
      Visible         =   0   'False
      Width           =   13395
      Begin VB.Frame EstimatingFrame 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   2145
         Index           =   2
         Left            =   7425
         TabIndex        =   336
         Top             =   885
         Width           =   6090
         Begin VB.CheckBox chkSageSqlEstConsolidateItems 
            Caption         =   "Consolidate items to lump sum when importing estimates."
            Height          =   195
            Left            =   1095
            TabIndex        =   389
            Top             =   960
            Width           =   5535
         End
         Begin VB.TextBox txtSageSqlEstServer 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2040
            TabIndex        =   102
            Top             =   0
            Width           =   3435
         End
         Begin VB.TextBox txtSageSqlEstPswd 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   3510
            PasswordChar    =   "*"
            TabIndex        =   105
            Top             =   510
            Width           =   1464
         End
         Begin VB.TextBox txtSageSqlEstUser 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2025
            TabIndex        =   104
            Top             =   510
            Width           =   1464
         End
         Begin HFSystem.VBCombo cboSageSqlEstDatabase 
            Height          =   240
            Left            =   2025
            TabIndex        =   103
            Top             =   255
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   423
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Database"
            Height          =   195
            Index           =   107
            Left            =   1245
            TabIndex        =   339
            Top             =   255
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SQL Server"
            Height          =   195
            Index           =   106
            Left            =   1110
            TabIndex        =   338
            Top             =   15
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Login/Password"
            Height          =   195
            Index           =   105
            Left            =   795
            TabIndex        =   337
            Top             =   510
            Width           =   1155
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   14
            Left            =   150
            Picture         =   "FOptions.frx":5C72
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
      End
      Begin HFSystem.VBCombo cboTOSystem 
         Height          =   240
         Left            =   2265
         TabIndex        =   110
         Top             =   4830
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   423
         Style           =   2
         Text            =   "Combo1"
      End
      Begin VB.TextBox txtSelectAtTakeoffItem 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   " "
         Top             =   3930
         Width           =   3495
      End
      Begin VB.TextBox txtIncentiveItem 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   " "
         Top             =   3420
         Width           =   3495
      End
      Begin VB.TextBox txtLandItem 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   " "
         Top             =   3675
         Width           =   3495
      End
      Begin VB.TextBox txtIncentiveCostPercent 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2265
         TabIndex        =   106
         Text            =   "55"
         Top             =   3165
         Width           =   675
      End
      Begin HFSystem.VBCombo cboEstimatingSystem 
         Height          =   240
         Left            =   2040
         TabIndex        =   97
         Top             =   450
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.Frame frmTOSystem 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1965
         Index           =   1
         Left            =   336
         TabIndex        =   307
         Top             =   5070
         Width           =   6252
         Begin VB.TextBox txtOnScreenServer 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1875
            TabIndex        =   114
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox txtOnScreenPWD 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   1875
            PasswordChar    =   "*"
            TabIndex        =   117
            Top             =   1605
            Width           =   1755
         End
         Begin VB.TextBox txtOnScreenUID 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1875
            TabIndex        =   116
            Text            =   " "
            Top             =   1350
            Width           =   1755
         End
         Begin VB.TextBox txtOnScreenMDB 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1875
            TabIndex        =   112
            Top             =   315
            Width           =   3615
         End
         Begin VB.OptionButton optOnScreenType 
            Caption         =   "MS Access Database"
            Height          =   195
            Index           =   0
            Left            =   690
            TabIndex        =   111
            Top             =   60
            Value           =   -1  'True
            Width           =   2805
         End
         Begin VB.OptionButton optOnScreenType 
            Caption         =   "MS SQL Server Database"
            Height          =   195
            Index           =   1
            Left            =   690
            TabIndex        =   113
            Top             =   600
            Width           =   2805
         End
         Begin VB.TextBox txtOnScreenDatabase 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1875
            TabIndex        =   115
            Top             =   1095
            Width           =   3615
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   7
            Left            =   0
            Picture         =   "FOptions.frx":653C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Server"
            Height          =   195
            Index           =   56
            Left            =   1305
            TabIndex        =   312
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Index           =   57
            Left            =   1080
            TabIndex        =   311
            Top             =   1605
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Login Name"
            Height          =   195
            Index           =   59
            Left            =   915
            TabIndex        =   310
            Top             =   1365
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "File Name"
            Height          =   195
            Index           =   60
            Left            =   1065
            TabIndex        =   309
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Database"
            Height          =   195
            Index           =   61
            Left            =   1080
            TabIndex        =   308
            Top             =   1110
            Width           =   690
         End
         Begin VB.Image cmdChooseFile 
            Height          =   240
            Index           =   3
            Left            =   5550
            Picture         =   "FOptions.frx":6E06
            Top             =   330
            Width           =   240
         End
      End
      Begin VB.Frame frmTOSystem 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1965
         Index           =   2
         Left            =   336
         TabIndex        =   314
         Top             =   5070
         Width           =   6252
         Begin VB.Image Image1 
            Height          =   480
            Index           =   11
            Left            =   0
            Picture         =   "FOptions.frx":7390
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Frame EstimatingFrame 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   2145
         Index           =   1
         Left            =   0
         TabIndex        =   318
         Top             =   705
         Width           =   6600
         Begin VB.OptionButton optPipelineCommunityStyle 
            Caption         =   "Community specific Plans and Options"
            Height          =   195
            Index           =   1
            Left            =   2025
            TabIndex        =   387
            Top             =   1680
            Width           =   3720
         End
         Begin VB.OptionButton optPipelineCommunityStyle 
            Caption         =   "Global Plans and Options"
            Height          =   195
            Index           =   0
            Left            =   2025
            TabIndex        =   386
            Top             =   1440
            Width           =   2775
         End
         Begin VB.TextBox txtBIMBaseModelNumber 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2010
            TabIndex        =   101
            Text            =   " "
            Top             =   975
            Width           =   1755
         End
         Begin VB.TextBox txtCGVisionsRootURI 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2025
            TabIndex        =   98
            Top             =   0
            Width           =   3615
         End
         Begin VB.TextBox txtCGVisionsPWD 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   2025
            PasswordChar    =   "*"
            TabIndex        =   100
            Top             =   510
            Width           =   1755
         End
         Begin VB.TextBox txtCGVisionsUID 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2025
            TabIndex        =   99
            Text            =   " "
            Top             =   255
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Model Style"
            Height          =   195
            Index           =   84
            Left            =   1080
            TabIndex        =   388
            Top             =   1425
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Base Model Number"
            Height          =   195
            Index           =   104
            Left            =   480
            TabIndex        =   327
            Top             =   990
            Width           =   1440
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   13
            Left            =   150
            Picture         =   "FOptions.frx":7C5A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Root URI"
            Height          =   195
            Index           =   101
            Left            =   1260
            TabIndex        =   321
            Top             =   15
            Width           =   675
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Index           =   98
            Left            =   1245
            TabIndex        =   320
            Top             =   510
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Security Context"
            Height          =   195
            Index           =   97
            Left            =   780
            TabIndex        =   319
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Options "
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
         Index           =   20
         Left            =   240
         TabIndex        =   322
         Top             =   2910
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   35
         X1              =   900
         X2              =   6300
         Y1              =   3030
         Y2              =   3030
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   34
         X1              =   900
         X2              =   6300
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "System"
         Height          =   195
         Index           =   69
         Left            =   810
         TabIndex        =   317
         Top             =   495
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Takeoff System"
         Height          =   195
         Index           =   83
         Left            =   1050
         TabIndex        =   313
         Top             =   4830
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Takeoff System "
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
         Index           =   19
         Left            =   240
         TabIndex        =   248
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   31
         X1              =   900
         X2              =   6300
         Y1              =   4695
         Y2              =   4695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   30
         X1              =   900
         X2              =   6300
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label lblEditList 
         AutoSize        =   -1  'True
         Caption         =   "Edit WBS default descriptions..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   10
         Left            =   1500
         TabIndex        =   247
         Top             =   4260
         Width           =   2235
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estimating System "
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
         Index           =   0
         Left            =   240
         TabIndex        =   212
         Top             =   240
         Width           =   1605
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   720
         X2              =   6300
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   720
         X2              =   6300
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Select At Takeoff"
         Height          =   195
         Index           =   27
         Left            =   930
         TabIndex        =   187
         Top             =   3930
         Width           =   1245
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   2
         Left            =   2265
         Picture         =   "FOptions.frx":889C
         Top             =   3930
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sales Incentives"
         Height          =   195
         Index           =   16
         Left            =   1005
         TabIndex        =   185
         Top             =   3420
         Width           =   1170
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   0
         Left            =   2265
         Picture         =   "FOptions.frx":89E6
         Top             =   3420
         Width           =   240
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   1
         Left            =   2265
         Picture         =   "FOptions.frx":8B30
         Top             =   3675
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Land Costs"
         Height          =   195
         Index           =   15
         Left            =   1380
         TabIndex        =   184
         Top             =   3675
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Incentive Cost%"
         Height          =   195
         Index           =   21
         Left            =   1035
         TabIndex        =   183
         Top             =   3165
         Width           =   1140
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Accounting Import Options "
      Height          =   7185
      Index           =   5
      Left            =   270
      TabIndex        =   291
      Tag             =   "Security"
      Top             =   7350
      Visible         =   0   'False
      Width           =   6795
      Begin VB.TextBox txtExcludedVendorTypes 
         BorderStyle     =   0  'None
         Height          =   5385
         Left            =   330
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   360
         Text            =   "FOptions.frx":8C7A
         Top             =   1500
         Width           =   6270
      End
      Begin VB.OptionButton optManageActiveVendors 
         Caption         =   "I will manage which vendors are active in HomeFront."
         Height          =   225
         Index           =   1
         Left            =   1500
         TabIndex        =   294
         Top             =   780
         Value           =   -1  'True
         Width           =   4935
      End
      Begin VB.OptionButton optManageActiveVendors 
         Caption         =   "Let accounting manage which vendors are active."
         Height          =   225
         Index           =   0
         Left            =   1500
         TabIndex        =   293
         Top             =   540
         Width           =   4035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Excluded Vendor Types - don't import these ones."
         Height          =   195
         Index           =   79
         Left            =   315
         TabIndex        =   296
         Top             =   1260
         Width           =   3510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Active Vendors"
         Height          =   195
         Index           =   78
         Left            =   315
         TabIndex        =   295
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Import Options"
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
         Index           =   25
         Left            =   240
         TabIndex        =   292
         Top             =   240
         Width           =   1245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   37
         X1              =   720
         X2              =   6120
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   36
         X1              =   720
         X2              =   6120
         Y1              =   345
         Y2              =   345
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Document Management"
      Height          =   7185
      Index           =   0
      Left            =   22995
      TabIndex        =   324
      Tag             =   "Security"
      Top             =   1410
      Visible         =   0   'False
      Width           =   6795
      Begin VSFlex8Ctl.VSFlexGrid gDocumentClasses 
         Height          =   6375
         Left            =   300
         TabIndex        =   326
         Top             =   525
         Width           =   6180
         _cx             =   10901
         _cy             =   11245
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":8C80
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Document Classes "
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
         Index           =   27
         Left            =   240
         TabIndex        =   325
         Top             =   240
         Width           =   1635
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   41
         X1              =   720
         X2              =   6120
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   40
         X1              =   720
         X2              =   6120
         Y1              =   330
         Y2              =   330
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gTOC 
      Height          =   7215
      Left            =   0
      TabIndex        =   323
      Top             =   0
      Width           =   2595
      _cx             =   1981944289
      _cy             =   1981952438
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483629
      BackColorSel    =   -2147483643
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FOptions.frx":8D2C
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Image Image2 
         Height          =   240
         Left            =   2040
         Picture         =   "FOptions.frx":8D92
         Top             =   390
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Corporate Info "
      Height          =   7185
      Index           =   1
      Left            =   15930
      TabIndex        =   230
      Top             =   5940
      Visible         =   0   'False
      Width           =   6795
      Begin VB.CommandButton cmdUserSecurity 
         Caption         =   "Setup User Permissions"
         Height          =   375
         Left            =   1620
         TabIndex        =   305
         Top             =   6480
         Width           =   2055
      End
      Begin VB.CheckBox chkEnforcePasswordComplexity 
         Caption         =   "Enforce password complexity"
         Height          =   195
         Left            =   2280
         TabIndex        =   304
         ToolTipText     =   "Ensure that user passwords are complex."
         Top             =   3780
         Width           =   2625
      End
      Begin VB.TextBox txtCompanyEmail 
         BorderStyle     =   0  'None
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   150
         TabIndex        =   164
         Top             =   2910
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyFax 
         BorderStyle     =   0  'None
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   163
         Top             =   2655
         Width           =   1755
      End
      Begin VB.TextBox txtCompanyName 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   154
         Text            =   " "
         Top             =   510
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyAddr1 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   155
         Text            =   " "
         Top             =   810
         Width           =   3375
      End
      Begin VB.TextBox txtContactPerson 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   161
         Text            =   " "
         Top             =   2145
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyPostalCode 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   159
         Text            =   " "
         Top             =   1575
         Width           =   1755
      End
      Begin VB.TextBox txtCompanyPhone 
         BorderStyle     =   0  'None
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   162
         Top             =   2400
         Width           =   1755
      End
      Begin VB.TextBox txtCompanyAddr2 
         BorderStyle     =   0  'None
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   156
         Top             =   1065
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyCity 
         BorderStyle     =   0  'None
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   157
         Top             =   1320
         Width           =   2565
      End
      Begin VSFlex8Ctl.VSFlexGrid gDivisions 
         Height          =   1815
         Left            =   300
         TabIndex        =   165
         Top             =   4410
         Width           =   5805
         _cx             =   1981949951
         _cy             =   1981942913
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":931C
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
         Ellipsis        =   2
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
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
      Begin HFSystem.VBCombo cboCompanyCountry 
         Height          =   240
         Left            =   2280
         TabIndex        =   160
         Top             =   1830
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
      End
      Begin HFSystem.VBCombo cboCompanyProvince 
         Height          =   240
         Left            =   4860
         TabIndex        =   158
         Top             =   1320
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   423
      End
      Begin HFSystem.VBCombo cboBuilderType 
         Height          =   240
         Left            =   2280
         TabIndex        =   276
         Top             =   3360
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   423
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Builder Type"
         Height          =   195
         Index           =   64
         Left            =   1050
         TabIndex        =   277
         Top             =   3375
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Divisions"
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
         Index           =   18
         Left            =   270
         TabIndex        =   243
         Top             =   4140
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Country"
         Height          =   195
         Index           =   53
         Left            =   1050
         TabIndex        =   240
         Top             =   1845
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Email Address"
         Height          =   195
         Index           =   52
         Left            =   1050
         TabIndex        =   239
         Top             =   2910
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax"
         Height          =   195
         Index           =   45
         Left            =   1050
         TabIndex        =   238
         Top             =   2655
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Corporate Info "
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
         Index           =   17
         Left            =   270
         TabIndex        =   232
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact"
         Height          =   195
         Index           =   50
         Left            =   1050
         TabIndex        =   237
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   195
         Index           =   49
         Left            =   1050
         TabIndex        =   236
         Top             =   810
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Postal Code"
         Height          =   195
         Index           =   48
         Left            =   1050
         TabIndex        =   235
         Top             =   1590
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone"
         Height          =   195
         Index           =   47
         Left            =   1050
         TabIndex        =   234
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "City/Prov"
         Height          =   195
         Index           =   39
         Left            =   1050
         TabIndex        =   233
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   27
         X1              =   780
         X2              =   6180
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   26
         X1              =   780
         X2              =   6180
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Company Name"
         Height          =   195
         Index           =   51
         Left            =   1080
         TabIndex        =   231
         Top             =   510
         Width           =   1125
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Community Standards"
      Height          =   7185
      Index           =   7
      Left            =   16035
      TabIndex        =   174
      Top             =   6135
      Visible         =   0   'False
      Width           =   6795
      Begin VB.CheckBox chkUseCommunityStandards 
         Caption         =   "Enable community standards to automatically replace items during takeoff."
         Height          =   195
         Left            =   330
         TabIndex        =   201
         Top             =   540
         Width           =   5895
      End
      Begin VB.CheckBox chkCommunityStandardsArePhaseSpecific 
         Caption         =   "Community standards are also phase specific."
         Height          =   195
         Left            =   330
         TabIndex        =   200
         Top             =   780
         Width           =   5895
      End
      Begin VSFlex8Ctl.VSFlexGrid gStdItems 
         Height          =   4935
         Left            =   300
         TabIndex        =   198
         Top             =   1470
         Width           =   6045
         _cx             =   1981950375
         _cy             =   1981948417
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":9523
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Standard Items"
         Height          =   195
         Index           =   14
         Left            =   300
         TabIndex        =   202
         Top             =   1230
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Community Standards "
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
         Index           =   15
         Left            =   240
         TabIndex        =   199
         Top             =   240
         Width           =   1890
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   29
         X1              =   720
         X2              =   6120
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   28
         X1              =   720
         X2              =   6120
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   6780
      TabIndex        =   177
      Top             =   7335
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   8100
      TabIndex        =   178
      Top             =   7335
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Sales Pricing Settings"
      Height          =   7185
      Index           =   8
      Left            =   10110
      TabIndex        =   175
      Tag             =   "Security"
      Top             =   11145
      Visible         =   0   'False
      Width           =   6795
      Begin VB.CheckBox chkAdjustIncentive 
         Caption         =   "Calculate incentive retail when incentive cost changes."
         Height          =   195
         Left            =   480
         TabIndex        =   131
         Top             =   750
         Width           =   5535
      End
      Begin VB.CheckBox chkUseMaxVendorPricing 
         Caption         =   "Use Maximum price if no vendor price found."
         Height          =   195
         Left            =   480
         TabIndex        =   135
         ToolTipText     =   $"FOptions.frx":95A6
         Top             =   1980
         Width           =   5535
      End
      Begin VB.CheckBox chkSellingPriceIncludesTax 
         Caption         =   "Selling prices include GST && PST less rebates."
         Height          =   195
         Left            =   480
         TabIndex        =   179
         Top             =   1725
         Width           =   5535
      End
      Begin VB.CheckBox chkAdjustSelling 
         Caption         =   "Calculate selling prices when costs change."
         Height          =   195
         Left            =   480
         TabIndex        =   130
         Top             =   510
         Width           =   5535
      End
      Begin VB.CheckBox chkLockPostedSheets 
         Caption         =   "Lock worksheets when they are published."
         Height          =   195
         Left            =   480
         TabIndex        =   132
         Top             =   1005
         Width           =   5535
      End
      Begin VB.CheckBox chkEstimatingWorksheetsOnly 
         Caption         =   "Hide marketing worksheet interface."
         Height          =   195
         Left            =   480
         TabIndex        =   133
         Top             =   1245
         Width           =   5535
      End
      Begin VB.CheckBox chkMarketingWorkSheetAutoRefreshCosts 
         Caption         =   "Refresh costs when a marketing worksheet opens."
         Height          =   195
         Left            =   480
         TabIndex        =   134
         Top             =   1485
         Width           =   5535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sales Worksheets "
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
         Index           =   7
         Left            =   240
         TabIndex        =   186
         Top             =   240
         Width           =   1605
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   14
         X1              =   720
         X2              =   6120
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   15
         X1              =   720
         X2              =   6120
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Departmental Approvals "
      Height          =   7185
      Index           =   13
      Left            =   10770
      TabIndex        =   297
      Top             =   10860
      Visible         =   0   'False
      Width           =   6795
      Begin VSFlex8Ctl.VSFlexGrid gDepartments 
         Height          =   6015
         Left            =   390
         TabIndex        =   299
         Top             =   600
         Width           =   2025
         _cx             =   1981943284
         _cy             =   1981950322
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":9633
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
      Begin VSFlex8Ctl.VSFlexGrid gDepartmentApprovers 
         Height          =   2865
         Left            =   2640
         TabIndex        =   300
         Top             =   600
         Width           =   3825
         _cx             =   1981946459
         _cy             =   1981944766
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":9692
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
      Begin VSFlex8Ctl.VSFlexGrid gDepartmentGLs 
         Height          =   2895
         Left            =   2610
         TabIndex        =   301
         Top             =   3690
         Width           =   3825
         _cx             =   1981946459
         _cy             =   1981944818
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":9741
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Departmental Approvals "
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
         Index           =   26
         Left            =   270
         TabIndex        =   298
         Top             =   240
         Width           =   2100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   39
         X1              =   780
         X2              =   6180
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   38
         X1              =   780
         X2              =   6180
         Y1              =   375
         Y2              =   375
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Role Definitions "
      Height          =   7185
      Index           =   12
      Left            =   11325
      TabIndex        =   271
      Tag             =   "Security"
      Top             =   10545
      Visible         =   0   'False
      Width           =   6795
      Begin VB.TextBox txtJobRoles 
         BorderStyle     =   0  'None
         Height          =   3840
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   138
         Text            =   "FOptions.frx":97C8
         Top             =   690
         Width           =   2625
      End
      Begin VB.TextBox txtCustomerRoles 
         BorderStyle     =   0  'None
         Height          =   1680
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   137
         Text            =   "FOptions.frx":97CE
         Top             =   2790
         Width           =   2625
      End
      Begin VB.TextBox txtVendorRoles 
         BorderStyle     =   0  'None
         Height          =   1680
         Left            =   330
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   136
         Text            =   "FOptions.frx":97D4
         Top             =   690
         Width           =   2625
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Job Contact Template"
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
         Index           =   24
         Left            =   3240
         TabIndex        =   274
         Top             =   450
         Width           =   1875
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Customer Contact Roles"
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
         Index           =   23
         Left            =   240
         TabIndex        =   273
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Contact Roles"
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
         Index           =   22
         Left            =   210
         TabIndex        =   272
         Top             =   420
         Width           =   1875
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Sales System Integration "
      Height          =   7185
      Index           =   3
      Left            =   13725
      TabIndex        =   171
      Top             =   990
      Visible         =   0   'False
      Width           =   6795
      Begin VB.Frame SalesFrame 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4635
         Index           =   1
         Left            =   0
         TabIndex        =   213
         Top             =   720
         Width           =   6675
         Begin VB.TextBox txtCrmClientID 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2280
            TabIndex        =   73
            Text            =   " "
            Top             =   0
            Width           =   3045
         End
         Begin VB.TextBox txtCrmApiKey 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   2280
            PasswordChar    =   "*"
            TabIndex        =   74
            Top             =   255
            Width           =   3045
         End
         Begin zybCombo.zybCombobox cboCrmEnvironment 
            Height          =   240
            Left            =   2280
            TabIndex        =   75
            Top             =   510
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   476
            Style           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ExtendedUI      =   0   'False
            DropDownWidth   =   0
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Environment"
            Height          =   195
            Index           =   118
            Left            =   1035
            TabIndex        =   358
            Top             =   510
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Client ID"
            Height          =   195
            Index           =   117
            Left            =   1035
            TabIndex        =   357
            Top             =   0
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "API Key"
            Height          =   195
            Index           =   116
            Left            =   1230
            TabIndex        =   356
            Top             =   255
            Width           =   960
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   390
            Picture         =   "FOptions.frx":97DA
            Stretch         =   -1  'True
            Top             =   30
            Width           =   480
         End
      End
      Begin VB.Frame SalesFrame 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3465
         Index           =   0
         Left            =   0
         TabIndex        =   226
         Top             =   705
         Width           =   6675
         Begin VB.Label lblEditList 
            AutoSize        =   -1  'True
            Caption         =   "Edit model series list..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   9
            Left            =   1500
            TabIndex        =   246
            Top             =   870
            Width           =   1545
         End
         Begin VB.Label lblEditList 
            AutoSize        =   -1  'True
            Caption         =   "Edit option groups..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   8
            Left            =   1500
            TabIndex        =   242
            Top             =   1230
            Width           =   1410
         End
         Begin VB.Label lblEditList 
            AutoSize        =   -1  'True
            Caption         =   "Edit option categories..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   6
            Left            =   1500
            TabIndex        =   241
            Top             =   1440
            Width           =   1665
         End
         Begin VB.Label lblEditList 
            AutoSize        =   -1  'True
            Caption         =   "Edit community phases..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   5
            Left            =   1500
            TabIndex        =   228
            Top             =   510
            Width           =   1755
         End
         Begin VB.Label lblEditList 
            AutoSize        =   -1  'True
            Caption         =   "Edit communities..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   4
            Left            =   1500
            TabIndex        =   227
            Top             =   300
            Width           =   1320
         End
      End
      Begin HFSystem.VBCombo cboSalesSystem 
         Height          =   240
         Left            =   2280
         TabIndex        =   72
         Top             =   450
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   423
      End
      Begin VB.Frame SalesFrame 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4785
         Index           =   2
         Left            =   0
         TabIndex        =   253
         Top             =   705
         Width           =   6675
         Begin VB.OptionButton optWebUploadbyDivision 
            Caption         =   "Prompt for divisions when uploading"
            Height          =   240
            Left            =   1350
            TabIndex        =   264
            Top             =   3300
            Width           =   3195
         End
         Begin VB.OptionButton optWebUploadbyCommunity 
            Caption         =   "Prompt for communities when uploading"
            Height          =   240
            Left            =   1350
            TabIndex        =   263
            Top             =   3030
            Value           =   -1  'True
            Width           =   3195
         End
         Begin VB.CheckBox chkAppendUOMtoOptionDesc 
            Caption         =   "Append UOM to option descriptions."
            Height          =   240
            Left            =   1350
            TabIndex        =   153
            Top             =   2640
            Width           =   4005
         End
         Begin VB.TextBox txtWebAdminUser 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   149
            Text            =   " "
            Top             =   255
            Width           =   1755
         End
         Begin VB.TextBox txtBuilderCode 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   148
            Text            =   " "
            Top             =   0
            Width           =   1755
         End
         Begin VB.TextBox txtWebUploadUser 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   151
            Text            =   " "
            Top             =   1005
            Width           =   1755
         End
         Begin VB.TextBox txtWebUploadPswd 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   2280
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   152
            Top             =   1260
            Width           =   1755
         End
         Begin VB.TextBox txtWebUploadPath 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2280
            TabIndex        =   150
            Text            =   " "
            Top             =   750
            Width           =   4335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Options:"
            Height          =   195
            Index           =   32
            Left            =   930
            TabIndex        =   262
            Top             =   2370
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   390
            Picture         =   "FOptions.frx":A0A4
            Stretch         =   -1  'True
            Top             =   30
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Builder Code"
            Height          =   195
            Index           =   26
            Left            =   1050
            TabIndex        =   258
            Top             =   0
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Administrator"
            Height          =   195
            Index           =   38
            Left            =   1050
            TabIndex        =   257
            Top             =   255
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Security Context"
            Height          =   195
            Index           =   40
            Left            =   1050
            TabIndex        =   256
            Top             =   1005
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Password"
            Height          =   195
            Index           =   41
            Left            =   1050
            TabIndex        =   255
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ftp Upload Path"
            Height          =   195
            Index           =   42
            Left            =   1050
            TabIndex        =   254
            Top             =   750
            Width           =   1155
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "System"
         Height          =   195
         Index           =   36
         Left            =   1050
         TabIndex        =   229
         Top             =   450
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sales System "
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
         Index           =   13
         Left            =   270
         TabIndex        =   214
         Top             =   240
         Width           =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   780
         X2              =   6180
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   780
         X2              =   6180
         Y1              =   375
         Y2              =   375
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Custom Job Properties"
      Height          =   7185
      Index           =   10
      Left            =   15765
      TabIndex        =   176
      Top             =   5625
      Visible         =   0   'False
      Width           =   6795
      Begin VB.CommandButton cmdEditField 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5340
         TabIndex        =   316
         Top             =   5805
         Width           =   735
      End
      Begin VB.CommandButton cmdDeleteField 
         Caption         =   "Delete"
         Height          =   315
         Left            =   4560
         TabIndex        =   169
         Top             =   5820
         Width           =   735
      End
      Begin VB.CommandButton cmdAddField 
         Caption         =   "Add"
         Height          =   315
         Left            =   3780
         TabIndex        =   168
         Top             =   5820
         Width           =   735
      End
      Begin VSFlex8Ctl.VSFlexGrid gCustomFields 
         Height          =   5235
         Left            =   300
         TabIndex        =   167
         Top             =   465
         Width           =   6045
         _cx             =   10663
         _cy             =   9234
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FOptions.frx":A96E
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
      Begin VB.Label Label1 
         Caption         =   "Job Properties"
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
         Height          =   315
         Index           =   22
         Left            =   240
         TabIndex        =   181
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame TabFrame 
      Caption         =   "Supply Chain Integration"
      Height          =   7185
      Index           =   11
      Left            =   15000
      TabIndex        =   249
      Top             =   4815
      Visible         =   0   'False
      Width           =   6795
      Begin VB.TextBox txtWebAttachmentsFolder 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1860
         TabIndex        =   129
         Text            =   " "
         Top             =   3645
         Width           =   4005
      End
      Begin VB.CheckBox chkUseDetailedFieldPOView 
         Caption         =   "Use detailed edit view for all Field PO users"
         Height          =   195
         Left            =   1650
         TabIndex        =   128
         Top             =   2805
         Width           =   3645
      End
      Begin VB.CheckBox chkFieldPOField 
         Caption         =   "Tax Group"
         Height          =   195
         Index           =   7
         Left            =   2940
         TabIndex        =   127
         Top             =   2115
         Width           =   1275
      End
      Begin VB.CheckBox chkFieldPOField 
         Caption         =   "Category"
         Height          =   195
         Index           =   6
         Left            =   2940
         TabIndex        =   126
         Top             =   1905
         Width           =   1275
      End
      Begin VB.CheckBox chkFieldPOField 
         Caption         =   "Cost Code"
         Height          =   195
         Index           =   5
         Left            =   2940
         TabIndex        =   125
         Top             =   1695
         Width           =   1275
      End
      Begin VB.CheckBox chkFieldPOField 
         Caption         =   "Extra"
         Height          =   195
         Index           =   4
         Left            =   2940
         TabIndex        =   124
         Top             =   1485
         Width           =   1275
      End
      Begin VB.CheckBox chkFieldPOField 
         Caption         =   "Job"
         Height          =   195
         Index           =   3
         Left            =   1650
         TabIndex        =   123
         Top             =   2115
         Width           =   1275
      End
      Begin VB.CheckBox chkFieldPOField 
         Caption         =   "Rate"
         Height          =   195
         Index           =   2
         Left            =   1650
         TabIndex        =   122
         Top             =   1905
         Width           =   1275
      End
      Begin VB.CheckBox chkFieldPOField 
         Caption         =   "Quantity"
         Height          =   195
         Index           =   1
         Left            =   1650
         TabIndex        =   121
         Top             =   1695
         Width           =   1275
      End
      Begin VB.CheckBox chkFieldPOField 
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   1650
         TabIndex        =   120
         Top             =   1485
         Width           =   1275
      End
      Begin VB.TextBox txtFieldPOPrefix 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2475
         MaxLength       =   50
         TabIndex        =   119
         Text            =   " "
         Top             =   840
         Width           =   3375
      End
      Begin HFSystem.VBCombo cboPurchasingManager 
         Height          =   240
         Left            =   2490
         TabIndex        =   118
         Top             =   585
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
      End
      Begin VB.Image cmdChooseFolder 
         Height          =   240
         Index           =   2
         Left            =   5910
         Picture         =   "FOptions.frx":A9F3
         Top             =   3615
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vendor and Customer Portals "
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
         Index           =   2
         Left            =   240
         TabIndex        =   269
         Top             =   3345
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Attachments Folder"
         Height          =   195
         Index           =   62
         Left            =   390
         TabIndex        =   270
         Top             =   3645
         Width           =   1365
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   750
         X2              =   6150
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   750
         X2              =   6150
         Y1              =   3465
         Y2              =   3465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Options"
         Height          =   195
         Index           =   1
         Left            =   1380
         TabIndex        =   261
         Top             =   2535
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fields Required on a request:"
         Height          =   195
         Index           =   0
         Left            =   1380
         TabIndex        =   260
         Top             =   1245
         Width           =   2070
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Field PO Settings"
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
         Index           =   21
         Left            =   240
         TabIndex        =   250
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Default Purchasing Manager"
         Height          =   195
         Index           =   58
         Left            =   390
         TabIndex        =   252
         Top             =   585
         Width           =   2025
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   33
         X1              =   750
         X2              =   6150
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   32
         X1              =   750
         X2              =   6150
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Field PO Prefix"
         Height          =   195
         Index           =   46
         Left            =   1380
         TabIndex        =   251
         Top             =   840
         Width           =   1035
      End
   End
   Begin VB.Frame TabFrame 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vendor Communications"
      Height          =   7185
      Index           =   9
      Left            =   6120
      TabIndex        =   278
      Tag             =   "Security"
      Top             =   30
      Visible         =   0   'False
      Width           =   13080
      Begin VB.Frame frmSendFromAddr 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   840
         Left            =   690
         TabIndex        =   379
         Top             =   6180
         Width           =   4755
         Begin VB.TextBox txtEmailAddress 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   990
            TabIndex        =   382
            Text            =   "info@HomeFront-Software.com"
            Top             =   510
            Width           =   3675
         End
         Begin VB.OptionButton optSendFromCorporateEmail 
            Caption         =   "Send from the corporate address"
            Height          =   240
            Left            =   120
            TabIndex        =   381
            Top             =   255
            Width           =   2715
         End
         Begin VB.OptionButton optSendFromCurrentUsersEmail 
            Caption         =   "Send from the current users email address"
            Height          =   240
            Left            =   120
            TabIndex        =   380
            Top             =   0
            Value           =   -1  'True
            Width           =   3375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Index           =   31
            Left            =   345
            TabIndex        =   383
            Top             =   525
            Width           =   570
         End
      End
      Begin VB.OptionButton optEmailProtocol 
         Caption         =   "Send using an SMTP server"
         Height          =   240
         Index           =   1
         Left            =   450
         TabIndex        =   57
         Top             =   4155
         Width           =   2355
      End
      Begin VB.OptionButton optEmailProtocol 
         Caption         =   "Send using the current users Outlook account"
         Height          =   240
         Index           =   0
         Left            =   450
         TabIndex        =   56
         Top             =   3915
         Value           =   -1  'True
         Width           =   4605
      End
      Begin VB.OptionButton optEmailProtocol 
         Caption         =   "Send using SMTP via Office 365"
         Height          =   240
         Index           =   2
         Left            =   450
         TabIndex        =   58
         Top             =   4395
         Width           =   2685
      End
      Begin VB.Frame frmEmailConfig 
         Caption         =   "SMTP via Office 365"
         Height          =   2310
         Index           =   2
         Left            =   435
         TabIndex        =   372
         Top             =   4785
         Visible         =   0   'False
         Width           =   5340
         Begin VB.TextBox txt365User 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1335
            TabIndex        =   68
            Top             =   1005
            Width           =   1950
         End
         Begin VB.TextBox txt365Pswd 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   4215
            PasswordChar    =   "*"
            TabIndex        =   69
            Top             =   1005
            Width           =   795
         End
         Begin VB.TextBox txt365Secret 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1335
            TabIndex        =   67
            Top             =   750
            Width           =   3675
         End
         Begin VB.TextBox txt365TenantID 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1335
            TabIndex        =   66
            Top             =   495
            Width           =   3675
         End
         Begin VB.TextBox txt365AppID 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1335
            TabIndex        =   65
            Top             =   240
            Width           =   3675
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "User"
            Height          =   195
            Index           =   123
            Left            =   900
            TabIndex        =   377
            Top             =   1020
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Index           =   122
            Left            =   3255
            TabIndex        =   376
            Top             =   1035
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Secret Value"
            Height          =   195
            Index           =   30
            Left            =   315
            TabIndex        =   375
            Top             =   765
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tenant ID"
            Height          =   195
            Index           =   29
            Left            =   510
            TabIndex        =   374
            Top             =   510
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Application ID"
            Height          =   195
            Index           =   28
            Left            =   240
            TabIndex        =   373
            Top             =   255
            Width           =   990
         End
      End
      Begin VB.Frame frmNotices 
         BorderStyle     =   0  'None
         Caption         =   "frmNotices(0)"
         Height          =   2175
         Index           =   0
         Left            =   360
         TabIndex        =   281
         Top             =   840
         Width           =   5985
         Begin VB.TextBox txtNotification 
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   705
            MultiLine       =   -1  'True
            TabIndex        =   53
            Text            =   "FOptions.frx":AF7D
            Top             =   90
            Width           =   5205
         End
         Begin VB.TextBox txtNotification 
            BorderStyle     =   0  'None
            Height          =   1740
            Index           =   1
            Left            =   75
            MultiLine       =   -1  'True
            TabIndex        =   54
            Text            =   "FOptions.frx":AF83
            Top             =   390
            Width           =   5835
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Subject"
            Height          =   195
            Index           =   14
            Left            =   60
            TabIndex        =   282
            Top             =   105
            Width           =   540
         End
      End
      Begin VB.CheckBox chkCCPrjMgr 
         Caption         =   "Always CC: the project manager"
         Height          =   285
         Left            =   450
         TabIndex        =   55
         Top             =   3525
         Width           =   3525
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   255
         Left            =   4560
         TabIndex        =   279
         Top             =   540
         Width           =   2055
         Begin VB.Label lblReplacementParameters 
            AutoSize        =   -1  'True
            Caption         =   "Replacement Parameters"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   0
            TabIndex        =   280
            Top             =   0
            Width           =   1785
         End
      End
      Begin VB.Frame frmNotices 
         BorderStyle     =   0  'None
         Caption         =   "frmNotices(0)"
         Height          =   2175
         Index           =   1
         Left            =   360
         TabIndex        =   283
         Top             =   840
         Visible         =   0   'False
         Width           =   5985
         Begin VB.TextBox txtNotification 
            BorderStyle     =   0  'None
            Height          =   1740
            Index           =   3
            Left            =   75
            MultiLine       =   -1  'True
            TabIndex        =   71
            Text            =   "FOptions.frx":AF89
            Top             =   390
            Width           =   5835
         End
         Begin VB.TextBox txtNotification 
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   705
            MultiLine       =   -1  'True
            TabIndex        =   70
            Text            =   "FOptions.frx":AF8F
            Top             =   90
            Width           =   5205
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Subject"
            Height          =   195
            Index           =   23
            Left            =   60
            TabIndex        =   284
            Top             =   105
            Width           =   540
         End
      End
      Begin MSComctlLib.TabStrip tabNotices 
         Height          =   2535
         Left            =   330
         TabIndex        =   285
         Top             =   510
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   4471
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PO Issued"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PO Canceled"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame frmEmailConfig 
         Caption         =   "SMTP"
         Height          =   2265
         Index           =   1
         Left            =   435
         TabIndex        =   368
         Top             =   4785
         Visible         =   0   'False
         Width           =   5340
         Begin VB.CheckBox chkSTARTTLS 
            Caption         =   "STARTTLS"
            Height          =   195
            Left            =   555
            TabIndex        =   62
            Top             =   1020
            Width           =   1290
         End
         Begin VB.TextBox txtEmailUserID 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   3435
            TabIndex        =   63
            Top             =   645
            Width           =   1350
         End
         Begin VB.CheckBox chkSSL 
            Caption         =   "SSL"
            Height          =   195
            Left            =   555
            TabIndex        =   61
            Top             =   795
            Width           =   675
         End
         Begin VB.CheckBox chkEmailAuthentication 
            Caption         =   "Require Authentication?"
            Height          =   255
            Left            =   555
            TabIndex        =   60
            Top             =   495
            Width           =   2325
         End
         Begin VB.TextBox txtEmailPwd 
            BorderStyle     =   0  'None
            Height          =   240
            IMEMode         =   3  'DISABLE
            Left            =   3435
            PasswordChar    =   "*"
            TabIndex        =   64
            Top             =   900
            Width           =   1350
         End
         Begin VB.TextBox txtSMTPServer 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1110
            TabIndex        =   59
            Top             =   255
            Width           =   3675
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Server:Port"
            Height          =   195
            Index           =   43
            Left            =   225
            TabIndex        =   371
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "User"
            Height          =   195
            Index           =   66
            Left            =   2730
            TabIndex        =   370
            Top             =   645
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Index           =   67
            Left            =   2475
            TabIndex        =   369
            Top             =   930
            Width           =   885
         End
      End
      Begin VB.Label lblTestEmail 
         AutoSize        =   -1  'True
         Caption         =   "Send a test message..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3555
         TabIndex        =   378
         Top             =   4410
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Notices "
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
         Index           =   4
         Left            =   240
         TabIndex        =   287
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Email Configuration "
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
         Index           =   11
         Left            =   240
         TabIndex        =   286
         Top             =   3240
         Width           =   1710
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   8
         X1              =   720
         X2              =   6120
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   9
         X1              =   720
         X2              =   6120
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   22
         X1              =   720
         X2              =   6120
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   23
         X1              =   720
         X2              =   6120
         Y1              =   3375
         Y2              =   3375
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Invoice"
      Height          =   195
      Index           =   112
      Left            =   255
      TabIndex        =   352
      Top             =   780
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cost Code"
      Height          =   195
      Index           =   85
      Left            =   0
      TabIndex        =   303
      Top             =   0
      Width           =   1155
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   173
      X2              =   173
      Y1              =   482
      Y2              =   -500
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   -83
      X2              =   10016
      Y1              =   482
      Y2              =   482
   End
End
Attribute VB_Name = "FOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FOptions::"
Public Cancel As Boolean

Private mIntacctEntitiesLoaded  As Boolean
Private mIntacctTranTypeListsLoaded  As Boolean
Private mIntacctEstimateTypeListsLoaded  As Boolean
Private mIntacctItemsLoaded  As Boolean
Private mIntacctOneTimeItemUOM  As String
Private mIntacctGstItemUOM  As String

Private mDivCodeChangeWarned As Boolean


Private Function MaxPOLength() As Integer
    Select Case cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex)
        
        Case asTimberline, asMasterBuilder
            MaxPOLength = 12
        
        Case Else
            MaxPOLength = 20
        
    End Select

End Function

Private Sub cboCompanyCountry_Click()
    cboCompanyProvince.Clear
    cboCompanyProvince.ListIndex = -1
    Call LoadComboBox(cboCompanyProvince, HFApp.Databases(dbHomefront), "select distinct state,'',0 from countrycodes where country=" & DbQuote(Str, cboCompanyCountry.Text) & "order by 1")
End Sub

Private Sub cboEstimatingSystem_Click()
    Dim i As Long
    For i = 1 To EstimatingFrame.UBound
        EstimatingFrame(i).Visible = i = cboEstimatingSystem.ListIndex
    Next
End Sub


Private Sub cboIntacctEntity_DropDown()
    Call LoadIntacctEntities
End Sub

Private Sub cboIntacctEntity_KeyDown(KeyCode As Integer, Shift As Integer)
    cboIntacctEntity.ListIndex = 0
End Sub

Private Sub cboIntacctEstimateType_DropDown()
    Call LoadIntacctEstimateTypes
End Sub






 

Private Sub cboIntacctItem_Click(Index As Integer)
    If Index = 0 Then
        mIntacctOneTimeItemUOM = GetComboBoxListKey(cboIntacctItem(Index))
    Else
        mIntacctGstItemUOM = GetComboBoxListKey(cboIntacctItem(Index))
    End If
End Sub

Private Sub cboIntacctItem_DropDown(Index As Integer)
    Call LoadIntacctItems
End Sub

Private Sub cboIntacctPOCOType_DropDown()
    Call LoadIntacctTranTypes
End Sub

Private Sub cboIntacctPOType_DropDown()
    Call LoadIntacctTranTypes
End Sub
Private Sub cboIntacctInvoiceType_DropDown()
    Call LoadIntacctTranTypes
End Sub

Private Sub LoadIntacctEstimateTypes()
On Error GoTo eh
    Dim rs As Recordset
    
    Dim a As String
    
    If mIntacctEstimateTypeListsLoaded Then Exit Sub
    
    'remember selection
    a = cboIntacctEstimateType.Text
    
    cboIntacctEstimateType.ClearList
    
    Set rs = Intacct_EstimateTypes(txtIntacctCompanyID.Text, txtIntacctUID.Text, txtIntacctPWD.Text, cboIntacctEntity.Text)
    While Not rs.EOF
        Call cboIntacctEstimateType.AddItem("" & rs("name"))
        rs.MoveNext
    Wend
    
    Call SetComboBoxListIndex(cboIntacctEstimateType, a)
    
    mIntacctEstimateTypeListsLoaded = True
eh: Exit Sub
End Sub

Private Sub LoadIntacctTranTypes()
On Error GoTo eh
    Dim rs As Recordset
    
    Dim a As String
    Dim b As String
    Dim c As String
    Dim d As String
    Dim E As String
    
    If mIntacctTranTypeListsLoaded Then Exit Sub
    
    'remember selection
    a = cboIntacctPOType.Text
    b = cboIntacctSubContractType.Text
    c = cboIntacctInvoiceType.Text
    d = cboIntacctPOCOType.Text
    E = cboIntacctSubContractCOType.Text
    
    
    cboIntacctPOType.ClearList
    cboIntacctSubContractType.ClearList
    cboIntacctInvoiceType.ClearList
    cboIntacctPOCOType.ClearList
    cboIntacctSubContractCOType.ClearList
    
    Set rs = Intacct_TransactionTypes(txtIntacctCompanyID.Text, txtIntacctUID.Text, txtIntacctPWD.Text, cboIntacctEntity.Text)
    While Not rs.EOF
        Call cboIntacctPOType.AddItem("" & rs("docid"))
        Call cboIntacctSubContractType.AddItem("" & rs("docid"))
        Call cboIntacctSubContractCOType.AddItem("" & rs("docid"))
        Call cboIntacctInvoiceType.AddItem("" & rs("docid"))
        Call cboIntacctPOCOType.AddItem("" & rs("docid"))
        rs.MoveNext
    Wend
    
    Call SetComboBoxListIndex(cboIntacctPOType, a)
    Call SetComboBoxListIndex(cboIntacctSubContractType, b)
    Call SetComboBoxListIndex(cboIntacctInvoiceType, c)
    Call SetComboBoxListIndex(cboIntacctPOCOType, d)
    Call SetComboBoxListIndex(cboIntacctSubContractCOType, E)
    
    mIntacctTranTypeListsLoaded = True
    
eh: Exit Sub
End Sub
Private Sub LoadIntacctEntities()
On Error GoTo eh
    Dim rs As Recordset
    
    Dim a As String
    
    If mIntacctEntitiesLoaded Then Exit Sub
    
    'remember selection
    a = cboIntacctEntity.Text
    
    cboIntacctEntity.ClearList
    Call cboIntacctEntity.AddItem("")
    
    Set rs = Intacct_Entities(txtIntacctCompanyID.Text, txtIntacctUID.Text, txtIntacctPWD.Text, "")
    While Not rs.EOF
        Call cboIntacctEntity.AddItem("" & rs("entityid"))
        rs.MoveNext
    Wend
    
    Call SetComboBoxListIndex(cboIntacctEntity, a)
    
    mIntacctEntitiesLoaded = True
eh: Exit Sub
End Sub
Private Sub LoadIntacctItems()
On Error GoTo eh
    Dim rs As Recordset
    Dim tag As String
    Dim a As String
    Dim b As String
    
    If mIntacctItemsLoaded Then Exit Sub
    
    a = cboIntacctItem(0).Text
    cboIntacctItem(0).ClearList
    
    b = cboIntacctItem(1).Text
    cboIntacctItem(1).ClearList
    
    Set rs = Intacct_Items(txtIntacctCompanyID.Text, txtIntacctUID.Text, txtIntacctPWD.Text, cboIntacctEntity.Text)
    tag = ""
    While Not rs.EOF
        cboIntacctItem(0).AddItem Trim("" & rs("ItemID"))
        cboIntacctItem(1).AddItem Trim("" & rs("ItemID"))
        tag = tag & Chr(1) & "" & rs("UOM")
        rs.MoveNext
    Wend
    cboIntacctItem(0).tag = Mid(tag, 2)
    cboIntacctItem(1).tag = Mid(tag, 2)
    
    Call SetComboBoxListIndex(cboIntacctItem(0), a)
    Call SetComboBoxListIndex(cboIntacctItem(1), b)
    
    mIntacctItemsLoaded = True
eh: Exit Sub
End Sub



Private Sub cboIntacctSubContractCOType_DropDown()
    Call LoadIntacctTranTypes
End Sub


Private Sub cboIntacctSubContractType_DropDown()
    Call LoadIntacctTranTypes
End Sub

Private Sub cboMBAPILevel_Click()
    cboMBInstance.Visible = cboMBAPILevel.Text = "v20"
    cboMBDrive.Visible = cboMBAPILevel.Text <> "v20"
End Sub




Private Sub cboMBCompany_DropDown()
On Error GoTo eh
    
    Dim s As String
    Dim i As Long
    Dim Folders As String
    Dim Companies As String
    Dim c As New Connection
    Dim rs As Recordset
    
    With cboMBCompany
        
        'clearing removes the current value
        s = .Text
        .Clear
        .Text = s
        .SelStart = 0
        .SelLength = Len(s)
        
        Screen.MousePointer = vbHourglass
        If cboMBAPILevel.Text = "v20" Then
            'connect to instance
            s = "Driver={SQL Server};Server=" & cboMBInstance.Text
            c.ConnectionTimeout = 1
            c.CommandTimeout = 1
            c.Open s
            'add db names to list
            s = "select name from sys.databases where database_id>4 order by name"
            Set rs = c.Execute(s)
            While Not rs.EOF
                .AddItem "" & rs(0)
                rs.MoveNext
            Wend
        Else
            'get folders
            s = VBA.Dir(cboMBDrive.Text & "\MB7\", vbDirectory)
            While s <> ""
                If s <> "." And s <> ".." And LCase(s) <> "database" Then Folders = Folders & Chr(1) & s
                s = VBA.Dir()
            Wend
            Folders = Mid(Folders, 2)
            'add company folders to list
            For i = 1 To Parse(Folders, , Chr(1))
                s = Parse(Folders, i, Chr(1))
                If FileExists(cboMBDrive.Text & "\MB7\" & s & "\Cmpany.dbf") Then
                    .AddItem s
                End If
            Next
        End If
        Screen.MousePointer = vbDefault
    End With
    
    
Exit Sub
eh:  'Screen.MousePointer = vbDefault
End Sub




Private Sub cboSageSqlEstDatabase_DropDown()
On Error Resume Next
    Dim c As New Connection
    Screen.MousePointer = vbHourglass
    c.CommandTimeout = 5
    c.ConnectionTimeout = 5
    c.Open "Driver={SQL Server};Server=" & Me.txtSageSqlEstServer.Text & ";Trusted_Connection=yes;"
    Call LoadComboBox(cboSageSqlEstDatabase, c, "select name,'',0 from sys.databases where name not in('master','tempdb','model','msdb')")
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboTOSystem_Click()
On Error Resume Next
    Dim i As Long
    For i = frmTOSystem.LBound To frmTOSystem.UBound
        frmTOSystem(i).Visible = cboTOSystem.ListIndex = i
    Next
End Sub

Private Sub chkAllowForecast_Click()
    chkWarnForecast.Enabled = chkAllowForecast.Value = vbChecked
    If Not chkWarnForecast.Enabled Then chkWarnForecast.Value = vbUnchecked
End Sub





Private Sub chkSSL_Click()
    If chkSSL.Value = vbChecked Then chkSTARTTLS.Value = vbUnchecked
End Sub

Private Sub chkSTARTTLS_Click()
    If chkSTARTTLS.Value = vbChecked Then chkSSL.Value = vbUnchecked
End Sub

Private Sub cmdChooseFile_Click(Index As Integer)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Select Case Index
        Case 0
            txtQuickBooksFile.SetFocus
            s = txtQuickBooksFile.Text
            If VBGetOpenFileName(s, , , , , , "Database File (*.QBW)|*.QBW", , , "Open Database", , Me.hwnd) Then
                txtQuickBooksFile.Text = s
            End If
            
            
        Case 2
            txtSimplyFile.SetFocus
            If VBGetOpenFileName(s, , , , , , "Database File (*.SAI)|*.SAI", , , "Open Database", , Me.hwnd) Then
                txtSimplyFile.Text = s
            End If
            
        Case 3
            txtOnScreenMDB.SetFocus
            If VBGetOpenFileName(s, , , , , , "MS Access files (*.mdb)|*.mdb", , , "Open Database", , Me.hwnd) Then
                optOnScreenType(0).Value = True
                txtOnScreenMDB.Text = s
            End If
            
    End Select
        
Exit Sub
eh: Call errHandler(SRCFILE & "cmdChooseFile_Click")
End Sub





Private Sub cmdEditField_Click()
    Call FAddProperty.EditField(gCustomFields.TextMatrix(gCustomFields.Row, gCustomFields.ColIndex("FieldName")))
    Call ReadCustomProperties
End Sub

Private Sub cmdUserSecurity_Click()
    Call HFApp.RunTask("EditSecurity")
End Sub


Private Sub gDepartmentApprovers_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gDepartmentApprovers
        Select Case .ColKey(Col)
            Case "invoicelimit":      .EditText = Val(.EditText)
        End Select
    End With
End Sub

Private Sub gDivisions_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gDivisions
        .ComboList = ""
        .EditMaxLength = 0
        Select Case .ColKey(Col)
            Case "ID":      Cancel = True
            Case "Code":    .EditMaxLength = 15
            Case "Company": .ComboList = "|..."
        End Select
    End With
End Sub

Private Sub gDivisions_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    
    With gDivisions
        Select Case .ColKey(Col)
            Case "Company":
                
                'system_setup and divisions both have address fields for some dumb reason so do this bs to keep them the same
                'copy this divisions address from the screen to the grid
                If .TextMatrix(Row, .ColIndex("id")) = HFApp.DivisionID Then
                    .TextMatrix(Row, .ColIndex("address1")) = txtCompanyAddr1.Text
                    .TextMatrix(Row, .ColIndex("address2")) = txtCompanyAddr2.Text
                    .TextMatrix(Row, .ColIndex("city")) = txtCompanyCity.Text
                    .TextMatrix(Row, .ColIndex("province")) = cboCompanyProvince.Text
                    .TextMatrix(Row, .ColIndex("country")) = cboCompanyCountry.Text
                    .TextMatrix(Row, .ColIndex("postal")) = txtCompanyPostalCode.Text
                    .TextMatrix(Row, .ColIndex("phone")) = txtCompanyPhone.Text
                    .TextMatrix(Row, .ColIndex("fax")) = txtCompanyFax.Text
                End If
                
                'edit values
                Call FCompany.ShowForm(gDivisions)
                
                'copy this divisions address from the grid to the screen
                If .TextMatrix(Row, .ColIndex("id")) = HFApp.DivisionID Then
                    txtCompanyAddr1.Text = .TextMatrix(Row, .ColIndex("address1"))
                    txtCompanyAddr2.Text = .TextMatrix(Row, .ColIndex("address2"))
                    txtCompanyCity.Text = .TextMatrix(Row, .ColIndex("city"))
                    Call SetComboBoxListIndex(cboCompanyCountry, .TextMatrix(Row, .ColIndex("country")))
                    Call SetComboBoxListIndex(cboCompanyProvince, .TextMatrix(Row, .ColIndex("province")))
                    txtCompanyPostalCode.Text = .TextMatrix(Row, .ColIndex("postal"))
                    txtCompanyPhone.Text = .TextMatrix(Row, .ColIndex("phone"))
                    txtCompanyFax.Text = .TextMatrix(Row, .ColIndex("fax"))
                End If
                
        End Select
    End With
End Sub

Private Sub gDivisions_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case True
'        Case KeyCode = vbKeyDelete
'            If gDivisions.Row > 0 And gDivisions.Row < gDivisions.Rows - 1 Then gDivisions.RowHidden(gDivisions.Row) = True
'    End Select
End Sub

Private Sub gDivisions_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With gDivisions
        
        If Col = .ColIndex("Code") And txtBuildProCompany.Text <> "" Then
            If Not mDivCodeChangeWarned Then
                Cancel = vbCancel = MsgBox("If you change the division code here you must also update the division in BuildPro. You must also notify support@ihyphen.com, requesting them to update your integration configuration.", vbInformation + vbOKCancel, "HomeFront")
                mDivCodeChangeWarned = Not Cancel
            End If
        End If
    
        If Row = .Rows - 1 Then .AddItem ""
        Call .AutoSize(0, 1, 2)
    End With
    
End Sub

Private Sub gDocumentClasses_AfterSort(ByVal Col As Long, Order As Integer)
    gDocumentClasses.AddItem ""
End Sub

Private Sub gDocumentClasses_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gDocumentClasses
        Cancel = .TextMatrix(Row, .ColIndex("DocumentClass")) = "" And Col <> .ColIndex("DocumentClass")
    End With
End Sub

Private Sub gDocumentClasses_BeforeSort(ByVal Col As Long, Order As Integer)
    gDocumentClasses.Rows = gDocumentClasses.Rows - 1
End Sub


Private Sub gEPOReasons_KeyDown(KeyCode As Integer, Shift As Integer)
    With gEPOReasons
    If KeyCode = vbKeyDelete And Shift <> 0 Then
    If .Row <> .Rows - 1 Then
        .RemoveItem .Row
    End If
    End If
    End With
End Sub

Private Sub gEPOReasons_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = gEPOReasons.Rows - 1 Then gEPOReasons.AddItem ""
End Sub


Private Sub gStdItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    Dim s As String
    With gStdItems
        If Row = .Rows - 1 Then .AddItem ""
        s = HFApp.SqlExec("select description from tblphaseitem where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, .TextMatrix(Row, 0)) & " and item=" & DbQuote(Str, .TextMatrix(Row, 1)), dbHomefront)(0)
        .TextMatrix(Row, 2) = s
        
    End With
End Sub

Private Sub gStdItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = 2
End Sub

Private Sub gStdItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    
    s = "select phase,item,description from tblphaseitem where DivisionID = " & HFApp.DivisionID
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s) Then
        With gStdItems
            .TextMatrix(.Row, 0) = FPickList.SelectedItem(1)
            .TextMatrix(.Row, 1) = FPickList.SelectedItem(2)
            .TextMatrix(.Row, 2) = FPickList.SelectedItem(3)
        End With
        Call gStdItems_AfterEdit(Row, Col)
    End If
    
    
End Sub

Private Sub gStdItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And gStdItems.Row > 0 And gStdItems.Row < gStdItems.Rows - 1 Then
        gStdItems.RowHidden(gStdItems.Row) = True
        gStdItems.Row = GridNextVisibleRow(gStdItems, gStdItems.Row)
    End If
End Sub



Private Sub lblEditList_Click(Index As Integer)
    Dim s As String
    
    Select Case Index
        
        Case 11 'book of account
            If HFApp.Options(AccountingSystem) = asTimberline Then
                s = "select BookOfAccount,WalletBankAccount,DataSourceID WalletPartyID,Sage300GLPrefixLength,'' Sage300GLPrefixes from DataSources"
            Else
                s = "select BookOfAccount,WalletBankAccount,DataSourceID WalletPartyID from DataSources"
            End If
            Call FDBGrid.ShowForm("Book of Accounts", s, "DataSources", "BookOfAccount", False, , "DataSourceID")
            Call HFApp.SqlExec("update DataSources set Datasourceid=newid() where datasourceid is null")
            
            s = cboBookOfAccount.Text
            Call LoadComboBox(cboBookOfAccount, HFApp.Databases(dbHomefront), "select distinct BookOfAccount,'',0 from DataSources order by 1")
            Call SetComboBoxListIndex(cboBookOfAccount, s)
        
        
        
        Case 10
            s = "select Item,Custom_Description from customDescriptions where item like 'WBS__' order by item"
            Call FDBGrid.ShowForm("WBS Codes", s, "customDescriptions", "Item", False)
        
        
        
        Case 2:
            s = "select divisionid,TaxGroup,Description, JCRate,NJCRate from taxgroups where DivisionID = " & HFApp.DivisionID
            Select Case cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex)
                Case asTimberline
                    Call FDBGrid.ShowForm("Tax Groups", s, "taxgroups", "divisionid,taxgroup", True, "divisionid")
                Case Else
                    Call FDBGrid.ShowForm("Tax Groups", s, "taxgroups", "divisionid,taxgroup", False, "divisionid")
            End Select
            Call HFApp.SqlExec("update taxgroups set grouprate=njcrate+jcrate")
                 
                 
        Case 0:  Call HFApp.RunTask("EditJCCostCodes")
        Case 1:  Call HFApp.RunTask("EditJCCategories")
        Case 3:  Call FChangeTaxes.Show(vbModal)
        Case 4:  Call FDBGrid.ShowForm("Communities", "select Area,Description,CompanyName,GSTNumber TaxNumber,Address1,Address2,City,Province,Zip PostalCode,County,Country,Phone,Fax,Inactive from tbllocality order by 1", "tbllocality", "area", cboSalesSystem.ListIndex <> 0)
        Case 5:  Call FDBGrid.ShowForm("Community Phases", "select Community,CommunityPhase,Description from CommunityPhase order by 1,2", "CommunityPhase", "Community,CommunityPhase", cboSalesSystem.ListIndex <> 0)
        Case 6:  Call FDBGrid.ShowForm("Option Categories", "select Group_Code,Category,Description,COMarkup from tblcategories", "tblCategories", "category", cboSalesSystem.ListIndex <> 0)
        
        Case 8:  Call FDBGrid.ShowForm("Option Groups", "select Major_Group,Description from tblmajorgroups", "tblMajorGroups", "Major_Group", cboSalesSystem.ListIndex <> 0)
    
        Case 9:  Call FDBGrid.ShowForm("Series List", "select Series,Description from tblSeries", "tblSeries", "series", cboSalesSystem.ListIndex <> asHomeFront)
    End Select
End Sub



Private Sub lblTestEmail_Click()
On Error Resume Next
    Call SaveEmailConfig
    HFApp.SendMail True, UserEmail, "", "test email", "testing 1 2 3", "", ""
    Screen.MousePointer = vbNormal
End Sub


Private Function UserEmail() As String
On Error Resume Next
    UserEmail = "" & HFApp.SqlExec("SELECT email FROM User_Manager WHERE User_ID=" & DbQuote(Str, HFApp.LoginID), dbHomefront)(0)
End Function

Private Sub optEmailProtocol_Click(Index As Integer)
    frmSendFromAddr.Visible = Not optEmailProtocol(0).Value
    frmEmailConfig(1).Visible = optEmailProtocol(1).Value
    frmEmailConfig(2).Visible = optEmailProtocol(2).Value
End Sub

Private Sub optPOByJobPOIndex_Click()
    Call optPOByJob_Click
End Sub
Private Sub optPOSequential_Click()
    Call optPOByJob_Click
End Sub
Private Sub optPOByJob_Click()

    Label1(17).Enabled = False
    txtPOSegment(1).Enabled = False
    txtPOSegment(2).Enabled = False
    txtPOSegment(3).Enabled = False
    Label1(73).Enabled = False
    txtPOSegment(4).Enabled = False
    txtPOSegment(5).Enabled = False
    txtPOSegment(6).Enabled = False
    txtPOSegment(7).Enabled = False
    txtPOSegment(8).Enabled = False
    lblNextPO.Enabled = False
    lblPOPrefix.Enabled = False
    txtNextPO.Enabled = False
    txtPOPrefix.Enabled = False
    cmdNextPO.Enabled = False
    
    Select Case True
    Case optPOByJob.Value
        Label1(17).Enabled = True
        txtPOSegment(1).Enabled = False
        txtPOSegment(2).Enabled = True
        txtPOSegment(3).Enabled = True
    Case optPOByJobPOIndex.Value
        Label1(73).Enabled = True
        txtPOSegment(4).Enabled = False
        txtPOSegment(5).Enabled = True
        txtPOSegment(6).Enabled = True
        txtPOSegment(7).Enabled = True
        txtPOSegment(8).Enabled = True
    Case Else
        lblNextPO.Enabled = True
        lblPOPrefix.Enabled = True
        txtNextPO.Enabled = True
        txtPOPrefix.Enabled = True
        cmdNextPO.Enabled = True
    End Select
    
    lblPONumberSample.Caption = SamplePONumber
    
End Sub

Private Sub cmdAddField_Click()
    Call FAddProperty.AddField
    Call ReadCustomProperties
End Sub


Private Sub cmdChooseFolder_Click(Index As Integer)

    Dim s As String
    Dim t As Textbox
    
    Select Case Index
        Case 0: Set t = txtTLFolder
        Case 1: Set t = txtTLARFolder
        Case 2: Set t = txtWebAttachmentsFolder
    End Select
        
    t.SetFocus
    s = t.Text
    If VBChooseFolder(0, BIF_RETURNONLYFSDIRS, s, , Me.hwnd, s) Then
        t.Text = s
    End If

End Sub


Private Sub cmdChooseItem_Click(Index As Integer)
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", "SELECT Phase,Item,Description FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID) Then
        Select Case Index
            Case 0
                txtIncentiveItem.tag = FPickList.SelectedItem("Phase") & Chr(1) & FPickList.SelectedItem("Item") & Chr(1) & FPickList.SelectedItem("Description")
                txtIncentiveItem.Text = FPickList.SelectedItem("Phase") & "/" & FPickList.SelectedItem("Item") & " -- " & FPickList.SelectedItem("Description")
            Case 1
                txtLandItem.tag = FPickList.SelectedItem("Phase") & Chr(1) & FPickList.SelectedItem("Item") & Chr(1) & FPickList.SelectedItem("Description")
                txtLandItem.Text = FPickList.SelectedItem("Phase") & "/" & FPickList.SelectedItem("Item") & " -- " & FPickList.SelectedItem("Description")
            Case 2
                txtSelectAtTakeoffItem.tag = FPickList.SelectedItem("Phase") & Chr(1) & FPickList.SelectedItem("Item") & Chr(1) & FPickList.SelectedItem("Description")
                txtSelectAtTakeoffItem.Text = FPickList.SelectedItem("Phase") & "/" & FPickList.SelectedItem("Item") & " -- " & FPickList.SelectedItem("Description")
        End Select
    End If
End Sub
Private Sub cmdDeleteField_Click()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    If vbNo = MsgBox("STOP!" & vbCrLf & vbCrLf & "If you delete this property it will be removed immediately." & vbCrLf & _
                     "The data it contains will be irretrievably lost. There is no" & vbCrLf & _
                     "cancel. There is no undo. You could recreate the property but" & vbCrLf & _
                     "the data will be gone FOREVER!" & vbCrLf & vbCrLf & _
                     "Are you sure this is what you want to do?" & vbCrLf _
                    , vbYesNo + vbCritical, "Confirm Delete") Then Exit Sub
    
    With gCustomFields
        On Error Resume Next
        s = "EXEC sp_unbindefault " & DbQuote(Str, "dbo.JobCustomFields." & vbQuote & .TextMatrix(.Row, .ColIndex("FieldName")) & vbQuote)
        Call HFApp.SqlExec(s, dbHomefront)
        
        s = "EXEC sp_unbindefault " & DbQuote(Str, "dbo.WorkticketCustomFlds." & vbQuote & .TextMatrix(.Row, .ColIndex("FieldName")) & vbQuote)
        Call HFApp.SqlExec(s, dbHomefront)
        
        Call HFApp.SqlExec("ALTER TABLE dbo.WorkticketCustomFlds DROP COLUMN " & vbQuote & .TextMatrix(.Row, .ColIndex("FieldName")) & vbQuote)
        On Error GoTo eh
            
        Call HFApp.SqlExec("ALTER TABLE dbo.JobCustomFields DROP COLUMN " & vbQuote & .TextMatrix(.Row, .ColIndex("FieldName")) & vbQuote)
        Call HFApp.SqlExec("DELETE FROM JobCustomFieldDefs WHERE Name=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FieldName"))))
        
    End With
    
    Call ReadCustomProperties
    
Exit Sub
eh: Call errHandler(SRCFILE & "Form_KeyDown")
    Call ReadCustomProperties
End Sub





Private Sub Form_Load()
On Error GoTo eh

    Dim i As Long
    Dim s As String
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim k() As String


    Dim c As Control
    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            c.BackColor = vbButtonFace
        End If
    Next
    
    With gTOC
        .Rows = 1
        .Cell(flexcpBackColor, 0, 0, 0, 1) = -2147483645
        For i = 0 To TabFrame.Count - 1
        
            TabFrame(i).BorderStyle = 0
            TabFrame(i).Move gTOC.Width + 1, 0, TabFrame(0).Width, TabFrame(0).Height
            .AddItem vbTab & TabFrame(i).Caption
            
        Next
        .Cell(flexcpPicture, 1, 1, .Rows - 1, 1) = Image2.Picture
        .Sort = flexSortStringAscending
    End With


    Call LoadComboBox(cboCompanyCountry, HFApp.Databases(dbHomefront), "select distinct country,'',0 from countrycodes order by 1")

    Call LoadComboBox(cboBookOfAccount, HFApp.Databases(dbHomefront), "select distinct BookOfAccount,'',0 from DataSources order by 1")


    
    cboMBAPILevel.AddItem "v18"
    cboMBAPILevel.AddItem "v19"
    cboMBAPILevel.AddItem "v20"
    
    cboTOSystem.AddItem "none"
    cboTOSystem.AddItem "On-Screen Takeoff"
    cboTOSystem.AddItem "PlanSwift"
    
    For Each drv In fso.Drives
    If drv.IsReady Then
        cboMBDrive.AddItem drv.DriveLetter & ":"
    End If
    Next
        
    With cboMBInstance
        Call RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", k)
        For i = 0 To UBound(k)
            If k(i) <> "" Then
                .AddItem MachineName() & "\" & k(i)
            End If
        Next
        Call RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", k, True)
        For i = 0 To UBound(k)
            If k(i) <> "" Then
                .AddItem MachineName() & "\" & k(i)
            End If
        Next
    End With
        
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gEPOReasons, , True)
    
    For i = 1 To AccountingFrame.UBound
        'AccountingFrame(i).Move AccountingFrame(1).left, AccountingFrame(1).Top
        AccountingFrame(i).Move 0, 930
    Next
    
    For i = 2 To EstimatingFrame.UBound
        EstimatingFrame(i).Move EstimatingFrame(1).left, EstimatingFrame(1).Top
    Next
    
    Call ReadData
    
    Call gTOC.Select(1, 1)
    
    Me.BackColor = vbButtonFace
Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
'Stop: Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gEPOReasons)
    
    mIntacctEntitiesLoaded = False
    mIntacctTranTypeListsLoaded = False
    mIntacctEstimateTypeListsLoaded = False
    mIntacctItemsLoaded = False
    
End Sub

Private Sub gCustomFields_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    cmdEditField.Enabled = gCustomFields.Cell(flexcpText, Row, gCustomFields.ColIndex("DataType")) = "Selection List"
    
    If Col = 0 Then
    Else
        Cancel = True
    End If
End Sub

Private Sub gCustomFields_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh
    Dim s As String
    
    s = ""
    s = s & "UPDATE JobCustomFieldDefs" & vbCrLf
    s = s & "SET Category=" & DbQuote(Str, gCustomFields.EditText) & vbCrLf
    s = s & "WHERE Name=" & DbQuote(Str, gCustomFields.TextMatrix(Row, 1)) & vbCrLf
    Call HFApp.SqlExec(s)
            
Exit Sub
eh: Call errHandler(SRCFILE & "gCustomFields_ValidateEdit")
End Sub

Private Sub gFormatting_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call ApplyFormating
End Sub

Private Sub gFormatting_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = 0
End Sub

Private Sub gFormatting_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim c As Long
    c = gFormatting.Cell(flexcpBackColor, Row, Col)
    If VBChooseColor(c, False, , , Me.hwnd) Then
        gFormatting.Cell(flexcpBackColor, Row, Col) = c
        Call ApplyFormating
    End If
End Sub

Private Sub ApplyFormating()
    Dim r As Long
    With gFormatting
        For r = 0 To 3
        
            'stupid flexgrid reads 0 as white - substitue something almost black
            If .Cell(flexcpBackColor, r, 3) = 0 Then .Cell(flexcpBackColor, r, 3) = 1
            If .Cell(flexcpBackColor, r, 2) = 0 Then .Cell(flexcpBackColor, r, 2) = 1
        
            .Cell(flexcpForeColor, r, 0) = .Cell(flexcpBackColor, r, 3)
            .Cell(flexcpBackColor, r, 0) = .Cell(flexcpBackColor, r, 2)
            .Cell(flexcpFontBold, r, 0) = InStr(1, .Cell(flexcpText, r, 1), "Bold")
            .Cell(flexcpFontItalic, r, 0) = InStr(1, .Cell(flexcpText, r, 1), "Italic")
        Next
    End With
End Sub


Private Sub lblReplacementParameters_Click()
    Dim s As String
    
    
    s = ""
    s = s & "You can use fields in your notification messages. When" & vbCrLf
    s = s & "the message is delivered the field will be replaced with" & vbCrLf
    s = s & "the actual value." & vbCrLf & vbCrLf
    s = s & "  Available fields:" & vbCrLf
    s = s & "    <%SenderName%>" & vbCrLf
    s = s & "    <%SenderCompany%>" & vbCrLf
    s = s & "    <%SenderFax%>" & vbCrLf
    s = s & "    <%SenderPhone%>" & vbCrLf
    s = s & "    <%SenderEmail%>" & vbCrLf
    s = s & "    <%RecipientName%>" & vbCrLf
    s = s & "    <%RecipientCompany%>" & vbCrLf
    s = s & "    <%RecipientFax%>" & vbCrLf
    s = s & "    <%RecipientPhone%>" & vbCrLf
    s = s & "    <%RecipientEmail%>" & vbCrLf
    s = s & "    <%Quote%>" & vbCrLf
    s = s & "    <%Job%>" & vbCrLf
    s = s & "    <%JobDescription%>" & vbCrLf
'    s = s & "    <%RFP%>" & vbCrLf
    s = s & "    <%POIndex%>" & vbCrLf
    s = s & "    <%POIndexDescription%>" & vbCrLf
    s = s & "    <%PONumber%>" & vbCrLf
    s = s & "    <%POCancellationReason%>" & vbCrLf
    s = s & "    <%Vendor%>" & vbCrLf
'    s = s & "    <%BidGUID%>" & vbCrLf
    s = s & "    <%VendorName%>" & vbCrLf
    s = s & "    <%Model%>" & vbCrLf
    s = s & "    <%MunicipalAddress%>" & vbCrLf
    s = s & "    <%LegalAddress%>" & vbCrLf
    s = s & "    <%County%>" & vbCrLf
    s = s & "    <%Township%>" & vbCrLf
    s = s & "    <%Community%>" & vbCrLf
    s = s & "    <%Phase%>" & vbCrLf
    MsgBox s, vbInformation, App.ProductName
    
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'OK
            If SaveData Then
                Unload Me
            End If
            
        Case 1 'cancel
            Cancel = True
            Unload Me
    End Select
End Sub

Private Sub ReadSupplyChain()
With HFApp.Options
    
    Call LoadComboBox(cboPurchasingManager, HFApp.Databases(dbHomefront), "select pmname,pm,0 from tblprojectmanager where purchaser=1")
    Call SetComboBoxListIndex(cboPurchasingManager, , .ValueByName("DefaultPurchasingManager"))
    txtFieldPOPrefix.Text = .ValueByName("FieldPOPrefix")
        
    chkFieldPOField(0).Value = IIf("1" = .ValueByName("Detail_Description"), vbChecked, vbUnchecked)
    chkFieldPOField(1).Value = IIf("1" = .ValueByName("Detail_Quantity"), vbChecked, vbUnchecked)
    chkFieldPOField(2).Value = IIf("1" = .ValueByName("Detail_Rate"), vbChecked, vbUnchecked)
    chkFieldPOField(3).Value = IIf("1" = .ValueByName("Detail_Job"), vbChecked, vbUnchecked)
    chkFieldPOField(4).Value = IIf("1" = .ValueByName("Detail_Extra"), vbChecked, vbUnchecked)
    chkFieldPOField(5).Value = IIf("1" = .ValueByName("Detail_CostCode"), vbChecked, vbUnchecked)
    chkFieldPOField(6).Value = IIf("1" = .ValueByName("Detail_Category"), vbChecked, vbUnchecked)
    chkFieldPOField(7).Value = IIf("1" = .ValueByName("Detail_TaxGroup"), vbChecked, vbUnchecked)
    chkUseDetailedFieldPOView.Value = IIf("1" = .ValueByName("ViewPurchasingManagerDetailsForm"), vbChecked, vbUnchecked)
    
    txtWebAttachmentsFolder.Text = .ValueByName("WebAttachmentsFolder")
    
End With
End Sub

Private Sub ReadData()
On Error GoTo eh
    
    Dim i  As Long
    Dim rs As Recordset
    Dim s As String
    
    'load tax groups into combo boxes
    Set rs = HFApp.SqlExec("SELECT '' TaxGroup, '' Description UNION ALL SELECT TaxGroup, Description FROM TaxGroups where DivisionID = " & HFApp.DivisionID)
    For i = 0 To 5
        cboTaxGroup(i).Clear
    Next
    While Not rs.EOF
        For i = 0 To 5
            cboTaxGroup(i).AddItem "" & rs(0)
        Next
        rs.MoveNext
    Wend
    
    mDivCodeChangeWarned = False
    
    'load variance cats into combo boxes
    Call LoadComboBox(cboVarianceCat, HFApp.Databases(dbHomefront), "SELECT '','',0 UNION ALL SELECT description,category,0 FROM standardcategories where isvariance=1 and DivisionID = " & HFApp.DivisionID)
    Call SetComboBoxListIndex(cboVarianceCat, , HFApp.Options.ValueByName("DefaultVarianceCategory"))
    chkUseVarianceReporting.Value = IIf(HFApp.Options.ValueByName("UseVarianceReporting") <> "False", vbChecked, vbUnchecked)
    
    
    Call ReadBuildProOptions
    Call ReadRoleDefinitions
    Call ReadDepartmentApprovals
    Call ReadSupplyChain
    
    With HFApp.Options
    
        cboBuilderType.Clear
        cboBuilderType.AddItem "Commercial"
        cboBuilderType.AddItem "Homebuilder"
        s = HFApp.Options.ValueByName("BuilderType")
        If Not IsIn(s, "Commercial", "Homebuilder") Then s = "Homebuilder"
        cboBuilderType.Text = s

        optUseAltCostCodesForCO(1).Value = .ValueByName("UseAltCostCodesForCO") = "True"
    
        chkPurchasingRequiresSalesApproval.Value = IIf(.ValueByName("PurchasingRequiresSalesApproval") = "True", vbChecked, vbUnchecked)
        If HFApp.SqlExec("select top 1 isnull(MaxVendorPricing,0) from system_setup where id=" & HFApp.DivisionID)(0) = True Then
            chkUseMaxVendorPricing.Value = 1
        Else
            chkUseMaxVendorPricing.Value = 0
        End If
        
        
        
        
        
        txtCompanyName.Text = .Value(CompanyName)
        txtContactPerson.Text = .Value(ContactPerson)
        txtCompanyAddr1.Text = .Value(address1)
        txtCompanyAddr2.Text = .Value(address2)
        txtCompanyCity.Text = .Value(city)
        Call SetComboBoxListIndex(cboCompanyCountry, .Value(Country))
        Call SetComboBoxListIndex(cboCompanyProvince, .Value(province))
        txtCompanyPostalCode.Text = .Value(zip)
        txtCompanyPhone.Text = .Value(phone)
        txtCompanyFax.Text = .Value(fax)
        txtCompanyEmail.Text = .Value(Email)
        chkEnforcePasswordComplexity.Value = IIf("True" = .ValueByName("EnforcePasswordComplexity"), vbChecked, vbUnchecked)
        
        
        Call ReadDivisions
            
        Call ReadDocumentManagement
        Call ReadEstimatingOptions
        Call ReadAccountingSystems
        Call ReadAccountingImportOptions
        Call ReadSalesSystems
        Call ReadCommunityStandards
        
        chkEstimatingWorksheetsOnly.Value = IIf(.Value(UseEstimatingWorksheetsOnly) = True, vbChecked, vbUnchecked)

        
        chkUseTaxGroups.Value = IIf(.Value(TaxGroupRequired) = "true", vbChecked, vbUnchecked)
        Call SetListIndex(cboTaxGroup(0), , .Value(TaxGroupLabour))
        Call SetListIndex(cboTaxGroup(1), , .Value(TaxGroupMaterial))
        Call SetListIndex(cboTaxGroup(2), , .Value(TaxGroupSubContract))
        Call SetListIndex(cboTaxGroup(3), , .Value(TaxGroupEquipment))
        Call SetListIndex(cboTaxGroup(4), , .Value(TaxGroupOverhead))
        Call SetListIndex(cboTaxGroup(5), , .Value(TaxGroupOther))
        
        
        txtIncentiveItem.tag = .Value(IncentivePhase) & Chr(1) & .Value(IncentiveItem) & Chr(1) & .Value(IncentiveItemDescription)
        txtIncentiveItem.Text = .Value(IncentivePhase) & "/" & .Value(IncentiveItem) & " -- " & .Value(IncentiveItemDescription)
        
        txtSelectAtTakeoffItem.tag = .Value(SelectAtTakeoffPhase) & Chr(1) & .Value(SelectAtTakeoffItem) & Chr(1) & .Value(SelectAtTakeoffDescription)
        txtSelectAtTakeoffItem.Text = .Value(SelectAtTakeoffPhase) & "/" & .Value(SelectAtTakeoffItem) & " -- " & .Value(SelectAtTakeoffDescription)
        txtIncentiveCostPercent.Text = Val("" & .Value(IncentiveCostPercent)) & "%"
        txtLandItem.tag = .Value(LandPhase) & Chr(1) & .Value(LandItem) & Chr(1) & .Value(LandItemDescription)
        txtLandItem.Text = .Value(LandPhase) & "/" & .Value(LandItem) & " -- " & .Value(LandItemDescription)
        
        
        chkSellingPriceIncludesTax.Value = IIf(.Value(includetax), vbChecked, vbUnchecked)
        chkAdjustSelling.Value = IIf(.Value(AdjustSellPriceWithCosts), vbChecked, vbUnchecked)
        chkAdjustIncentive.Value = IIf(.ValueByName("AdjustIncentiveRetailWithCosts") <> "False", vbChecked, vbUnchecked)
        
        chkLockPostedSheets.Value = IIf(.Value(LockPostedSalesWorksheets), vbChecked, vbUnchecked)
        
        chkLetPurchaserChangeSaleQty.Value = IIf("" & .ValueByName("LetPurchaserChangeSaleQty") = "True", vbChecked, vbUnchecked)
        chkUseComponents.Value = IIf("" & .ValueByName("UseComponents") = "True", vbChecked, vbUnchecked)
        
        chkZeroRateOnRefeshCosts.Value = IIf(.Value(ZeroRateOnRefreshCosts), vbChecked, vbUnchecked)
        chkZeroRateOnChangeVendor.Value = IIf(.Value(ZeroRateOnChangeVendor), vbChecked, vbUnchecked)
        
        
        chkMarketingWorkSheetAutoRefreshCosts.Value = IIf(.Value(MarketingWorkSheetAutoRefreshCosts), vbChecked, vbUnchecked)
        chkCanAddContractItemsFromPurchasing.Value = IIf(.Value(CanAddContractItemsFromPurchasing) = "True", vbChecked, vbUnchecked)
        chkUsePricingFromEstimating.Value = IIf(.Value(UsePricingFromEstimating), vbUnchecked, vbChecked)
        
        
        
        optPOSequential.Value = False
        optPOByJobPOIndex.Value = False
        optPOByJob.Value = False
        Select Case "" & .ValueByName("PONumberingStyle")
            Case "Job":            optPOByJob.Value = True
            Case "JobPOIndex":     optPOByJobPOIndex.Value = True
            Case Else:             optPOSequential.Value = True
        End Select
        txtPOSegment(1) = .Value(POFormatSegment1)
        txtPOSegment(2) = .Value(POFormatSegment2)
        txtPOSegment(3) = .Value(POFormatSegment3)
        txtPOSegment(4) = .ValueByName("POFormatSegment4")
        txtPOSegment(5) = .ValueByName("POFormatSegment5")
        txtPOSegment(6) = .ValueByName("POFormatSegment6")
        txtPOSegment(7) = .ValueByName("POFormatSegment7")
        txtPOSegment(8) = .ValueByName("POFormatSegment8")
        txtPOPrefix.Text = "" & .Value(POPrefix)
        txtNextPO.Text = Val("" & HFApp.SqlExec("SELECT MAX(PONumber) FROM PONUmbers")(0)) + 1
        Call optPOByJob_Click
        
        
        cboZeroQtyTakeoffMode.AddItem "Prompt me"
        cboZeroQtyTakeoffMode.AddItem "Always skip"
        cboZeroQtyTakeoffMode.AddItem "Always add"
        cboZeroQtyTakeoffMode.ListIndex = Val(.Value(ZeroQtyTakeoffs))
        
        
        
        
        'notices and delivery
        txtNotification(0).Text = .Value(Msg_SendPO_Subject)
        txtNotification(1).Text = .Value(Msg_SendPO_Body)
        txtNotification(2).Text = .Value(Msg_CancelPO_Subject)
        txtNotification(3).Text = .Value(Msg_CancelPO_Body)
        
        
        
    End With
    
    Call ReadEmailConfig
    
    With gFormatting
        .Cell(flexcpText, 0, 1) = HFApp.Options.Value(Format_QtyRateEQZero_FontStyle)
        .Cell(flexcpText, 1, 1) = HFApp.Options.Value(Format_QtyRateLTZero_FontStyle)
        .Cell(flexcpText, 2, 1) = HFApp.Options.Value(Format_NonSysRate_FontStyle)
        .Cell(flexcpText, 3, 1) = HFApp.Options.Value(Format_InvalidData_FontStyle)
        
        .Cell(flexcpBackColor, 0, 2) = HFApp.Options.Value(Format_QtyRateEQZero_BackColor)
        .Cell(flexcpBackColor, 1, 2) = HFApp.Options.Value(Format_QtyRateLTZero_BackColor)
        .Cell(flexcpBackColor, 2, 2) = HFApp.Options.Value(Format_NonSysRate_BackColor)
        .Cell(flexcpBackColor, 3, 2) = HFApp.Options.Value(Format_InvalidData_BackColor)

        .Cell(flexcpBackColor, 0, 3) = HFApp.Options.Value(Format_QtyRateEQZero_ForeColor)
        .Cell(flexcpBackColor, 1, 3) = HFApp.Options.Value(Format_QtyRateLTZero_ForeColor)
        .Cell(flexcpBackColor, 2, 3) = HFApp.Options.Value(Format_NonSysRate_ForeColor)
        .Cell(flexcpBackColor, 3, 3) = HFApp.Options.Value(Format_InvalidData_ForeColor)
    End With
    
    
    Call ApplyFormating
    
    Call ReadCustomProperties

Exit Sub
eh: Call errHandler(SRCFILE & "ReadData")
End Sub

Private Sub ReadEmailConfig()
    With HFApp.Options
        chkCCPrjMgr.Value = IIf(.ValueByName("SendPO_CCPrjMgr") <> "False", vbChecked, vbUnchecked)
                
        If .ValueByName("SendPO_ViaSMTP") = "True" Then
            optEmailProtocol(1).Value = True
        ElseIf .ValueByName("SendPO_Via365SMTP") = "True" Then
            optEmailProtocol(2).Value = True
        Else
            optEmailProtocol(0).Value = True
        End If
        
        'smtp
        txtSMTPServer.Text = .ValueByName("SendPO_SMTPServer")
        chkEmailAuthentication.Value = IIf(.ValueByName("Email_Authenticate") = "True", 1, 0)
        chkSSL.Value = IIf(.ValueByName("Email_SSL") = "True", 1, 0)
        chkSTARTTLS.Value = IIf(.ValueByName("Email_STARTTLS") = "True", 1, 0)
        txtEmailUserID.Text = .ValueByName("Email_User")
        txtEmailPwd.Text = .ValueByName("Email_Pswd")
        optSendFromCurrentUsersEmail.Value = IIf(.ValueByName("SendPO_FromCurrentUsersEmail") = "True", 1, 0)
        optSendFromCorporateEmail.Value = Not optSendFromCurrentUsersEmail.Value
        txtEmailAddress.Text = .Value(Estimator_Email)
        
        '365
        txt365AppID.Text = .ValueByName("Smtp365EmailAppID")
        txt365TenantID.Text = .ValueByName("Smtp365EmailTenantID")
        txt365Secret.Text = .ValueByName("Smtp365EmailSecret")
        txt365User.Text = .ValueByName("Smtp365EmailUser")
        txt365Pswd.Text = .ValueByName("Smtp365EmailPswd")
         
        frmEmailConfig(1).Caption = ""
        frmEmailConfig(2).Caption = ""
        frmEmailConfig(2).Move frmEmailConfig(1).left, frmEmailConfig(1).Top
        Call optEmailProtocol_Click(0)
        
        
    End With
End Sub

Private Sub SaveEmailConfig()
    With HFApp.Options
        .ValueByName("SendPO_CCPrjMgr") = chkCCPrjMgr.Value = vbChecked
        .ValueByName("SendPO_ViaOutlook") = optEmailProtocol(0).Value
        .ValueByName("SendPO_ViaSMTP") = optEmailProtocol(1).Value
        .ValueByName("SendPO_Via365SMTP") = optEmailProtocol(2).Value
        
        'smtp
        .ValueByName("SendPO_SMTPServer") = txtSMTPServer.Text
        .ValueByName("Email_Authenticate") = IIf(chkEmailAuthentication.Value = vbChecked, "True", "False")
        .ValueByName("Email_SSL") = IIf(chkSSL.Value = vbChecked, "True", "False")
        .ValueByName("EnableSSL") = IIf(chkSSL.Value = vbChecked, "True", "False")
        .ValueByName("Email_STARTTLS") = IIf(chkSTARTTLS.Value = vbChecked, "True", "False")
        .ValueByName("Email_User") = txtEmailUserID.Text
        .ValueByName("SMTP_User") = txtEmailUserID.Text
        .ValueByName("SMTP_Password") = txtEmailPwd.Text
        .ValueByName("Email_Pswd") = txtEmailPwd.Text
        .ValueByName("SendPO_FromCurrentUsersEmail") = optSendFromCurrentUsersEmail.Value
        .Value(Estimator_Email) = txtEmailAddress.Text
        
        '365
        .ValueByName("Smtp365EmailAppID") = txt365AppID.Text
        .ValueByName("Smtp365EmailTenantID") = txt365TenantID.Text
        .ValueByName("Smtp365EmailSecret") = txt365Secret.Text
        .ValueByName("Smtp365EmailUser") = txt365User.Text
        .ValueByName("Smtp365EmailPswd") = txt365Pswd.Text
        
    End With
End Sub


Private Sub ReadSalesSystems()

    cboSalesSystem.AddItem "No sales system integration"
    cboSalesSystem.AddItem "HomeFront Profit Builder"
    cboSalesSystem.AddItem "Builder 1440"
    cboSalesSystem.ListIndex = Val(HFApp.Options.Value(SalesSystem))

    txtBuilderCode.Text = HFApp.Options.Value(BuilderCode)
    txtWebAdminUser.Text = HFApp.Options.Value(WebAdminUser)
    txtWebUploadPath.Text = HFApp.Options.Value(WebUploadPath)
    txtWebUploadUser.Text = HFApp.Options.Value(WebUploadUser)
    txtWebUploadPswd.Text = HFApp.Options.Value(WebUploadPswd)
    chkAppendUOMtoOptionDesc.Value = IIf(HFApp.Options.ValueByName("AppendUOMtoOptionDesc") = "True", vbChecked, vbUnchecked)

    optWebUploadbyCommunity.Value = HFApp.Options.ValueByName("WebUploadByCommunity") <> "False"
    optWebUploadbyDivision.Value = Not optWebUploadbyCommunity.Value



    cboCrmEnvironment.Clear
    cboCrmEnvironment.AddItem "User Acceptance Testing (UAT)"
    cboCrmEnvironment.AddItem "Production"
    Call SetComboBoxListIndex(cboCrmEnvironment, HFApp.Options.ValueByName("CRMEnvironment"))
    txtCrmClientID.Text = HFApp.Options.ValueByName("CRMClientID")
    txtCrmApiKey.Text = HFApp.Options.ValueByName("CRMApiKey")




End Sub

Private Sub ReadCommunityStandards()
    Dim rs As Recordset
    
    chkUseCommunityStandards.Value = IIf(HFApp.Options.ValueByName("UseCommunityStandards") = "True", vbChecked, vbUnchecked)
    chkCommunityStandardsArePhaseSpecific.Value = IIf(HFApp.Options.ValueByName("CommunityStandardsArePhaseSpecific") = "True", vbChecked, vbUnchecked)
    With gStdItems
        .Rows = 1
        Set rs = HFApp.SqlExec("select s.phase,s.item,i.description from communitystandarditems s left outer join tblphaseitem i on (i.phase=s.phase and i.item=s.item) where i.DivisionID = " & HFApp.DivisionID & " order by 1,2,3", dbHomefront)
        While Not rs.EOF
            .AddItem "" & rs(0) & vbTab & rs(1) & vbTab & rs(2)
            rs.MoveNext
        Wend
        .AddItem ""
        Call .AutoSize(0, 2)
    End With
    
End Sub

Private Sub SaveCommunityStandards()
On Error Resume Next
    Dim i As Long
    HFApp.Options.ValueByName("UseCommunityStandards") = chkUseCommunityStandards.Value = vbChecked
    HFApp.Options.ValueByName("CommunityStandardsArePhaseSpecific") = chkCommunityStandardsArePhaseSpecific.Value = vbChecked
    With gStdItems
    
        For i = .Rows - 2 To 1 Step -1
            If .RowHidden(i) Then
                Call HFApp.SqlExec("delete from CommunityStandardItems where phase=" & DbQuote(Str, .TextMatrix(i, 0)) & " and item=" & DbQuote(Str, .TextMatrix(i, 1)), dbHomefront)
                Call HFApp.SqlExec("delete from CommunityStandards where stdphase=" & DbQuote(Str, .TextMatrix(i, 0)) & " and stditem=" & DbQuote(Str, .TextMatrix(i, 1)), dbHomefront)
                Call .RemoveItem(i)
            Else
                Call HFApp.SqlExec("insert into CommunityStandardItems(phase,item) values(" & DbQuote(Str, .TextMatrix(i, 0)) & "," & DbQuote(Str, .TextMatrix(i, 1)) & ")", dbHomefront)
            End If
        Next
    
    End With
    
End Sub


Private Sub ReadAccountingImportOptions()
    Dim s As String
    
    optManageActiveVendors(1).Value = HFApp.Options.ValueByName("ManageActiveVendors") = "True"
    optManageActiveVendors(0).Value = Not optManageActiveVendors(1).Value


    s = HFApp.Options.ValueByName("ExcludedVendorTypes")
    s = Replace(s, "'", "")
    s = Replace(s, ",", vbCrLf)
    txtExcludedVendorTypes.Text = s

End Sub


Private Sub SaveAccountingImportOptions()
    Dim s As String
    
    HFApp.Options.ValueByName("ManageActiveVendors") = "" & optManageActiveVendors(1).Value
    
    s = txtExcludedVendorTypes.Text
    s = Replace(s, vbCrLf, ",")
    
    s = Replace(s, "     ,", ",")
    s = Replace(s, "    ,", ",")
    s = Replace(s, "   ,", ",")
    s = Replace(s, "  ,", ",")
    s = Replace(s, " ,", ",")
    
    s = Replace(s, ",     ", ",")
    s = Replace(s, ",    ", ",")
    s = Replace(s, ",   ", ",")
    s = Replace(s, ",  ", ",")
    s = Replace(s, ", ", ",")
    
    s = Replace(s, ",,,,,", ",")
    s = Replace(s, ",,,,", ",")
    s = Replace(s, ",,,", ",")
    s = Replace(s, ",,", ",")
    
    s = "'" & Replace(s, ",", "','") & "'"
    
    HFApp.Options.ValueByName("ExcludedVendorTypes") = s
    
    
End Sub

Private Sub ReadAccountingSystems()
    Dim i As Long
    Dim csv As String
    Dim s As String
    Dim rs As Recordset
    
    Dim tax As String
    Dim non As String
    
    For i = 1 To AccountingFrame.UBound
        AccountingFrame(i).BorderStyle = 0
    Next
    
    
    With HFApp.Options
        chkPostSummarizedBudgets.Value = IIf(.Value(PostSummarizedBudgets), vbChecked, vbUnchecked)
        chkPostAssembliesAsSalesInvoices.Value = IIf(.ValueByName("PostAssembliesAsSalesInvoices") = "True", vbChecked, vbUnchecked)
        chkPostAssembliesAsSalesInvoices.Visible = .Value(AccountingSystem) = asQuickBooks
        chkPostJCExtraAsSubJob.Value = IIf(.ValueByName("PostJCExtraAsSubJob") = "True", vbChecked, vbUnchecked)
        chkPostPOQtyToAccounting.Value = IIf(.ValueByName("PostPOQtyToAccounting") = "True", vbChecked, vbUnchecked)
    
        chkPostJCExtraAsSubJob.Top = chkPostAssembliesAsSalesInvoices.Top
    
    End With
    
    
    
    With cboAccountingSystem
        .AddItem "No accounting integration":              .ItemData(.NewIndex) = AccountingSystems.asNone
    
        .AddItem "Sage 300 Construction and Real Estate":  .ItemData(.NewIndex) = asTimberline
        .AddItem "Sage 100 Contractor":                    .ItemData(.NewIndex) = asMasterBuilder
        .AddItem "Sage 50 Accounting":                     .ItemData(.NewIndex) = asSimply
        .AddItem "Sage Intacct":                           .ItemData(.NewIndex) = asIntacct
        .AddItem "QuickBooks":                             .ItemData(.NewIndex) = asQuickBooks
        .AddItem "Quickbooks Online":                      .ItemData(.NewIndex) = asQuickbooksOnline
        .AddItem "Xero":                                   .ItemData(.NewIndex) = asXero
            
        Call SetListIndex(cboAccountingSystem, Val(HFApp.Options.Value(AccountingSystem)))
    End With
    
    
    s = "select bookofaccount from divisions where divisionid=" & HFApp.DivisionID
    s = "" & HFApp.SqlExec(s, dbHomefront)(0)
    If s = "" Then s = "Main"
    Call SetComboBoxListIndex(cboBookOfAccount, s)
    If cboBookOfAccount.ListIndex = -1 Then cboBookOfAccount.ListIndex = 0
    
                
    'Sage Intacct
    txtIntacctCompanyID.Text = HFApp.Options.ValueByName("IntacctCompanyID")
    txtIntacctUID.Text = HFApp.Options.ValueByName("IntacctUID")
    txtIntacctPWD.Text = HFApp.Options.ValueByName("IntacctPWD")
    
    Call cboIntacctEntity.AddItem(HFApp.Options.ValueByName("IntacctEntity"))
    Call cboIntacctPOType.AddItem(HFApp.Options.ValueByName("IntacctPOType"))
    Call cboIntacctSubContractType.AddItem(HFApp.Options.ValueByName("IntacctSubContractType"))
    Call cboIntacctInvoiceType.AddItem(HFApp.Options.ValueByName("IntacctInvoiceType"))
    Call cboIntacctEstimateType.AddItem(HFApp.Options.ValueByName("IntacctEstimateType"))
    Call cboIntacctPOCOType.AddItem(HFApp.Options.ValueByName("IntacctPOCOType"))
    Call cboIntacctSubContractCOType.AddItem(HFApp.Options.ValueByName("IntacctSubcontractCOType"))
    Call cboIntacctItem(0).AddItem(HFApp.Options.ValueByName("IntacctOneTimeItemID"))
    Call cboIntacctItem(1).AddItem(HFApp.Options.ValueByName("IntacctGstItemID"))
    cboIntacctEntity.ListIndex = 0
    cboIntacctPOType.ListIndex = 0
    cboIntacctPOCOType.ListIndex = 0
    cboIntacctSubContractType.ListIndex = 0
    cboIntacctSubContractCOType.ListIndex = 0
    cboIntacctInvoiceType.ListIndex = 0
    cboIntacctEstimateType.ListIndex = 0
    cboIntacctItem(0).ListIndex = 0
    cboIntacctItem(1).ListIndex = 0
    mIntacctOneTimeItemUOM = HFApp.Options.ValueByName("IntacctOneTimeItemUOM")
    mIntacctGstItemUOM = HFApp.Options.ValueByName("IntacctGstItemUOM")
    txtIntacctBudgetGL.Text = HFApp.Options.ValueByName("IntacctEstimateGLBudgetID")
    With cboPOReferenceFld
        .Clear
        .AddItem "Job Number"
        .AddItem "PO Number"
        .ListIndex = IIf(HFApp.Options.ValueByName("IntacctPOReferenceFld") = "PO Number", 1, 0)
    End With
    With cboInvReferenceFld
        .Clear
        .AddItem "Job Number"
        .AddItem "Job Description"
        .AddItem "Invoice Description"
        .AddItem "PO Number"
        .ListIndex = Decode(HFApp.Options.ValueByName("IntacctInvReferenceFld"), "Job Number", 0, "Invoice Description", 2, 1)
    End With
    
    
    
    
    'Xero
'    txtXeroJobName.Text = HFApp.Options.ValueByName("XeroJobName")
'    txtXeroCostCodeName.Text = HFApp.Options.ValueByName("XeroCostCodeName")
'    If txtXeroJobName.Text = "" Then txtXeroJobName.Text = "Job"
'    If txtXeroCostCodeName.Text = "" Then txtXeroCostCodeName.Text = "Cost Code"
    
    'Timberline
    txtTLFolder.Text = HFApp.Options.Value(Timberline_Data_Path)
    txtTLARFolder.Text = HFApp.Options.Value(TL_AR_Data_Path)
    txtTLUser.Text = HFApp.Options.Value(Timberline_UID)
    txtTLPswd.Text = HFApp.Options.Value(Timberline_PWD)
    
    
    
    'Master Builder
    s = HFApp.Options.ValueByName("Sage100APILevel")
    If s = "" Then s = "v19"
    Call SetComboBoxListIndex(cboMBAPILevel, s)
    Call cboMBAPILevel_Click
    
    'v18&19 - datafolder was a full path to a folder
    'with v20 it changed to a sql server instance name
    s = FileDrive(HFApp.Options.Value(MasterBuilderDataFolder))
    If s = "" Then s = "c:\"
    Call SetComboBoxListIndex(cboMBDrive, s)
    cboMBInstance.Text = HFApp.Options.Value(MasterBuilderDataFolder)
    
    cboMBCompany.Text = HFApp.Options.Value(MasterBuilderCompany)
    txtMBUser.Text = HFApp.Options.Value(MasterBuilderUID)
    txtMBPswd.Text = HFApp.Options.Value(MasterBuilderPWD)
    chkmbUseSubAcct = IIf(HFApp.Options.ValueByName("UseJobAsSubAcct") = "True", 1, 0)
    
    'Quick Books
    txtQuickBooksFile.Text = HFApp.Options.Value(QuickBooksDataFile)
    cboQuickBooksVersion.AddItem "US"
    cboQuickBooksVersion.AddItem "CA"
    Call SetListIndex(cboQuickBooksVersion, , HFApp.Options.ValueByName("AccountingVersion"))
    If cboQuickBooksVersion.ListIndex = -1 Then cboQuickBooksVersion.ListIndex = 0
    optQBJobHeirarchy(0).Value = HFApp.Options.ValueByName("QuickBooksJobStyle") = "Simple"
    
    'Quick Books Online
    cboQuickBooksOnlineVersion.AddItem "US"
    cboQuickBooksOnlineVersion.ListIndex = 0
    optQBOJobHeirarchy(0).Value = HFApp.Options.ValueByName("QuickBooksJobStyle") = "Simple"
    
    
    'Simply Accounting
    txtSimplyFile.Text = HFApp.Options.Value(SimplyDataFile)
    txtSimplyUser.Text = HFApp.Options.ValueByName("SimplyUID")
    txtSimplyPswd.Text = HFApp.Options.ValueByName("SimplyPWD")
    Call LoadComboBox(cboSimplyInternalCustomer, HFApp.Databases(dbHomefront), "select description,arcustomer,0 from arcustomers where internalcustomer=1")
    Call SetComboBoxListIndex(cboSimplyInternalCustomer, , HFApp.Options.ValueByName("SimplyInternalCustomer"))

    

End Sub

Private Sub SaveSalesSystem()
    
    HFApp.Options.Value(SalesSystem) = Max(cboSalesSystem.ListIndex, 0)
        
    HFApp.Options.Value(BuilderCode) = txtBuilderCode.Text
    HFApp.Options.Value(WebAdminUser) = txtWebAdminUser.Text
    HFApp.Options.Value(WebUploadPath) = txtWebUploadPath.Text
    HFApp.Options.Value(WebUploadUser) = txtWebUploadUser.Text
    HFApp.Options.Value(WebUploadPswd) = txtWebUploadPswd.Text
    HFApp.Options.ValueByName("AppendUOMtoOptionDesc") = chkAppendUOMtoOptionDesc.Value = vbChecked
    HFApp.Options.ValueByName("WebUploadByCommunity") = "" & optWebUploadbyCommunity.Value
    
    HFApp.Options.ValueByName("CRMEnvironment") = cboCrmEnvironment.Text
    HFApp.Options.ValueByName("CRMClientID") = txtCrmClientID.Text
    HFApp.Options.ValueByName("CRMApiKey") = txtCrmApiKey.Text
    
End Sub


Private Sub SaveAccountingSystem()
    
    HFApp.Options.Value(AccountingSystem) = Max(cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex), 0)
    HFApp.Options.Value(Use_Timberline) = HFApp.Options.Value(AccountingSystem) = asTimberline
    
    HFApp.SqlExec "update divisions set bookofaccount=" & DbQuote(Str, Me.cboBookOfAccount.Text) & " where divisionid=" & HFApp.DivisionID
    
    'only implemented in qb
    HFApp.Options.ValueByName("PostAssembliesAsSalesInvoices") = False
    
    

    HFApp.Options.Value(WarnForecastsWhenRefreshingPOs) = chkWarnForecast.Value = vbChecked
    HFApp.Options.Value(PostSummarizedBudgets) = chkPostSummarizedBudgets.Value = vbChecked
    HFApp.Options.Value(PostSummarizedPOs) = False
    HFApp.Options.ValueByName("PostJCExtraAsSubJob") = chkPostJCExtraAsSubJob.Value = vbChecked
    HFApp.Options.ValueByName("PostPOQtyToAccounting") = chkPostPOQtyToAccounting.Value = vbChecked
    
    
    Select Case cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex)
        
        Case 0 'none
            HFApp.Options.Value(Timberline_Data_Path) = ""
            HFApp.Options.Value(TL_AR_Data_Path) = ""
            HFApp.Options.Value(Timberline_UID) = ""
            HFApp.Options.Value(Timberline_PWD) = ""
            HFApp.Options.Value(MasterBuilderCompany) = ""
            HFApp.Options.Value(MasterBuilderDataFolder) = ""
            HFApp.Options.Value(MasterBuilderUID) = ""
            HFApp.Options.Value(MasterBuilderPWD) = ""
            
            
        Case 1 'timberline
            HFApp.Options.Value(Timberline_Data_Path) = txtTLFolder.Text
            HFApp.Options.Value(TL_AR_Data_Path) = txtTLARFolder.Text
            HFApp.Options.Value(Timberline_UID) = txtTLUser.Text
            HFApp.Options.Value(Timberline_PWD) = txtTLPswd.Text
            HFApp.Options.Value(MasterBuilderCompany) = ""
            HFApp.Options.Value(MasterBuilderDataFolder) = ""
            HFApp.Options.Value(MasterBuilderUID) = ""
            HFApp.Options.Value(MasterBuilderPWD) = ""
        
        
        Case 2 'masterbuilder
            HFApp.Options.Value(Timberline_Data_Path) = ""
            HFApp.Options.Value(TL_AR_Data_Path) = ""
            HFApp.Options.Value(Timberline_UID) = txtMBUser.Text
            HFApp.Options.Value(Timberline_PWD) = txtMBPswd.Text
            HFApp.Options.Value(MasterBuilderCompany) = cboMBCompany.Text
            
            HFApp.Options.ValueByName("Sage100APILevel") = cboMBAPILevel.Text
            If cboMBAPILevel.Text = "v20" Then
                HFApp.Options.Value(MasterBuilderDataFolder) = cboMBInstance.Text
            Else
                HFApp.Options.Value(MasterBuilderDataFolder) = PathAppend(left(cboMBDrive.Text, 1) & ":", "MB7", cboMBCompany.Text)
            End If
            HFApp.Options.ValueByName("MasterBuilderEdition") = "US" 'cboMBEdition.Text
            HFApp.Options.Value(MasterBuilderUID) = txtMBUser.Text
            HFApp.Options.Value(MasterBuilderPWD) = txtMBPswd.Text
            HFApp.Options.ValueByName("UseJobAsSubAcct") = IIf(chkmbUseSubAcct.Value = 1, "True", "False")
            

        Case 10 'quickbooks online
            HFApp.Options.ValueByName("AccountingVersion") = cboQuickBooksOnlineVersion.Text
            HFApp.Options.ValueByName("QuickBooksJobStyle") = IIf(optQBOJobHeirarchy(0).Value, "Simple", "Heirarchical")
            
 
        Case 3 'quickbooks
            HFApp.Options.Value(QuickBooksDataFile) = txtQuickBooksFile.Text
            If cboQuickBooksVersion.Text = "" Then cboQuickBooksVersion.Text = "US"
            HFApp.Options.ValueByName("AccountingVersion") = cboQuickBooksVersion.Text
            HFApp.Options.ValueByName("QuickBooksJobStyle") = IIf(optQBJobHeirarchy(0).Value, "Simple", "Heirarchical")
            HFApp.Options.ValueByName("UseQBARByJob") = chkQBPostRevenueToJob.Value = vbChecked
            HFApp.Options.ValueByName("PostAssembliesAsSalesInvoices") = chkPostAssembliesAsSalesInvoices.Value = vbChecked
                
                
        Case 4 'simply accounting
            HFApp.Options.Value(SimplyDataFile) = txtSimplyFile.Text
            HFApp.Options.ValueByName("SimplyUID") = txtSimplyUser.Text
            HFApp.Options.ValueByName("SimplyPWD") = txtSimplyPswd.Text
            HFApp.Options.ValueByName("SimplyInternalCustomer") = GetComboBoxListKey(cboSimplyInternalCustomer)
            
        Case 7 'Xero
            HFApp.Options.ValueByName("XeroJobName") = "" 'txtXeroJobName.Text
            HFApp.Options.ValueByName("XeroCostCodeName") = "" 'txtXeroCostCodeName.Text
                
        Case 9 'Sage Intacct
            HFApp.Options.ValueByName("IntacctCompanyID") = txtIntacctCompanyID.Text
            HFApp.Options.ValueByName("IntacctUID") = txtIntacctUID.Text
            HFApp.Options.ValueByName("IntacctPWD") = txtIntacctPWD.Text
            HFApp.Options.ValueByName("IntacctEntity") = cboIntacctEntity.Text
            HFApp.Options.ValueByName("IntacctPOType") = cboIntacctPOType.Text
            HFApp.Options.ValueByName("IntacctSubContractType") = cboIntacctSubContractType.Text
            HFApp.Options.ValueByName("IntacctSubContractCOType") = cboIntacctSubContractCOType.Text
            HFApp.Options.ValueByName("IntacctInvoiceType") = cboIntacctInvoiceType.Text
            HFApp.Options.ValueByName("IntacctEstimateType") = cboIntacctEstimateType.Text
            HFApp.Options.ValueByName("IntacctPOCOType") = cboIntacctPOCOType.Text
            HFApp.Options.ValueByName("IntacctOneTimeItemID") = cboIntacctItem(0).Text
            HFApp.Options.ValueByName("IntacctOneTimeItemUOM") = mIntacctOneTimeItemUOM
            HFApp.Options.ValueByName("IntacctGstItemID") = cboIntacctItem(1).Text
            HFApp.Options.ValueByName("IntacctGstItemUOM") = mIntacctGstItemUOM
            HFApp.Options.ValueByName("IntacctEstimateGLBudgetID") = txtIntacctBudgetGL.Text
            HFApp.Options.ValueByName("IntacctPOReferenceFld") = cboPOReferenceFld.Text
            HFApp.Options.ValueByName("IntacctInvReferenceFld") = cboInvReferenceFld.Text

    End Select


    HFApp.Options.ValueByName("CategoriesRequired") = True

End Sub

Private Sub cboAccountingSystem_Click()
    Dim i As Long
    For i = 1 To AccountingFrame.UBound
        AccountingFrame(i).Visible = AccountingFrame(i).Caption = cboAccountingSystem.Text
    Next
    chkPostSummarizedBudgets.Enabled = cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex) <> asQuickBooks
    If cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex) = asQuickBooks Then chkPostSummarizedBudgets.Value = vbChecked
    
    lblEditList(9).Visible = cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex) = 0
    chkPostAssembliesAsSalesInvoices.Visible = cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex) = asQuickBooks
    
    lblPONumberSample.Caption = SamplePONumber
    
    chkPostJCExtraAsSubJob.Visible = cboAccountingSystem.ItemData(cboAccountingSystem.ListIndex) = asIntacct

End Sub

Private Sub cboSalesSystem_Click()
    Dim i As Long
    For i = 0 To SalesFrame.UBound
        SalesFrame(i).Visible = i = cboSalesSystem.ListIndex
    Next
End Sub

Private Sub ReadCustomProperties()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim ct As Recordset
    Dim FieldName As String
    
    s = ""
    Set rs = HFApp.SqlExec("SELECT DISTINCT Category FROM JobCustomFieldDefs")
    While Not rs.EOF
        If "" & rs(0) <> "" Then s = s & "|" & rs(0)
        rs.MoveNext
    Wend
    gCustomFields.ColComboList(0) = s
    
    
    Set ct = HFApp.SqlExec("SELECT * FROM JobCustomFieldDefs ORDER BY Category,Name")
    Set rs = HFApp.SqlExec("SELECT * FROM JobCustomFields WHERE 1=2")
    
    With gCustomFields
        .Rows = 1
        While Not ct.EOF
            FieldName = "" & ct("Name")
            i = .Rows
            .AddItem ""
            .Cell(flexcpText, i, .ColIndex("Category")) = "" & ct("Category")
            .Cell(flexcpText, i, .ColIndex("FieldName")) = FieldName
            If ct("PickList") <> "" Then
                .Cell(flexcpText, i, .ColIndex("DataType")) = "Selection List"
            Else
                .Cell(flexcpText, i, .ColIndex("DataType")) = ColumnDefinitionPretty(rs.fields(FieldName))
            End If
            
            
            ct.MoveNext
        Wend
        Call .AutoSize(0, .Cols - 1)
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "ReadCustomProperties")
End Sub
        
Private Sub WriteSupplyChain()
With HFApp.Options

    
    
    .ValueByName("DefaultPurchasingManager") = GetComboBoxListKey(cboPurchasingManager)
    .ValueByName("FieldPOPrefix") = txtFieldPOPrefix.Text

    .ValueByName("Detail_Description") = IIf(chkFieldPOField(0).Value = vbChecked, "1", "0")
    .ValueByName("Detail_Quantity") = IIf(chkFieldPOField(1).Value = vbChecked, "1", "0")
    .ValueByName("Detail_Rate") = IIf(chkFieldPOField(2).Value = vbChecked, "1", "0")
    .ValueByName("Detail_Job") = IIf(chkFieldPOField(3).Value = vbChecked, "1", "0")
    .ValueByName("Detail_Extra") = IIf(chkFieldPOField(4).Value = vbChecked, "1", "0")
    .ValueByName("Detail_CostCode") = IIf(chkFieldPOField(5).Value = vbChecked, "1", "0")
    .ValueByName("Detail_Category") = IIf(chkFieldPOField(6).Value = vbChecked, "1", "0")
    .ValueByName("Detail_TaxGroup") = IIf(chkFieldPOField(7).Value = vbChecked, "1", "0")
    
    .ValueByName("ViewPurchasingManagerDetailsForm") = IIf(chkUseDetailedFieldPOView.Value = vbChecked, "1", "0")
    
    .ValueByName("WebAttachmentsFolder") = txtWebAttachmentsFolder.Text

End With
End Sub
Private Function SaveData() As Boolean
    Dim i As Long
    Dim s As String
    Dim c As New adodb.Connection
    Screen.MousePointer = vbHourglass
    
    
    Call WriteRoleDefinitions
    Call WriteBuildProOptions
    Call WriteDepartmentApprovals
    Call WriteDocumentManagement
    Call WriteSupplyChain
    
    With HFApp.Options
        .ValueByName("BuilderType") = cboBuilderType.Text
        .ValueByName("UseAltCostCodesForCO") = optUseAltCostCodesForCO(1).Value
        .ValueByName("DefaultVarianceCategory") = GetComboBoxListKey(cboVarianceCat)
        .ValueByName("UseVarianceReporting") = chkUseVarianceReporting.Value = vbChecked
        .ValueByName("PurchasingRequiresSalesApproval") = chkPurchasingRequiresSalesApproval.Value = vbChecked
        
        
     
        .Value(CompanyName) = txtCompanyName.Text
        .Value(ContactPerson) = txtContactPerson.Text
        .Value(address1) = txtCompanyAddr1.Text
        .Value(address2) = txtCompanyAddr2.Text
        .Value(city) = txtCompanyCity.Text
        .Value(province) = cboCompanyProvince.Text
        .Value(zip) = txtCompanyPostalCode.Text
        .Value(Country) = cboCompanyCountry.Text
        .ValueByName("EnforcePasswordComplexity") = IIf(chkEnforcePasswordComplexity.Value = vbChecked, "True", "False")
        
        .Value(phone) = txtCompanyPhone.Text
        .Value(fax) = txtCompanyFax.Text
        .Value(Email) = txtCompanyEmail.Text
        
        .Value(TaxGroupRequired) = chkUseTaxGroups.Value = vbChecked
        .Value(TaxGroupLabour) = cboTaxGroup(0).Text
        .Value(TaxGroupMaterial) = cboTaxGroup(1).Text
        .Value(TaxGroupSubContract) = cboTaxGroup(2).Text
        .Value(TaxGroupEquipment) = cboTaxGroup(3).Text
        .Value(TaxGroupOverhead) = cboTaxGroup(4).Text
        .Value(TaxGroupOther) = cboTaxGroup(5).Text
        .Value(UseEstimatingWorksheetsOnly) = chkEstimatingWorksheetsOnly.Value = vbChecked
        .Value(IncentivePhase) = Parse(txtIncentiveItem.tag, 1, Chr(1))
        .Value(IncentiveItem) = Parse(txtIncentiveItem.tag, 2, Chr(1))
        .Value(IncentiveItemDescription) = Parse(txtIncentiveItem.tag, 3, Chr(1))
        .ValueByName("LetPurchaserChangeSaleQty") = chkLetPurchaserChangeSaleQty.Value = vbChecked
        .ValueByName("UseComponents") = chkUseComponents.Value = vbChecked
        
        .Value(SelectAtTakeoffPhase) = Parse(txtSelectAtTakeoffItem.tag, 1, Chr(1))
        .Value(SelectAtTakeoffItem) = Parse(txtSelectAtTakeoffItem.tag, 2, Chr(1))
        .Value(SelectAtTakeoffDescription) = Parse(txtSelectAtTakeoffItem.tag, 3, Chr(1))
        
        .Value(IncentiveCostPercent) = Val(txtIncentiveCostPercent.Text)
        .Value(LandPhase) = Parse(txtLandItem.tag, 1, Chr(1))
        .Value(LandItem) = Parse(txtLandItem.tag, 2, Chr(1))
        .Value(LandItemDescription) = Parse(txtLandItem.tag, 3, Chr(1))
        
        Select Case True
        Case optPOByJob.Value:          .ValueByName("PONumberingStyle") = "Job"
        Case optPOByJobPOIndex.Value:   .ValueByName("PONumberingStyle") = "JobPOIndex"
        Case Else:                      .ValueByName("PONumberingStyle") = "Sequential"
        End Select
        .ValueByName("POFormatSegment1") = Val("" & txtPOSegment(1).Text)
        .ValueByName("POFormatSegment2") = "" & txtPOSegment(2).Text
        .ValueByName("POFormatSegment3") = Val("" & txtPOSegment(3).Text)
        .ValueByName("POFormatSegment4") = Val("" & txtPOSegment(4).Text)
        .ValueByName("POFormatSegment5") = "" & txtPOSegment(5).Text
        .ValueByName("POFormatSegment6") = Val("" & txtPOSegment(6).Text)
        .ValueByName("POFormatSegment7") = "" & txtPOSegment(7).Text
        .ValueByName("POFormatSegment8") = Val("" & txtPOSegment(8).Text)
        
        .Value(AdjustSellPriceWithCosts) = chkAdjustSelling.Value = vbChecked
        .ValueByName("AdjustIncentiveRetailWithCosts") = chkAdjustIncentive.Value = vbChecked
        
        .Value(LockPostedSalesWorksheets) = chkLockPostedSheets.Value = vbChecked
        .Value(AllowForecastsWhenRefreshingPOs) = chkAllowForecast.Value = vbChecked
        
        .Value(MarketingWorkSheetAutoRefreshCosts) = chkMarketingWorkSheetAutoRefreshCosts.Value = vbChecked
        
        .Value(includetax) = chkSellingPriceIncludesTax.Value = vbChecked
        .Value(ZeroRateOnRefreshCosts) = chkZeroRateOnRefeshCosts.Value = vbChecked
        .Value(ZeroRateOnChangeVendor) = chkZeroRateOnChangeVendor.Value = vbChecked
        .Value(CanAddContractItemsFromPurchasing) = chkCanAddContractItemsFromPurchasing.Value = vbChecked
        .Value(UsePricingFromEstimating) = chkUsePricingFromEstimating.Value = vbUnchecked
        .Value(POPrefix) = txtPOPrefix.Text
        
        'i = 0
        'If optZeroQtyTakeoffMode(1).Value Then i = 1
        'If optZeroQtyTakeoffMode(2).Value Then i = 2
        '.Value(ZeroQtyTakeoffs) = i
        .Value(ZeroQtyTakeoffs) = cboZeroQtyTakeoffMode.ListIndex
        
        .Value(Msg_SendPO_Subject) = txtNotification(0).Text
        .Value(Msg_SendPO_Body) = txtNotification(1).Text
        .Value(Msg_CancelPO_Subject) = txtNotification(2).Text
        .Value(Msg_CancelPO_Body) = txtNotification(3).Text
        
        
    
    

        
    End With
    Call HFApp.SqlExec("Update system_setup set MaxVendorPricing=" & DbQuote(Bit, Me.chkUseMaxVendorPricing.Value) & " where id=" & HFApp.DivisionID)
    
    Call SaveEmailConfig
    Call SaveAccountingSystem
    Call SaveAccountingImportOptions
    Call SaveSalesSystem
    Call SaveEstimatingOptions
    Call SaveCommunityStandards
    
    
    
    
    With gFormatting
        HFApp.Options.Value(Format_QtyRateEQZero_FontStyle) = .Cell(flexcpText, 0, 1)
        HFApp.Options.Value(Format_QtyRateLTZero_FontStyle) = .Cell(flexcpText, 1, 1)
        HFApp.Options.Value(Format_NonSysRate_FontStyle) = .Cell(flexcpText, 2, 1)
        HFApp.Options.Value(Format_InvalidData_FontStyle) = .Cell(flexcpText, 3, 1)
        
        HFApp.Options.Value(Format_QtyRateEQZero_BackColor) = .Cell(flexcpBackColor, 0, 2)
        HFApp.Options.Value(Format_QtyRateLTZero_BackColor) = .Cell(flexcpBackColor, 1, 2)
        HFApp.Options.Value(Format_NonSysRate_BackColor) = .Cell(flexcpBackColor, 2, 2)
        HFApp.Options.Value(Format_InvalidData_BackColor) = .Cell(flexcpBackColor, 3, 2)

        HFApp.Options.Value(Format_QtyRateEQZero_ForeColor) = .Cell(flexcpBackColor, 0, 3)
        HFApp.Options.Value(Format_QtyRateLTZero_ForeColor) = .Cell(flexcpBackColor, 1, 3)
        HFApp.Options.Value(Format_NonSysRate_ForeColor) = .Cell(flexcpBackColor, 2, 3)
        HFApp.Options.Value(Format_InvalidData_ForeColor) = .Cell(flexcpBackColor, 3, 3)
    End With

    Call HFApp.Options.SaveData



    Call SaveDivisions
    'ensure change to system_setup address is also saved to division
    s = ""
    s = s & "update d set" & vbCrLf
    s = s & " address1= s.address1" & vbCrLf
    s = s & ",address2= s.address2" & vbCrLf
    s = s & ",city = s.city" & vbCrLf
    s = s & ",province = s.province" & vbCrLf
    s = s & ",country = s.country" & vbCrLf
    s = s & ",postal = s.zip" & vbCrLf
    s = s & ",phone = s.phone" & vbCrLf
    s = s & ",fax = s.fax" & vbCrLf
    s = s & "from system_setup s" & vbCrLf
    s = s & "join divisions d on s.id=d.divisionid" & vbCrLf
    s = s & "where s.id=" & HFApp.DivisionID
    Call HFApp.SqlExec(s)
    
    SaveData = True

    Screen.MousePointer = vbDefault
End Function

Private Function UnMask(s As String) As String
    Dim n As String
    Dim i As Long
    For i = 1 To Len(s)
        If IsNumeric(Mid(s, i, 1)) Then n = n & Mid(s, i, 1)
    Next
    UnMask = n
End Function


Private Sub optQBJobHeirarchy_Click(Index As Integer)
    chkQBPostRevenueToJob.Enabled = optQBJobHeirarchy(0).Value
End Sub







Private Sub tabNotices_Click()
    Dim i As Long
    For i = 0 To 1
        frmNotices(i).Visible = tabNotices.SelectedItem.Index = i + 1
    Next
End Sub


Private Sub txtCompanyFax_Validate(Cancel As Boolean)
    txtCompanyFax.Text = FormatPhone(txtCompanyFax.Text, Me.cboCompanyCountry.Text)
End Sub

Private Sub txtCompanyPhone_Validate(Cancel As Boolean)
    txtCompanyPhone.Text = FormatPhone(txtCompanyPhone.Text, cboCompanyCountry.Text)
End Sub



Private Sub txtCrmApiKey_GotFocus()
    SelectAll txtCrmApiKey
End Sub


Private Sub txtCrmClientID_GotFocus()
    SelectAll txtCrmClientID
End Sub

Private Sub txtEmailAddress_GotFocus()
    SelectAll txtEmailAddress
End Sub

Private Sub txtIncentiveCostPercent_Validate(Cancel As Boolean)
    txtIncentiveCostPercent = Val(txtIncentiveCostPercent.Text) & "%"
End Sub

Private Sub txtIncentiveItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbAltMask
        Case KeyCode = vbKeyDelete
            txtIncentiveItem.tag = ""
            txtIncentiveItem.Text = ""
        Case Else
            Call cmdChooseItem_Click(0)
    End Select
End Sub
Private Sub txtLandItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbAltMask
        Case KeyCode = vbKeyDelete
            txtLandItem.tag = ""
            txtLandItem.Text = ""
        Case Else
            Call cmdChooseItem_Click(1)
    End Select
End Sub

Private Function SamplePONumber() As String

    Dim i1 As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    
    Dim i4 As Integer
    Dim i5 As Integer
    Dim i6 As Integer
    Dim i7 As Integer
    Dim i8 As Integer
    
    Dim Job As String
    Dim POIndex As String
    Dim PO  As String
    
    Dim POLength As Integer
    POLength = MaxPOLength
    
    
    Select Case True
        Case optPOByJob.Value
            i2 = Min(POLength, Max(0, Len(txtPOSegment(2).Text)))
            i3 = Min(POLength - i2, Max(0, Val(txtPOSegment(3).Text)))
            i1 = Min(POLength, Max(0, POLength - i3 - i2))
            txtPOSegment(1).Text = i1
            txtPOSegment(3).Text = i3
    
            Job = "AAAAAAAAAAAAAAAAAAAAAA"
            PO = "9999999999999999999999"
    
            SamplePONumber = Right(Job, i1) & txtPOSegment(2).Text & left(Format(PO, String(i3, "0")), i3)
        
        Case optPOByJobPOIndex.Value
            i5 = Min(POLength, Max(0, Len(txtPOSegment(5).Text)))
            i7 = Min(POLength, Max(0, Len(txtPOSegment(7).Text)))
            i6 = Min(POLength - i5 - i7, Max(0, Val(txtPOSegment(6).Text)))
            i8 = Min(POLength - i5 - i7 - i6, Max(0, Val(txtPOSegment(8).Text)))
            i4 = Min(POLength, Max(0, POLength - i8 - i7 - i6 - i5))
            txtPOSegment(4).Text = i4
            txtPOSegment(6).Text = i6
            txtPOSegment(8).Text = i8
            
            Job = "AAAAAAAAAAAAAAAAAAAAAA"
            POIndex = "BBBBBBBBBBBBBBBBBBBBBB"
            PO = "9999999999999999999999"
    
            SamplePONumber = left(Job, i4) & txtPOSegment(5).Text & Right(POIndex, i6) & txtPOSegment(7).Text & left(PO, i8)
        
        Case Else
            SamplePONumber = Mid(Me.txtPOPrefix.Text + Me.txtNextPO.Text, 1, POLength)
            
    End Select
    
End Function



Private Sub txtNextPO_Validate(Cancel As Boolean)
    txtNextPO.Text = Val(txtNextPO.Text)
End Sub

Private Sub txtNotification_DblClick(Index As Integer)
    Dim s As String
    s = txtNotification(Index).Text
    If FComments.Edit(s, txtNotification(Index), , "") Then txtNotification(Index).Text = s
End Sub


Private Sub txtNotification_GotFocus(Index As Integer)
    SelectAll txtNotification(Index)
End Sub



Private Sub txtPOPrefix_Change()
    lblPONumberSample.Caption = SamplePONumber
End Sub

Private Sub txtPOPrefix_GotFocus()
    SelectAll txtPOPrefix
End Sub

Private Sub txtPOSegment_Change(Index As Integer)
    If Not IsIn(Index, 2, 5, 7) Then
        txtPOSegment(Index).Text = Val(txtPOSegment(Index).Text)
    End If
    lblPONumberSample.Caption = SamplePONumber
    SelectAll txtPOSegment(Index)
End Sub

Private Sub gTOC_RowColChange()
    Dim i As Integer
    
    With gTOC
        .Cell(flexcpFontBold, 0, 1, .Rows - 1, 1) = False
        .Cell(flexcpFontBold, 0, 1) = True
        .Cell(flexcpFontBold, .Row, 1) = True
        
        For i = 0 To TabFrame.Count - 1
            TabFrame(i).ZOrder 0
            TabFrame(i).Enabled = TabFrame(i).Caption = .TextMatrix(.Row, 1)
            TabFrame(i).Visible = TabFrame(i).Enabled
        Next
        
    End With

End Sub


Private Sub cmdNextPO_Click()
On Error GoTo eh
    Dim nextPO As String
    Dim s As String
    Dim rs As Recordset
    
    nextPO = InputBox(vbCrLf & vbCrLf & vbCrLf & "Enter the Next PO Number", "")
    If nextPO = "" Then Exit Sub
    If IsNumeric(nextPO) = False Then
        MsgBox "Invalid entry. Next PO Number must be an integer value", vbExclamation, App.ProductName
        Exit Sub
    Else
        If nextPO - Int(Val(nextPO)) <> 0 Then
            MsgBox "Invalid entry. Next PO Number must be an integer value", vbExclamation, App.ProductName
            Exit Sub
        Else
            Set rs = HFApp.SqlExec("select ponumber from pomaster where isnumeric(ponumber)=1 and cast(ponumber as int)>=" & DbQuote(Num, nextPO) & " order by cast(ponumber as int) desc")
            If Not rs.EOF Then
                MsgBox "Invalid entry. HomeFront already has PO's", vbExclamation, App.ProductName
                Exit Sub
            End If
        End If
    End If
    
    
    If nextPO <> 0 Then
        Call HFApp.SqlExec("DELETE FROM PONumbers", dbHomefront)
        Call HFApp.SqlExec("DBCC CHECKIDENT(PONumbers, RESEED, " & nextPO - 1 & ")", dbHomefront)
        Call HFApp.SqlExec("INSERT INTO PONumbers(TStmp) VALUES(getdate())", dbHomefront)
        txtNextPO.Text = nextPO
    End If
    
    lblPONumberSample.Caption = SamplePONumber
    
Exit Sub
eh: MsgBox Parse(Err.Description, Parse(Err.Description, , "]"), "]"), vbExclamation, App.ProductName
End Sub

Private Sub SaveEstimatingOptions()
    Dim i As Long
    Dim s As String
    
    
    HFApp.Options.ValueByName("EstimatingSystem") = cboEstimatingSystem.ListIndex
    
    
    'CGVisions
    HFApp.Options.ValueByName("CGVisionsRootURI") = txtCGVisionsRootURI.Text
    HFApp.Options.ValueByName("CGVisionsUID") = txtCGVisionsUID.Text
    HFApp.Options.ValueByName("CGVisionsPWD") = txtCGVisionsPWD.Text
    HFApp.Options.ValueByName("BIMBaseModelNumber") = txtBIMBaseModelNumber.Text
    HFApp.Options.ValueByName("BIMCommunitySpecificPlansAndOptions") = optPipelineCommunityStyle(1).Value

    
    
    'New Sage Estimating
    HFApp.Options.ValueByName("SageSqlEstServer") = txtSageSqlEstServer.Text
    HFApp.Options.ValueByName("SageSqlEstDatabase") = cboSageSqlEstDatabase.Text
    HFApp.Options.ValueByName("SageSqlEstUser") = txtSageSqlEstUser.Text
    HFApp.Options.ValueByName("SageSqlEstPswd") = txtSageSqlEstPswd.Text
    HFApp.Options.ValueByName("SageSqlEstConsolidateItems") = chkSageSqlEstConsolidateItems.Value = vbChecked
    
    
    
    HFApp.Options.ValueByName("TakeoffSystem") = Choose(cboTOSystem.ListIndex + 1, "none", "On-Screen", "PlanSwift")
    
    HFApp.Options.ValueByName("OnScreenType") = IIf(optOnScreenType(1).Value, "SQL", "MDB")
    HFApp.Options.ValueByName("OnScreenMDB") = txtOnScreenMDB.Text
    HFApp.Options.ValueByName("OnScreenServer") = txtOnScreenServer.Text
    HFApp.Options.ValueByName("OnScreenDatabase") = txtOnScreenDatabase.Text
    HFApp.Options.ValueByName("OnScreenUID") = txtOnScreenUID.Text
    HFApp.Options.ValueByName("OnScreenPWD") = txtOnScreenPWD.Text
    
End Sub

Private Sub ReadDivisions()
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    With gDivisions
        .Rows = 1
        s = "select * from divisions order by divisioncode"
        Set rs = HFApp.SqlExec(s, dbHomefront)
        While Not rs.EOF
            Call .AddItem("")
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = "" & rs("DivisionID")
            .TextMatrix(.Rows - 1, .ColIndex("Code")) = "" & rs("DivisionCode")
            .TextMatrix(.Rows - 1, .ColIndex("Company")) = "" & rs("DivisionName")
            .TextMatrix(.Rows - 1, .ColIndex("Address1")) = "" & rs("Address1")
            .TextMatrix(.Rows - 1, .ColIndex("Address2")) = "" & rs("Address2")
            .TextMatrix(.Rows - 1, .ColIndex("City")) = "" & rs("City")
            .TextMatrix(.Rows - 1, .ColIndex("Province")) = "" & rs("Province")
            .TextMatrix(.Rows - 1, .ColIndex("Postal")) = "" & rs("Postal")
            .TextMatrix(.Rows - 1, .ColIndex("Country")) = "" & rs("Country")
            .TextMatrix(.Rows - 1, .ColIndex("County")) = "" & rs("County")
            .TextMatrix(.Rows - 1, .ColIndex("TaxNumber")) = "" & rs("TaxNumber")
            .TextMatrix(.Rows - 1, .ColIndex("Phone")) = "" & rs("Phone")
            .TextMatrix(.Rows - 1, .ColIndex("Fax")) = "" & rs("Fax")
            .TextMatrix(.Rows - 1, .ColIndex("BalSheetPrefix")) = "" & rs("BalSheetPrefix")
            .TextMatrix(.Rows - 1, .ColIndex("IncomeStmtPrefix")) = "" & rs("IncomeStmtPrefix")
            rs.MoveNext
        Wend
        .AddItem ""
        Call .AutoSize(0, 1, 2)
    End With
End Sub

Private Sub SaveDivisions()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    With gDivisions
'        For i = .Rows - 2 To 1 Step -1
'            If .RowHidden(i) Then
'                s = "delete from divisions where divisionid=" & DbQuote(Num, .TextMatrix(i, .ColIndex("ID")))
'                Call HFApp.SqlExec(s, dbHomefront)
'                s = "delete from divisioncommunities where divisionid=" & DbQuote(Num, .TextMatrix(i, .ColIndex("ID")))
'                Call HFApp.SqlExec(s, dbHomefront)
'                Call .RemoveItem(i)
'            End If
'        Next
        For i = .Rows - 2 To 1 Step -1
            If .ValueMatrix(i, .ColIndex("ID")) = 0 Then
                
                s = "insert into divisions(divisioncode,divisionname) values(" & DbQuote(Str, .TextMatrix(i, .ColIndex("code"))) & "," & DbQuote(Str, .TextMatrix(i, .ColIndex("company"))) & ")"
                Call HFApp.SqlExec(s, dbHomefront)
                .TextMatrix(i, .ColIndex("ID")) = HFApp.SqlIdentity("divisions", dbHomefront)
                
                s = ""
                s = s & "insert into System_Setup(ID, Export_Path, Import_Path, Def_Rep_Path, Cost_Code_Section1, Cost_Code_Section2, Cost_Code_Section3, Cost_Code_Section4, Cost_Code_Separator, Cost_Code_Mask, Job_Section1, Job_Section2, Job_Section3, Job_Separator, Job_Mask, GL_AR_Account, GL_Cost_Dep_Account, Timberline_UID, Timberline_PWD, Deposit_DSN, Deposit_Debit, Deposit_credit, Mortgage_Debit, Mortgage_Credit, Sale_Debit, Sale_credit, Adjustment_Credit, GST_Credit, GST_Rebate_Credit, Holdback_Debit, Holdback_Credit, Draw, Tax_Group, tax_id, USER_DSN, Comm_Cost_code, Comm_Category, Next_Batch, Next_Dep_Seq, Next_Mort_Seq, Next_Comm_Seq, Next_Adjust_Seq, Next_Mort_Adv_Id, Tim_Cost_Table, Tim_Category_Table, Estimator_Email, Contract_Document, Cust_Section1, Cust_Section2, Cust_Section3, Cust_Separator, Customer_Section1, Customer_Section2, Customer_Section3, Customer_Separator, Customer_Mask, Database_Version, Post_Deposit_Macro" & vbCrLf
                s = s & ", Post_Deposit_Path, Post_Mortgage_Macro, Post_Mortgage_Path, Post_Comm_Macro, Post_Comm_Path" & vbCrLf
                s = s & ", Post_Adj_Macro, Post_Adj_Path, Process_Sales_Macro, Process_Sales_Path, Backup_Folder, Use_Timberline, GST_Rate, GST_Rebate_Rate, Document_Folder, Acct_Receivable, Days_Cutoff_Notice, Days_Deposit_Notice, Tim_Extra_Table, TL_AR_Data_Path, Show_CDN_GST, Sales_Override, Timberline_Data_Path, Lot_Inv_debit, Lot_Inv_credit, Req_Prefix_Job, Land_Cost_Code, Land_Category, GL_Lot_Inv_Account, GL_Lot_Pay_Account, Lot_Inventory_Macro, Lot_Inventory_Path, Use_Addition_Changes, Def_Lending_Company" & vbCrLf
                s = s & ", Def_Law_Firm, GL_Prefix, Next_Lot_No, Lot_LIAB_GL, Working_Progress, Def_Template_Path, ExportExcel_Path, ModelByArea, UseMajorGroup, UseElevation, Next_Job_Incr, Cardel, UsePst, pst_rate, PST_Province, UseSelection, UseSeqCustomer, Next_Customer_No, Comm_Type, DefaultFee, fiscal_year, UseSpecWordDoc, Def_Spec_Doc, UseCM, Display_CatDesc, CompanyName, CompanyLogo, Address1, Address2, City, Province, Zip, Phone, FAX, Email, ReqCategory, Default_Category, MultiFamily, TLJCACCTFORMAT, UseServiceModule, ServiceModuleUDL" & vbCrLf
                s = s & ", Use_french, Use_Prefix, ShowMajorGroupDesc, CreateJobLot, ApplyLotCostJob, Community_Cost_Code, Taxes_Cost_Code, Interest_Cost_Code, Misc_Cost_Code, GlobalJobIncr, TLMunicipal_Add, ContractDocInWord, ModelSpecDocInWord, DefaultCity, DefaultProvince, UseINCPrefix_Adj" & vbCrLf
                s = s & ", accounting_approve, allow_customprice, french_contract, GST_No, DisableMiscApproved, dep_ar_cr_inv, Cut_Off_Table_Name, Cut_Off_Field_Name, Monitor_DT_TL, Monitor_DT_TL_Field, UseUserLog, UseHFWeb, HFWebUDL, EnableSelectionWeb, Interior_Viewable, EnableBankingLink, COPrefix, ACPrefix, DCPrefix, USE_POUPLOADWEB, POUPLOAD_FROM, POUPLOAD_DSN, SchRevenue, CORevenue, PSTAcct, PSTRebateAcct, Job1Incr, Job2Incr, Job3Incr, Job1Table, Job2Table, Job3Table, Job1Field, Job2Field, Job3Field, Job1Ornt, Job2Ornt, Job3Ornt, AreaSpecSalesPerson, CostCodeSaleClosing, UpdateDefLotStatus, DefLotStatus, POUPLOAD_HEADER, POUPLOAD_DETAILS, UsePO, MonitorUserLog, LogFolder, HMSUID, HMSPWD, USESCHEDULE, POUPLOAD_FIELD, IncludeTax, NETTAXRATE, ModelTaxRate, LotTaxRate" & vbCrLf
                s = s & ", MSOPTTaxRate, GLOPTTaxRate, DCOPTTaxRate, CustomOptionsTaxRate, RevenueAdjGLAcct, FileAttachmentBase, AlternatingGridColor, DateFormat, PhoneMask, ApproveCOReqQuote, NoImage, taxadjust_to_Bldr, BalanceSheetMaxAcct" & vbCrLf
                s = s & ", showMsgOnEachLoad, ShowSpecDocument, OptionByArea, ModelsbyArea_Phase, OptionByAreaPhase, CopyACCuttoff, TLDateFormat, AdjustJCExtra, SalesJCCategory, UsePricingFromEstimating, SendFaxUsingPurchasersInfo, GlobalOptionByArea, GlobalOptionByAreaPhase, DCOptionByArea, DCOptionByAreaPhase, DCRevenue, ReservationDays, RescindDays, ExemptwithGST, AutoAssignARCust, MultiUnitsasExtras, LogApptinOutlook, BuilderCode, WebAdminUser, WebUploadPath, WebUploadUser, WebUploadPswd, Country, ContactPerson, BuilderName, RemoveDetailsonSpecConvert, PromptForSpecOptions, COEntryPriortoPurch, ShowAssignCustOnly, UseEquipmentCostType, TLAccountLevel, DisableEmailEstimator, PurchasingRequiresSalesApproval, MaxVendorPricing, UseOutlookforEmail, ApptDuration, ApptDefaultText" & vbCrLf
                s = s & ", LotOnlyJobOverride)" & vbCrLf
                s = s & "select " & .TextMatrix(i, .ColIndex("ID")) & ", Export_Path, Import_Path, Def_Rep_Path, Cost_Code_Section1, Cost_Code_Section2, Cost_Code_Section3, Cost_Code_Section4, Cost_Code_Separator, Cost_Code_Mask, Job_Section1, Job_Section2, Job_Section3, Job_Separator, Job_Mask, GL_AR_Account, GL_Cost_Dep_Account, Timberline_UID, Timberline_PWD, Deposit_DSN, Deposit_Debit, Deposit_credit, Mortgage_Debit, Mortgage_Credit, Sale_Debit, Sale_credit, Adjustment_Credit, GST_Credit, GST_Rebate_Credit, Holdback_Debit, Holdback_Credit, Draw, Tax_Group, tax_id, USER_DSN, Comm_Cost_code, Comm_Category, Next_Batch, Next_Dep_Seq, Next_Mort_Seq, Next_Comm_Seq, Next_Adjust_Seq, Next_Mort_Adv_Id, Tim_Cost_Table, Tim_Category_Table, Estimator_Email, Contract_Document, Cust_Section1, Cust_Section2, Cust_Section3, Cust_Separator, Customer_Section1, Customer_Section2, Customer_Section3, Customer_Separator, Customer_Mask, Database_Version, Post_Deposit_Macro" & vbCrLf
                s = s & ", Post_Deposit_Path, Post_Mortgage_Macro, Post_Mortgage_Path, Post_Comm_Macro, Post_Comm_Path" & vbCrLf
                s = s & ", Post_Adj_Macro, Post_Adj_Path, Process_Sales_Macro, Process_Sales_Path, Backup_Folder, Use_Timberline, GST_Rate, GST_Rebate_Rate, Document_Folder, Acct_Receivable, Days_Cutoff_Notice, Days_Deposit_Notice, Tim_Extra_Table, TL_AR_Data_Path, Show_CDN_GST, Sales_Override, Timberline_Data_Path, Lot_Inv_debit, Lot_Inv_credit, Req_Prefix_Job, Land_Cost_Code, Land_Category, GL_Lot_Inv_Account, GL_Lot_Pay_Account, Lot_Inventory_Macro, Lot_Inventory_Path, Use_Addition_Changes, Def_Lending_Company" & vbCrLf
                s = s & ", Def_Law_Firm, GL_Prefix, Next_Lot_No, Lot_LIAB_GL, Working_Progress, Def_Template_Path, ExportExcel_Path, ModelByArea, UseMajorGroup, UseElevation, Next_Job_Incr, Cardel, UsePst, pst_rate, PST_Province, UseSelection, UseSeqCustomer, Next_Customer_No, Comm_Type, DefaultFee, fiscal_year, UseSpecWordDoc, Def_Spec_Doc, UseCM, Display_CatDesc, CompanyName, CompanyLogo, Address1, Address2, City, Province, Zip, Phone, FAX, Email, ReqCategory, Default_Category, MultiFamily, TLJCACCTFORMAT, UseServiceModule, ServiceModuleUDL" & vbCrLf
                s = s & ", Use_french, Use_Prefix, ShowMajorGroupDesc, CreateJobLot, ApplyLotCostJob, Community_Cost_Code, Taxes_Cost_Code, Interest_Cost_Code, Misc_Cost_Code, GlobalJobIncr, TLMunicipal_Add, ContractDocInWord, ModelSpecDocInWord, DefaultCity, DefaultProvince, UseINCPrefix_Adj" & vbCrLf
                s = s & ", accounting_approve, allow_customprice, french_contract, GST_No, DisableMiscApproved, dep_ar_cr_inv, Cut_Off_Table_Name, Cut_Off_Field_Name, Monitor_DT_TL, Monitor_DT_TL_Field, UseUserLog, UseHFWeb, HFWebUDL, EnableSelectionWeb, Interior_Viewable, EnableBankingLink, COPrefix, ACPrefix, DCPrefix, USE_POUPLOADWEB, POUPLOAD_FROM, POUPLOAD_DSN, SchRevenue, CORevenue, PSTAcct, PSTRebateAcct, Job1Incr, Job2Incr, Job3Incr, Job1Table, Job2Table, Job3Table, Job1Field, Job2Field, Job3Field, Job1Ornt, Job2Ornt, Job3Ornt, AreaSpecSalesPerson, CostCodeSaleClosing, UpdateDefLotStatus, DefLotStatus, POUPLOAD_HEADER, POUPLOAD_DETAILS, UsePO, MonitorUserLog, LogFolder, HMSUID, HMSPWD, USESCHEDULE, POUPLOAD_FIELD, IncludeTax, NETTAXRATE, ModelTaxRate, LotTaxRate" & vbCrLf
                s = s & ", MSOPTTaxRate, GLOPTTaxRate, DCOPTTaxRate, CustomOptionsTaxRate, RevenueAdjGLAcct, FileAttachmentBase, AlternatingGridColor, DateFormat, PhoneMask, ApproveCOReqQuote, NoImage, taxadjust_to_Bldr, BalanceSheetMaxAcct" & vbCrLf
                s = s & ", showMsgOnEachLoad, ShowSpecDocument, OptionByArea, ModelsbyArea_Phase, OptionByAreaPhase, CopyACCuttoff, TLDateFormat, AdjustJCExtra, SalesJCCategory, UsePricingFromEstimating, SendFaxUsingPurchasersInfo, GlobalOptionByArea, GlobalOptionByAreaPhase, DCOptionByArea, DCOptionByAreaPhase, DCRevenue, ReservationDays, RescindDays, ExemptwithGST, AutoAssignARCust, MultiUnitsasExtras, LogApptinOutlook, BuilderCode, WebAdminUser, WebUploadPath, WebUploadUser, WebUploadPswd, Country, ContactPerson, BuilderName, RemoveDetailsonSpecConvert, PromptForSpecOptions, COEntryPriortoPurch, ShowAssignCustOnly, UseEquipmentCostType, TLAccountLevel, DisableEmailEstimator, PurchasingRequiresSalesApproval, MaxVendorPricing, UseOutlookforEmail, ApptDuration, ApptDefaultText" & vbCrLf
                s = s & ", LotOnlyJobOverride" & vbCrLf
                s = s & "from System_Setup where id =" & HFApp.DivisionID & vbCrLf
                HFApp.SqlExec s
                
                s = ""
                s = s & "insert into ReportAccess(Divisionid, Report_Name, Description, Report_Path, User_Access, Report_Group, Job_Mask_Req, Custom_Report)" & vbCrLf
                s = s & "select " & DbQuote(Num, .TextMatrix(i, .ColIndex("ID"))) & ", Report_Name, Description, Report_Path, User_Access, Report_Group, Job_Mask_Req, Custom_Report" & vbCrLf
                s = s & "from ReportAccess" & vbCrLf
                s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront)
                
                s = ""
                s = s & "insert into AppOptions(UID, OptionName, OptionValue, UStmp, TStmp, divisionid)" & vbCrLf
                s = s & "select UID, OptionName, OptionValue, UStmp, TStmp, " & DbQuote(Num, .TextMatrix(i, .ColIndex("ID"))) & vbCrLf
                s = s & "from AppOptions" & vbCrLf
                s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront)
                
                
                If MsgBox("Would you like to Copy the Estimating Item Database, and Po Indexes from this Division to the new Division?", vbYesNo) = vbYes Then
                    
                    s = "insert into dbo.tblPhaseItem(Phase, Item, PriceLink, Description, Notes, POIndex, JCCostCode, JCCategory, OrderUOM, TakeoffUOM, UpdateEstimating, UStmp, TStmp, ConversionFactor, Price, TaxGroup, WastePercent, RoundDir, Roundto, PhaseSortOrder, ItemSortOrder, flag, PartNumber, CostCategory, ItemNumber, IsQuote, UseInFieldPO, AltJCCostCode, AltJCCategory, Formula, ModifiedDate, ModifiedBy, PriceModifiedDate, PriceModifiedBy, Location, WBS01, WBS02, WBS03, WBS04, WBS05, WBS06, WBS07, WBS08, WBS09, WBS10, WBS11, WBS12, WBS13, WBS14, WBS15, WBS16, WBS17, WBS18, WBS19, WBS20, WBS21, WBS22, WBS23, WBS24, WBS25, WBS26, WBS27, WBS28, WBS29, WBS30, WBS31, WBS32, WBS33, WBS34, WBS35, WBS36, WBS37, WBS38, WBS39, WBS40, OptionID, Color, OptionCategory, RetailPretax, DivisionID) "
                    s = s & "select Phase, Item, PriceLink, Description, Notes, POIndex, JCCostCode, JCCategory, OrderUOM, TakeoffUOM, UpdateEstimating, UStmp, TStmp, ConversionFactor, Price, TaxGroup, WastePercent, RoundDir, Roundto, PhaseSortOrder, ItemSortOrder, flag, PartNumber, CostCategory, ItemNumber, IsQuote, UseInFieldPO, AltJCCostCode, AltJCCategory, Formula, ModifiedDate, ModifiedBy, PriceModifiedDate, PriceModifiedBy, Location, WBS01, WBS02, WBS03, WBS04, WBS05, WBS06, WBS07, WBS08, WBS09, WBS10, WBS11, WBS12, WBS13, WBS14, WBS15, WBS16, WBS17, WBS18, WBS19, WBS20, WBS21, WBS22, WBS23, WBS24, WBS25, WBS26, WBS27, WBS28, WBS29, WBS30, WBS31, WBS32, WBS33, WBS34, WBS35, WBS36, WBS37, WBS38, WBS39, WBS40, OptionID, Color, OptionCategory, RetailPretax," & .TextMatrix(i, .ColIndex("ID"))
                    s = s & " From tblPhaseItem where DivisionID = " & HFApp.DivisionID
                    Call HFApp.SqlExec(s, dbHomefront)
                    
                    s = "insert into tblEstPhases(Phase, Description, GroupPhase, UpdateEstimating, UStmp, TStmp, SortOrder, GroupPhaseValue, DivisionID)"
                    s = s & " select Phase, Description, GroupPhase, UpdateEstimating, UStmp, TStmp, SortOrder, GroupPhaseValue," & .TextMatrix(i, .ColIndex("ID")) & " from tblEstPhases where DivisionID = " & HFApp.DivisionID
                    Call HFApp.SqlExec(s, dbHomefront)
                    
                    s = "insert into tblPOIndex( POIndex, Notes, JCCostCode, JCCategory, ForecastPercent1, ForecastPercent2, ForecastPercent3, ForecastPercent4, ForecastPercent5, ForecastPercent6, ForecastPercent7, ForecastPercent8, ForecastPercent9, ForecastPercent10, ForecastPercent11, ForecastPercent12, StandardText, Description, POGroup, HideQty, HidePrice, TotalOnly, POFormat, SchedTaskID, ReleaseTaskID, PaymentTerm, FOB, ShipVia, Terms, UStmp, TStmp, RetainagePercent, POType, DivisionID)"
                    s = s & " select POIndex, Notes, JCCostCode, JCCategory, ForecastPercent1, ForecastPercent2, ForecastPercent3, ForecastPercent4, ForecastPercent5, ForecastPercent6, ForecastPercent7, ForecastPercent8, ForecastPercent9, ForecastPercent10, ForecastPercent11, ForecastPercent12, StandardText, Description, POGroup, HideQty, HidePrice, TotalOnly, POFormat, SchedTaskID, ReleaseTaskID, PaymentTerm, FOB, ShipVia, Terms, UStmp, TStmp, RetainagePercent, POType," & .TextMatrix(i, .ColIndex("ID"))
                    s = s & " From tblPOIndex where DivisionID = " & HFApp.DivisionID
                    Call HFApp.SqlExec(s, dbHomefront)
                End If
                
            End If
            
            
            s = ""
            s = s & "UPDATE divisions" & vbCrLf
            s = s & "   SET divisioncode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("code"))) & vbCrLf
            s = s & "      ,divisionname=" & DbQuote(Str, .TextMatrix(i, .ColIndex("company"))) & vbCrLf
            s = s & "      ,Address1=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Address1"))) & vbCrLf
            s = s & "      ,Address2=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Address2"))) & vbCrLf
            s = s & "      ,City=" & DbQuote(Str, .TextMatrix(i, .ColIndex("City"))) & vbCrLf
            s = s & "      ,Province=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Province"))) & vbCrLf
            s = s & "      ,Postal=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Postal"))) & vbCrLf
            s = s & "      ,Country=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Country"))) & vbCrLf
            s = s & "      ,County=" & DbQuote(Str, .TextMatrix(i, .ColIndex("County"))) & vbCrLf
            s = s & "      ,TaxNumber=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxNumber"))) & vbCrLf
            s = s & "      ,Fax=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Fax"))) & vbCrLf
            s = s & "      ,Phone=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phone"))) & vbCrLf
            s = s & "      ,BalSheetPrefix=" & DbQuote(Str, .TextMatrix(i, .ColIndex("BalSheetPrefix"))) & vbCrLf
            s = s & "      ,IncomeStmtPrefix=" & DbQuote(Str, .TextMatrix(i, .ColIndex("IncomeStmtPrefix"))) & vbCrLf
            s = s & " WHERE divisionid=" & DbQuote(Num, .TextMatrix(i, .ColIndex("ID"))) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
            

        Next
    End With

'mk -- disable salessheetcosts after removing the option from UI
'mk -- re-enable salessheetcosts for Tribute Communities but don't add the UI option

'    s = ""
'    s = s & "insert appoptions(divisionid,uid,OptionName,OptionValue)" & vbCrLf
'    s = s & "select d.divisionid,'','RecordSalesSheetCosts','False'" & vbCrLf
'    s = s & "from divisions d" & vbCrLf
'    s = s & "left join appoptions o on d.DivisionID=o.divisionid and o.OptionName='RecordSalesSheetCosts'" & vbCrLf
'    s = s & "where o.divisionid is null" & vbCrLf
'    s = s & "" & vbCrLf
'    s = s & "update AppOptions set OptionValue='False' where OptionName='RecordSalesSheetCosts'" & vbCrLf
'    Call HFApp.SqlExec(s, dbHomefront)

Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "SaveDivisions", s)
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub ReadEstimatingOptions()
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    
    cboEstimatingSystem.AddItem "none"
    cboEstimatingSystem.AddItem "CG Visions BIM Pipeline"
    cboEstimatingSystem.AddItem "Sage Estimating"
    cboEstimatingSystem.ListIndex = Val(HFApp.Options.ValueByName("EstimatingSystem"))
    
    
    'CGVisions
    txtCGVisionsRootURI.Text = HFApp.Options.ValueByName("CGVisionsRootURI")
    txtCGVisionsUID.Text = HFApp.Options.ValueByName("CGVisionsUID")
    txtCGVisionsPWD.Text = HFApp.Options.ValueByName("CGVisionsPWD")
    txtBIMBaseModelNumber.Text = HFApp.Options.ValueByName("BIMBaseModelNumber")
    
    optPipelineCommunityStyle(0).Value = True
    optPipelineCommunityStyle(1).Value = HFApp.Options.ValueByName("BIMCommunitySpecificPlansAndOptions") = "true"
    
    
    'New Sage Estimating
    txtSageSqlEstServer.Text = HFApp.Options.ValueByName("SageSqlEstServer")
    cboSageSqlEstDatabase.Text = HFApp.Options.ValueByName("SageSqlEstDatabase")
    txtSageSqlEstUser.Text = HFApp.Options.ValueByName("SageSqlEstUser")
    txtSageSqlEstPswd.Text = HFApp.Options.ValueByName("SageSqlEstPswd")
    chkSageSqlEstConsolidateItems.Value = IIf(HFApp.Options.ValueByName("SageSqlEstConsolidateItems") = "True", vbChecked, vbUnchecked)
    
    
    
    
    Select Case HFApp.Options.ValueByName("TakeoffSystem")
        Case "On-Screen":   cboTOSystem.ListIndex = 1
        Case "PlanSwift":   cboTOSystem.ListIndex = 2
        Case Else:          cboTOSystem.ListIndex = 0
    End Select
    
    txtOnScreenMDB.Text = HFApp.Options.ValueByName("OnScreenMDB")
    txtOnScreenServer.Text = HFApp.Options.ValueByName("OnScreenServer")
    txtOnScreenDatabase.Text = HFApp.Options.ValueByName("OnScreenDatabase")
    txtOnScreenUID.Text = HFApp.Options.ValueByName("OnScreenUID")
    txtOnScreenPWD.Text = HFApp.Options.ValueByName("OnScreenPWD")
    optOnScreenType(1).Value = HFApp.Options.ValueByName("OnScreenType") = "SQL"
    optOnScreenType(0).Value = Not optOnScreenType(1).Value
    
    
    
End Sub










Private Sub txtOnScreenMDB_Change()
    optOnScreenType(0).Value = True
End Sub
Private Sub txtOnScreenServer_Change()
    optOnScreenType(1).Value = True
End Sub
Private Sub txtOnScreenDatabase_Change()
    optOnScreenType(1).Value = True
End Sub
Private Sub txtOnScreenUID_Change()
    optOnScreenType(1).Value = True
End Sub
Private Sub txtOnScreenPWD_Change()
    optOnScreenType(1).Value = True
End Sub
Private Sub txtOnScreenMDB_GotFocus()
    SelectAll txtOnScreenMDB
End Sub
Private Sub txtOnScreenServer_GotFocus()
    SelectAll txtOnScreenServer
End Sub
Private Sub txtOnScreenDatabase_GotFocus()
    SelectAll txtOnScreenDatabase
End Sub
Private Sub txtOnScreenUID_GotFocus()
    SelectAll txtOnScreenUID
End Sub
Private Sub txtOnScreenPWD_GotFocus()
    SelectAll txtOnScreenPWD
End Sub

Private Sub txtSelectAtTakeoffItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbAltMask
        Case KeyCode = vbKeyDelete
            txtSelectAtTakeoffItem.tag = ""
            txtSelectAtTakeoffItem.Text = ""
        Case Else
            Call cmdChooseItem_Click(2)
    End Select
End Sub

Private Sub txtWebAttachmentsFolder_GotFocus()
    SelectAll txtWebAttachmentsFolder
End Sub

Private Sub WriteBuildProOptions()
On Error Resume Next
    
    Dim r As Long
    Dim s As String
    
    With HFApp.Options
        .ValueByName("BuildProCompanyCode") = txtBuildProCompany.Text
        .ValueByName("BuildProEnvironment") = cboBuildProEnvironment.Text
        .ValueByName("BuildProUID") = txtBuildProUID.Text
        .ValueByName("BuildProPwd") = txtBuildProPwd.Text
        .ValueByName("BuildProSendPOsImmediately") = IIf(chkBuildProSendPOsImmediately.Value = vbChecked, "true", "false")
        .ValueByName("BuildProDontSendPhases") = IIf(chkBuildProSendPhases.Value = vbUnchecked, "true", "false")
        
        .ValueByName("EnableTarionFields") = IIf(chkTarion.Value = vbChecked, "true", "false")
        
        If cboBuildProEnvironment.Text = "Production" Then
            .ValueByName("BuildProURL") = "https://xml.hyphensolutions.com/httpreceive.aspx"
            .ValueByName("BuildProIntegrationURL") = "https://integration.hyphensolutions.com/BuildProIntegration.svc"
        Else
            .ValueByName("BuildProURL") = "https://uatxml.hyphensolutions.com/httpreceive.aspx"
            .ValueByName("BuildProIntegrationURL") = "https://uatintegration.hyphensolutions.com/BuildProIntegration.svc"
        End If
        
        .ValueByName("BuildProMultiFamily") = chkMultiFamily.Value = vbChecked
        .ValueByName("BuildProWarrantyOwnerType") = Trim(Me.txtBuildProWarrantyOwnerType.Text)
        .ValueByName("BuildProWarrantyCoOwnerType") = Trim(Me.txtBuildProWarrantyCoOwnerType.Text)
        
        .ValueByName("ScheduleTemplates") = PDL(txtScheduleTemplates.Text)
    End With
    
    s = "delete appoptions where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and optionname like 'EPOReason%'"
    HFApp.SqlExec s
        
    With gEPOReasons
    For r = 1 To .Rows - 2
        s = ""
        s = s & "insert appoptions(uid,divisionid,optionname,optionvalue) values(''"
        s = s & "," & DbQuote(Num, HFApp.DivisionID)
        s = s & "," & DbQuote(Str, "EPOReason" & .TextMatrix(r, 0))
        s = s & "," & DbQuote(Str, .TextMatrix(r, 1))
        s = s & ")"
        HFApp.SqlExec s
        
    Next
    End With

End Sub
Private Sub ReadBuildProOptions()
    Dim r As Long
    Dim s As String
    Dim rs As Recordset
    
    With HFApp.Options
        cboBuildProEnvironment.Clear
        cboBuildProEnvironment.AddItem "User Acceptance Testing (UAT)"
        cboBuildProEnvironment.AddItem "Production"
        Call SetComboBoxListIndex(cboBuildProEnvironment, .ValueByName("BuildProEnvironment"))
        txtBuildProCompany.Text = .ValueByName("BuildProCompanyCode")
        txtBuildProUID.Text = .ValueByName("BuildProUID")
        txtBuildProPwd.Text = .ValueByName("BuildProPwd")
        txtScheduleTemplates.Text = Replace(HFApp.Options.ValueByName("ScheduleTemplates"), "|", vbCrLf)
    
        chkBuildProSendPOsImmediately.Value = IIf(.ValueByName("BuildProSendPOsImmediately") = "true", vbChecked, vbUnchecked)
        chkBuildProSendPhases.Value = IIf(.ValueByName("BuildProDontSendPhases") = "true", vbUnchecked, vbChecked)
        chkTarion.Value = IIf(.ValueByName("EnableTarionFields") = "true", vbChecked, vbUnchecked)
        
        chkMultiFamily.Value = IIf(.ValueByName("BuildProMultiFamily") = "True", vbChecked, vbUnchecked)
        
        txtBuildProWarrantyOwnerType.Text = Trim(.ValueByName("BuildProWarrantyOwnerType"))
        txtBuildProWarrantyCoOwnerType.Text = Trim(.ValueByName("BuildProWarrantyCoOwnerType"))
    
    End With
    
    With gEPOReasons
        s = "select substring(OptionName,10,9999) ReasonCode,OptionValue Category from appoptions where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and optionname like 'EPOReason%'"
        Set rs = HFApp.SqlExec(s, dbHomefront)
        
        .Rows = 1
        r = 1
        While Not rs.EOF
            .AddItem "" & rs("ReasonCode") & vbTab & rs("Category")
            rs.MoveNext
        Wend
        .AddItem ""
    
    
        Set rs = HFApp.SqlExec("select Category,Description from standardcategories where divisionid=" & DbQuote(Num, HFApp.DivisionID))
        .ColComboList(1) = .BuildComboList(rs, "Category,Description")
    
    
    End With
    
End Sub

Private Sub WriteRoleDefinitions()
    HFApp.Options.ValueByName("VendorContactRoles") = PDL(txtVendorRoles.Text)
    HFApp.Options.ValueByName("CustomerContactRoles") = PDL(txtCustomerRoles.Text)
    HFApp.Options.ValueByName("JobContactRoles") = PDL(txtJobRoles.Text)
End Sub

Private Sub ReadRoleDefinitions()
    txtVendorRoles.Text = Replace(HFApp.Options.ValueByName("VendorContactRoles"), "|", vbCrLf)
    txtCustomerRoles.Text = Replace(HFApp.Options.ValueByName("CustomerContactRoles"), "|", vbCrLf)
    txtJobRoles.Text = Replace(HFApp.Options.ValueByName("JobContactRoles"), "|", vbCrLf)
End Sub

Private Function PDL(Value As String) As String
    'converts vbcrlf delimited string to pipe delimited string
    'removes embedded blank lines and trims each element
    Dim i As Long
    Dim s As String
    Dim t As String
    s = ""
    For i = 1 To Parse(Value, , vbCrLf)
        t = Trim(Parse(Value, i, vbCrLf))
        t = Replace(t, "|", "")
        If t <> "" Then s = s & "|" & t
    Next
    PDL = Mid(s, 2)
End Function


Private Sub ReadDepartmentApprovals()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    
    
    With gDepartments
        .Rows = 1
        s = "select * from departments where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by name"
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem "" & rs("deptid") & vbTab & rs("name")
            rs.MoveNext
        Wend
        .AddItem ""
    End With

    
    With gDepartmentApprovers
        .Rows = 1
        s = ""
        s = s & "select d.deptid,p.pm,da.supervisor,da.invoicelimit" & vbCrLf
        s = s & "from departments d " & vbCrLf
        s = s & "join departmentapprovers da on d.deptid=da.deptid" & vbCrLf
        s = s & "join tblprojectmanager p on da.pm=p.pm" & vbCrLf
        s = s & "where d.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "order by 4 desc, 3 asc" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem "" & rs("deptid") & vbTab & rs("pm") & vbTab & rs("supervisor") & vbTab & rs("invoicelimit")
            .RowHidden(.Rows - 1) = True
            rs.MoveNext
        Wend
        .AddItem ""
        .RowHidden(.Rows - 1) = True
        Call .AutoSize(0, .Cols - 1)
        
    End With

    
    With gDepartmentGLs
        .Rows = 1
        s = ""
        s = s & "select d.deptid,g.account,g.description " & vbCrLf
        s = s & "from departments d " & vbCrLf
        s = s & "join departmentglaccounts dg on d.deptid=dg.deptid" & vbCrLf
        s = s & "join glaccounts g on dg.account=g.account and d.divisionid=g.divisionid" & vbCrLf
        s = s & "where d.divisionid=" & DbQuote(Num, HFApp.DivisionID)
        s = s & "order by 2" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem "" & rs("deptid") & vbTab & rs("account") & vbTab & rs("description")
            .RowHidden(.Rows - 1) = True
            rs.MoveNext
        Wend
        .AddItem ""
        .RowHidden(.Rows - 1) = True
    End With
    

Exit Sub
eh: Call errHandler(SRCFILE & "ReadDepartmentApprovals")
End Sub

Private Sub WriteDepartmentApprovals()
On Error GoTo eh

    Dim s As String
    Dim r As Long
    Dim i As Long
    Dim newid As Long
    Dim oldid As String
    
    With gDepartments
    For r = .Rows - 2 To 1 Step -1
        Select Case .RowData(r)
            Case "insert"
                s = ""
                s = s & "insert into departments(divisionid,name) values" & vbCrLf
                s = s & "(" & DbQuote(Num, HFApp.DivisionID)
                s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("name")))
                s = s & ")"
                Call HFApp.SqlExec(s)
                oldid = .TextMatrix(r, .ColIndex("deptid"))
                newid = HFApp.SqlIdentity("departments")
                .TextMatrix(r, .ColIndex("deptid")) = newid
                For i = 1 To gDepartmentGLs.Rows - 1
                    If gDepartmentGLs.TextMatrix(i, gDepartmentGLs.ColIndex("deptid")) = oldid Then gDepartmentGLs.TextMatrix(i, gDepartmentGLs.ColIndex("deptid")) = newid
                Next
                For i = 1 To gDepartmentApprovers.Rows - 1
                    If gDepartmentApprovers.TextMatrix(i, gDepartmentApprovers.ColIndex("deptid")) = oldid Then gDepartmentApprovers.TextMatrix(i, gDepartmentApprovers.ColIndex("deptid")) = newid
                Next
                
                
            Case "update"
                s = ""
                s = s & "update departments set" & vbCrLf
                s = s & "name=" & DbQuote(Str, .TextMatrix(r, .ColIndex("name"))) & vbCrLf
                s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                s = s & "and deptid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("deptid")))
                Call HFApp.SqlExec(s)
                
            Case "delete"
                s = ""
                s = s & "delete departments where deptid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("deptid"))) & vbCrLf
                s = s & "delete departmentglaccounts where deptid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("deptid"))) & vbCrLf
                s = s & "delete departmentapprovers where deptid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("deptid"))) & vbCrLf
                Call HFApp.SqlExec(s)
                Call .RemoveItem(r)
                
        End Select
    Next
    End With




    With gDepartmentGLs
    For r = .Rows - 2 To 1 Step -1
        Select Case .RowData(r)
            Case "insert"
                s = ""
                s = s & "insert into departmentglaccounts(deptid,account) values" & vbCrLf
                s = s & "(" & DbQuote(Num, .TextMatrix(r, .ColIndex("deptid")))
                s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("account")))
                s = s & ")"
                Call HFApp.SqlExec(s)
                
            Case "update"
                s = ""
                s = s & "update departmentglaccounts set" & vbCrLf
                s = s & "account=" & DbQuote(Str, .TextMatrix(r, .ColIndex("account"))) & vbCrLf
                s = s & "where deptid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("deptid")))
                Call HFApp.SqlExec(s)
                
            Case "delete"
                s = ""
                s = s & "delete departmentglaccounts" & vbCrLf
                s = s & "where deptid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("deptid"))) & vbCrLf
                s = s & "and account=" & DbQuote(Str, .TextMatrix(r, .ColIndex("account")))
                Call HFApp.SqlExec(s)
                Call .RemoveItem(r)
                
        End Select
    Next
    End With


    With gDepartmentApprovers
    
    'remove any deleted rows from the grid
    Call .RemoveItem(.Rows - 1)
    For r = .Rows - 1 To 1 Step -1
        If .RowData(r) = "delete" Then
            Call .RemoveItem(r)
        End If
    Next
    
    'sort grid so row number is the sortorder in db
    .ColPosition(.ColIndex("deptid")) = 0
    .ColPosition(.ColIndex("invoicelimit")) = 1
    .ColPosition(.ColIndex("pm")) = 2
    .Col = 0
    .ColSel = 2
    .Sort = flexSortGenericDescending
        
    'delete all approvers in this division
    s = "delete from departmentapprovers where deptid in(select deptid from departments where divisionid=" & DbQuote(Num, HFApp.DivisionID) & ")"
    Call HFApp.SqlExec(s)
    
    'write each approver to db
    For r = 1 To .Rows - 1
        s = ""
        s = s & "insert into departmentapprovers(deptid,pm,supervisor,invoicelimit,sortorder) values" & vbCrLf
        s = s & "(" & DbQuote(Num, .TextMatrix(r, .ColIndex("deptid")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("pm")))
        s = s & "," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Supervisor")))
        s = s & "," & DbQuote(Cur, .TextMatrix(r, .ColIndex("invoicelimit")))
        s = s & "," & DbQuote(Num, r)
        s = s & ")"
        Call HFApp.SqlExec(s)
    Next
    
    
    End With



Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "WriteDepartmentApprovals", s)
        Screen.MousePointer = vbDefault
    End If
End Sub



Private Sub gDepartments_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim b As Boolean
    With gDepartments
        If Row = .Rows - 1 Then
            .RowData(Row) = "insert"
            .TextMatrix(Row, .ColIndex("deptid")) = CreateGUID
            .AddItem ""
            Call gDepartments_BeforeSelChange(0, 0, Row, 0, b)
        End If
        If .RowData(Row) = "" Then
            .RowData(Row) = "update"
        End If
    End With
End Sub
Private Sub gDepartments_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim b As Boolean
    With gDepartments
        If .Row < 1 Or .Row = .Rows - 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift = vbCtrlMask Then
            .RowData(.Row) = "delete"
            .RowHidden(.Row) = True
            .Row = GridNextVisibleRow(gDepartments, .Row)
        End If
    End With
End Sub
Private Sub gDepartments_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    Dim deptid As String
    Dim r As Long
    
    
    deptid = gDepartments.TextMatrix(NewRowSel, gDepartments.ColIndex("deptid"))
    If gDepartments.RowHidden(NewRowSel) Then deptid = ""
    With gDepartmentGLs
        For r = 1 To .Rows - 2
            .RowHidden(r) = .RowData(r) = "delete" Or deptid <> .TextMatrix(r, gDepartments.ColIndex("deptid"))
        Next
        .RowHidden(.Rows - 1) = deptid = ""
    End With
    With gDepartmentApprovers
        For r = 1 To .Rows - 2
            .RowHidden(r) = .RowData(r) = "delete" Or deptid <> .TextMatrix(r, gDepartments.ColIndex("deptid"))
        Next
        .RowHidden(.Rows - 1) = deptid = ""
    End With
End Sub



Private Sub gDepartmentGLs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    gDepartmentGLs.ComboList = "..."
End Sub
Private Sub gDepartmentGLs_KeyDown(KeyCode As Integer, Shift As Integer)
    With gDepartmentGLs
        If .Row < 1 Or .Row = .Rows - 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift = vbCtrlMask And .RowHidden(.Row) = False Then
            .RowData(.Row) = "delete"
            .RowHidden(.Row) = True
            .Row = GridNextVisibleRow(gDepartmentGLs, .Row)
        End If
    End With
End Sub
Private Sub gDepartmentGLs_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim i As Long
    Dim r As Long
    
    s = "select Account,Description from glaccounts where divisionid=" & DbQuote(Num, HFApp.DivisionID)
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Account", s, , , , , , True) Then
    With gDepartmentGLs
        r = Row
        For i = 1 To FPickList.SelectedItems
            If i = 1 And r <> .Rows - 1 Then
                .RowData(r) = "update"
            Else
                If r <> .Rows - 1 Then
                    r = r + 1
                End If
                
                .AddItem "", r
                .RowData(r) = "insert"
            End If
            .TextMatrix(r, .ColIndex("deptid")) = SelectedDeptID
            .TextMatrix(r, .ColIndex("account")) = FPickList.SelectedItem("account", i)
            .TextMatrix(r, .ColIndex("description")) = FPickList.SelectedItem("description", i)
        Next
    End With
    End If
    
End Sub
Private Function SelectedDeptID() As String
    SelectedDeptID = gDepartments.TextMatrix(gDepartments.Row, gDepartments.ColIndex("deptid"))
End Function
Private Sub gDepartmentApprovers_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gDepartmentApprovers
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "pm":      .ComboList = "..."
            Case "InvoiceLimit"
        End Select
    End With
End Sub

Private Sub gDepartmentApprovers_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim i As Long
    Dim r As Long
    
    s = "select PM,PMName Name from tblprojectmanager"
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Approver", s, , , , , , True) Then
    With gDepartmentApprovers
        r = Row
        For i = 1 To FPickList.SelectedItems
            If i = 1 And r <> .Rows - 1 Then
                .RowData(r) = "update"
            Else
                If r <> .Rows - 1 Then
                    r = r + 1
                End If
                
                .AddItem "", r
                .RowData(r) = "insert"
            End If
            .TextMatrix(r, .ColIndex("deptid")) = SelectedDeptID
            .TextMatrix(r, .ColIndex("pm")) = FPickList.SelectedItem("pm", i)
            .TextMatrix(r, .ColIndex("invoicelimit")) = 0
        Next
    End With
    End If
End Sub


Private Sub gDepartmentApprovers_KeyDown(KeyCode As Integer, Shift As Integer)
    With gDepartmentApprovers
        If .Row < 1 Or .Row = .Rows - 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift = vbCtrlMask And .RowHidden(.Row) = False Then
            .RowData(.Row) = "delete"
            .RowHidden(.Row) = True
            .Row = GridNextVisibleRow(gDepartmentApprovers, .Row)
        End If
    End With
End Sub


Private Sub gDocumentClasses_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim b As Boolean
    With gDocumentClasses
        If .Row < 1 Or .Row = .Rows - 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift = vbCtrlMask Then
            .RowData(.Row) = "delete"
            .RowHidden(.Row) = True
            .Row = GridNextVisibleRow(gDocumentClasses, .Row)
        End If
    End With
End Sub


Private Sub gDocumentClasses_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim b As Boolean
    Dim s As String
    With gDocumentClasses
                
        If .ColKey(Col) = "DocumentClass" Then
            'clean the path
            'slashes must not have slashes or spaces before or after
            'path must not start or end with a slash
            s = .EditText
            While InStr(1, s, "\ ", vbTextCompare) <> 0
                s = Replace(s, "\ ", "\")
            Wend
            While InStr(1, s, " \", vbTextCompare) <> 0
                s = Replace(s, " \", "\")
            Wend
            While InStr(1, s, "\\", vbTextCompare) <> 0
                s = Replace(s, "\\", "\")
            Wend
            While left(s, 1) = "\"
                s = Mid(s, 2)
            Wend
            While Right(s, 1) = "\"
                s = Mid(s, 1, Len(s) - 1)
            Wend
            .EditText = s
            
            If s = "" Then
                Cancel = True
                Exit Sub
            End If
            
        End If
        
        If Row = .Rows - 1 Then
            .RowData(Row) = "insert"
            .AddItem ""
        End If
        If .RowData(Row) = "" Then .RowData(Row) = "update"
    End With
End Sub

Private Sub ReadDocumentManagement()
    Dim s As String
    Dim rs As Recordset
    Dim r As Integer
    With gDocumentClasses
        
        .Rows = 1
        r = 0
        s = "select * from dms_documentclasses order by documentclass"
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem "" & rs("customervisible") & vbTab & rs("vendorvisible") & vbTab & rs("documentclass")
            r = r + 1
            .Cell(flexcpData, r, .ColIndex("documentclass")) = "" & rs("documentclass")
            rs.MoveNext
        Wend
        .AddItem ""
        
    End With

End Sub
Private Sub WriteDocumentManagement()
On Error GoTo eh

    Dim s As String
    Dim r As Long
    Dim i As Long
    
    With gDocumentClasses
    For r = .Rows - 2 To 1 Step -1
        Select Case .RowData(r)
            Case "insert"
                s = ""
                s = s & "insert into dms_documentclasses(documentclass,vendorvisible,customervisible) values" & vbCrLf
                s = s & "(" & DbQuote(Str, .TextMatrix(r, .ColIndex("documentclass")))
                s = s & "," & DbQuote(Bit, .TextMatrix(r, .ColIndex("vendorvisible")))
                s = s & "," & DbQuote(Bit, .TextMatrix(r, .ColIndex("customervisible")))
                s = s & ")"
                Call HFApp.SqlExec(s)
                .Cell(flexcpData, r, .ColIndex("documentclass")) = .TextMatrix(r, .ColIndex("documentclass"))
                
            Case "update"
                s = ""
                s = s & "update dms_documentclasses set" & vbCrLf
                s = s & " documentclass=" & DbQuote(Str, .TextMatrix(r, .ColIndex("documentclass"))) & vbCrLf
                s = s & ",vendorvisible=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("vendorvisible"))) & vbCrLf
                s = s & ",customervisible=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("customervisible"))) & vbCrLf
                s = s & "where documentclass=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("documentclass"))) & vbCrLf
                Call HFApp.SqlExec(s)
                .Cell(flexcpData, r, .ColIndex("documentclass")) = .TextMatrix(r, .ColIndex("documentclass"))
                
            Case "delete"
                s = ""
                s = s & "delete dms_documentclasses where documentclass=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("documentclass"))) & vbCrLf
                Call HFApp.SqlExec(s)
                Call .RemoveItem(r)
                
        End Select
    Next
    End With


Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "WriteDocumentManagement", s)
        Screen.MousePointer = vbDefault
    End If
End Sub


