VERSION 5.00
Begin VB.Form FHome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Home"
   ClientHeight    =   10590
   ClientLeft      =   1650
   ClientTop       =   675
   ClientWidth     =   15180
   Icon            =   "FHome.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10590
   ScaleWidth      =   15180
   Begin VB.Image Task 
      Height          =   960
      Index           =   31
      Left            =   1020
      MouseIcon       =   "FHome.frx":000C
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":0316
      Tag             =   "Custom Requests"
      Top             =   3750
      Width           =   690
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   16
      X1              =   4080
      X2              =   4080
      Y1              =   2580
      Y2              =   3330
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   6
      X1              =   2400
      X2              =   3720
      Y1              =   2820
      Y2              =   2100
   End
   Begin VB.Image Task 
      Height          =   720
      Index           =   30
      Left            =   3795
      MouseIcon       =   "FHome.frx":2658
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":2962
      Tag             =   "Field PO's"
      Top             =   3420
      Width           =   660
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   28
      Left            =   7785
      MouseIcon       =   "FHome.frx":4264
      Picture         =   "FHome.frx":456E
      ToolTipText     =   "no accounting system is configured"
      Top             =   300
      Width           =   885
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   26
      Left            =   12067
      MouseIcon       =   "FHome.frx":7094
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":739E
      Tag             =   "Library Mass Change"
      Top             =   7380
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   15
      Visible         =   0   'False
      X1              =   2310
      X2              =   3720
      Y1              =   1110
      Y2              =   600
   End
   Begin VB.Image Task 
      Height          =   855
      Index           =   25
      Left            =   3705
      MouseIcon       =   "FHome.frx":A0AC
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":A3B6
      Tag             =   "Send RFQ's"
      Top             =   330
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   14
      X1              =   9570
      X2              =   8430
      Y1              =   8670
      Y2              =   8670
   End
   Begin VB.Image Task 
      Height          =   1080
      Index           =   24
      Left            =   9375
      MouseIcon       =   "FHome.frx":C960
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":CC6A
      Tag             =   "WriteToWeb"
      Top             =   8355
      Width           =   1260
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   13
      X1              =   8520
      X2              =   9660
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Image Task 
      Height          =   960
      Index           =   18
      Left            =   9630
      MouseIcon       =   "FHome.frx":1138C
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":11696
      Tag             =   "default vendors"
      Top             =   6480
      Width           =   765
   End
   Begin VB.Image Task 
      Height          =   870
      Index           =   8
      Left            =   7755
      MouseIcon       =   "FHome.frx":13DD8
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":140E2
      Tag             =   "vendor setup"
      Top             =   6480
      Width           =   705
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   12
      X1              =   2130
      X2              =   3510
      Y1              =   6840
      Y2              =   7620
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   11
      X1              =   2340
      X2              =   3750
      Y1              =   1230
      Y2              =   1860
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   10
      X1              =   4500
      X2              =   5670
      Y1              =   1890
      Y2              =   1260
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   9
      X1              =   4470
      X2              =   5790
      Y1              =   2100
      Y2              =   2670
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   1
      X1              =   6540
      X2              =   7920
      Y1              =   2910
      Y2              =   3690
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   8
      X1              =   6540
      X2              =   9870
      Y1              =   2760
      Y2              =   2040
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   7
      X1              =   6510
      X2              =   7770
      Y1              =   1110
      Y2              =   690
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   5
      X1              =   6270
      X2              =   7650
      Y1              =   8670
      Y2              =   8670
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   0
      X1              =   2100
      X2              =   3510
      Y1              =   8640
      Y2              =   7800
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   2
      X1              =   4200
      X2              =   5490
      Y1              =   7650
      Y2              =   6750
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   3
      X1              =   6300
      X2              =   7710
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Flow 
      BorderColor     =   &H00808000&
      Index           =   4
      X1              =   4200
      X2              =   5490
      Y1              =   7800
      Y2              =   8640
   End
   Begin VB.Image Task 
      Height          =   870
      Index           =   10
      Left            =   1425
      MouseIcon       =   "FHome.frx":161C4
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":164CE
      Tag             =   "purchase orders"
      Top             =   8400
      Width           =   645
   End
   Begin VB.Image Task 
      Height          =   990
      Index           =   12
      Left            =   5295
      MouseIcon       =   "FHome.frx":182F8
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":18602
      Tag             =   "edit models and options"
      Top             =   8400
      Width           =   1155
   End
   Begin VB.Image Task 
      Height          =   1035
      Index           =   14
      Left            =   7575
      MouseIcon       =   "FHome.frx":1C214
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":1C51E
      Tag             =   "open an estimating worksheet"
      Top             =   8385
      Width           =   1050
   End
   Begin VB.Image Task 
      Height          =   1035
      Index           =   13
      Left            =   1230
      MouseIcon       =   "FHome.frx":1FE84
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":2018E
      Tag             =   "community standards"
      Top             =   6480
      Width           =   1020
   End
   Begin VB.Image Task 
      Height          =   720
      Index           =   17
      Left            =   12120
      MouseIcon       =   "FHome.frx":238CC
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":23BD6
      Tag             =   "community setup"
      Top             =   2340
      Width           =   945
   End
   Begin VB.Image Task 
      Height          =   990
      Index           =   16
      Left            =   12240
      MouseIcon       =   "FHome.frx":26018
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":26322
      Tag             =   "job cost codes"
      Top             =   600
      Width           =   690
   End
   Begin VB.Image Task 
      Height          =   750
      Index           =   19
      Left            =   12045
      MouseIcon       =   "FHome.frx":2877C
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":28A86
      Tag             =   "series"
      Top             =   4020
      Width           =   960
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   20
      Left            =   12240
      MouseIcon       =   "FHome.frx":2B048
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":2B352
      Tag             =   "option groups"
      Top             =   5760
      Width           =   585
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   23
      Left            =   1650
      MouseIcon       =   "FHome.frx":2D02C
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":2D336
      Tag             =   "Prepare Job Quote"
      Top             =   795
      Width           =   630
   End
   Begin VB.Image Task 
      Height          =   945
      Index           =   22
      Left            =   13770
      MouseIcon       =   "FHome.frx":2F1F8
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":2F502
      Tag             =   "community phases setup"
      Top             =   2340
      Width           =   795
   End
   Begin VB.Image Task 
      Height          =   975
      Index           =   21
      Left            =   13860
      MouseIcon       =   "FHome.frx":31CA4
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":31FAE
      Tag             =   "option categories"
      Top             =   5760
      Width           =   915
   End
   Begin VB.Image Task 
      Height          =   975
      Index           =   0
      Left            =   13740
      MouseIcon       =   "FHome.frx":34EA8
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":351B2
      Tag             =   "job cost categories"
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Task 
      Height          =   855
      Index           =   7
      Left            =   7950
      MouseIcon       =   "FHome.frx":37DA0
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":380AA
      Tag             =   "send purchase orders"
      Top             =   3420
      Width           =   795
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   4
      Left            =   7815
      MouseIcon       =   "FHome.frx":3A48C
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":3A796
      Tag             =   "post budgets"
      Top             =   300
      Width           =   885
   End
   Begin VB.Image lblRegion 
      Height          =   1485
      Index           =   1
      Left            =   60
      Picture         =   "FHome.frx":3D2BC
      Tag             =   "job"
      Top             =   5310
      Width           =   435
   End
   Begin VB.Image lblRegion 
      Height          =   1740
      Index           =   0
      Left            =   60
      Picture         =   "FHome.frx":3F506
      Tag             =   "job"
      Top             =   60
      Width           =   450
   End
   Begin VB.Image Task 
      Height          =   960
      Index           =   15
      Left            =   13830
      MouseIcon       =   "FHome.frx":41EF8
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":42202
      Tag             =   "project managers"
      Top             =   4020
      Width           =   810
   End
   Begin VB.Image Task 
      Height          =   930
      Index           =   9
      Left            =   5520
      MouseIcon       =   "FHome.frx":44B44
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":44E4E
      Tag             =   "edit item prices"
      Top             =   6480
      Width           =   690
   End
   Begin VB.Image Task 
      Height          =   765
      Index           =   11
      Left            =   3540
      MouseIcon       =   "FHome.frx":47078
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":47382
      Tag             =   "edit item database"
      Top             =   7440
      Width           =   630
   End
   Begin VB.Shape Region 
      BackColor       =   &H00F3F0E5&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   10485
      Index           =   2
      Left            =   11370
      Top             =   60
      Width           =   3735
   End
   Begin VB.Image Task 
      Height          =   765
      Index           =   3
      Left            =   5775
      MouseIcon       =   "FHome.frx":48D44
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":4904E
      Tag             =   "issue budgets"
      Top             =   870
      Width           =   915
   End
   Begin VB.Image Task 
      Height          =   1005
      Index           =   5
      Left            =   5700
      MouseIcon       =   "FHome.frx":4B538
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":4B842
      Tag             =   "issue po's"
      Top             =   2550
      Width           =   1065
   End
   Begin VB.Shape Region 
      BackColor       =   &H00F3F0E5&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5145
      Index           =   1
      Left            =   60
      Top             =   5310
      Width           =   11220
   End
   Begin VB.Image Task 
      Height          =   675
      Index           =   2
      Left            =   3780
      MouseIcon       =   "FHome.frx":4F10C
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":4F416
      Tag             =   "Job setup"
      Top             =   1770
      Width           =   690
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   1
      Left            =   1440
      MouseIcon       =   "FHome.frx":50CF4
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":50FFE
      Tag             =   "Retrieve Jobs from Sales Center"
      Top             =   2550
      Width           =   1320
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   27
      Left            =   1440
      MouseIcon       =   "FHome.frx":54F28
      Picture         =   "FHome.frx":55232
      ToolTipText     =   "no sales system is configured"
      Top             =   2550
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   29
      Left            =   9870
      MouseIcon       =   "FHome.frx":5915C
      Picture         =   "FHome.frx":59466
      ToolTipText     =   "no accounting system is configured"
      Top             =   1740
      Width           =   885
   End
   Begin VB.Image Task 
      Height          =   915
      Index           =   6
      Left            =   9840
      MouseIcon       =   "FHome.frx":5BF8C
      MousePointer    =   99  'Custom
      Picture         =   "FHome.frx":5C296
      Tag             =   "post purchase orders"
      Top             =   1740
      Width           =   885
   End
   Begin VB.Shape Region 
      BackColor       =   &H00F3F0E5&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5145
      Index           =   0
      Left            =   90
      Top             =   90
      Width           =   11220
   End
   Begin VB.Label lblScreenSize 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"FHome.frx":5EDBC
      Height          =   10605
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15120
   End
End
Attribute VB_Name = "FHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Call ConfigForm
End Sub

Public Sub ConfigForm()

    'hide sales stuff if not installed
    Task(1).Visible = HFApp.Options(SalesSystem) <> SalesSystems.asNone
    Task(27).Visible = Not Task(1).Visible
    
    Task(24).Visible = HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440
    Flow(14).Visible = HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440
    
    'hide accounting stuff if not installed
    Task(4).Visible = HFApp.Options(AccountingSystem) <> AccountingSystems.asNone
    Task(28).Visible = Not Task(4).Visible
    
    Task(6).Visible = HFApp.Options(AccountingSystem) <> AccountingSystems.asNone
    Task(29).Visible = Not Task(6).Visible
    
    If Not HFApp.UserPermission("SetupJobs") Then Task(2).Enabled = False
    If Not HFApp.UserPermission("EditAssemblies") Then Task(12).Enabled = False
    If Not HFApp.UserPermission("EditItemDb") Then Task(11).Enabled = False
    If Not HFApp.UserPermission("SetupVendorPricing") Then Task(9).Enabled = False
    If Not HFApp.UserPermission("SetupPOIndex") Then Task(10).Enabled = False
    If Not HFApp.UserPermission("OpenEstimatingWorkSheets") Then Task(14).Enabled = False
    If Not HFApp.UserPermission("SetupVendors") Then Task(8).Enabled = False
    If Not HFApp.UserPermission("SetupDefaultVendor") Then Task(18).Enabled = False
    If Not HFApp.UserPermission("SendPO") Then Task(7).Enabled = False
    If Not HFApp.UserPermission("PostCommitments") Then Task(6).Enabled = False
    If Not HFApp.UserPermission("PostBudgets") Then Task(4).Enabled = False
    If Not HFApp.UserPermission("SetupCommunity") Then Task(17).Enabled = False
    If Not HFApp.UserPermission("SetupCommunity") Then Task(22).Enabled = False
    If Not HFApp.UserPermission("Setupseries") Then Task(19).Enabled = False
    If Not HFApp.UserPermission("SetupMajorGroup") Then Task(20).Enabled = False
    If Not HFApp.UserPermission("SetupCategory") Then Task(21).Enabled = False
    If HFApp.UserPermission("EditAssemblies") Then Task(26).Visible = True
    If Not HFApp.UserPermission("EditAssemblies") Then Task(26).Enabled = False
    If Not HFApp.UserPermission("TaskEntQuote") Then Task(31).Enabled = False
    
End Sub

Private Sub Form_Resize()
On Error GoTo eh
Const margin = 60
    Dim i As Long
        
    'normalized scaling factors
    Dim X1 As Double
    Dim Y1 As Double
    Dim X2 As Double
    Dim Y2 As Double
    
    'form is minimized
    If Me.ScaleHeight < 1 Then Exit Sub
    
    'Regions
    Region(0).Move margin, margin, (Me.ScaleWidth - 3 * margin) * 3 / 4, (Me.ScaleHeight - 3 * margin) / 2
    lblRegion(0).Move Region(0).Left, Region(0).Top
    Region(1).Move margin, Region(0).Height + 2 * margin, Region(0).Width, Region(0).Height
    lblRegion(1).Move Region(1).Left, Region(1).Top
    Region(2).Move Region(0).Width + 2 * margin, margin, Me.ScaleWidth - 3 * margin - Region(0).Width, Me.ScaleHeight - 2 * margin
    
    'Tasks
    For i = Task.LBound To Task.UBound
        X1 = (Task(i).Left + (Task(i).Width / 2)) / lblScreenSize.Width
        Y1 = (Task(i).Top + (Task(i).Height / 2)) / lblScreenSize.Height
        Task(i).Move X1 * Me.ScaleWidth - (Task(i).Width / 2), Y1 * Me.ScaleHeight - (Task(i).Height / 2)

'        X1 = Task(i).left / lblScreenSize.Width
'        Y1 = Task(i).Top / lblScreenSize.Height
'        Task(i).Move X1 * Me.ScaleWidth, Y1 * Me.ScaleHeight

     Next
    
    'Lines
    For i = Flow.LBound To Flow.UBound
        X1 = Flow(i).X1 / lblScreenSize.Width
        Y1 = Flow(i).Y1 / lblScreenSize.Height
        X2 = Flow(i).X2 / lblScreenSize.Width
        Y2 = Flow(i).Y2 / lblScreenSize.Height
        Flow(i).X1 = X1 * Me.ScaleWidth
        Flow(i).X2 = X2 * Me.ScaleWidth
        Flow(i).Y1 = Y1 * Me.ScaleHeight
        Flow(i).Y2 = Y2 * Me.ScaleHeight
    Next
    
    lblScreenSize.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
eh: Exit Sub
End Sub


Private Sub Task_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Task(Index).tag <> "" And Button = vbLeftButton Then
        Call FMain.RunTask(Task(Index).tag, Shift = vbCtrlMask)
    End If
End Sub


