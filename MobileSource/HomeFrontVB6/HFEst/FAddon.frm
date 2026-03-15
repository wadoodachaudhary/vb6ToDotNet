VERSION 5.00
Begin VB.Form FAddon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Addon"
   ClientHeight    =   3810
   ClientLeft      =   10830
   ClientTop       =   2295
   ClientWidth     =   4575
   Icon            =   "FAddon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4575
   Begin VB.OptionButton optBasis 
      Caption         =   "Running Total"
      Height          =   600
      Index           =   2
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   900
      Width           =   900
   End
   Begin VB.OptionButton optBasis 
      Caption         =   "Previous SubTotal"
      Height          =   600
      Index           =   1
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   900
      Width           =   900
   End
   Begin VB.OptionButton optBasis 
      Caption         =   "Cost"
      Height          =   600
      Index           =   0
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   900
      Width           =   900
   End
   Begin VB.TextBox txtTitle 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   240
      Width           =   3675
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3300
      TabIndex        =   17
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2130
      TabIndex        =   16
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Frame frmCategories 
      Caption         =   " Categories "
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
      Height          =   1485
      Left            =   2700
      TabIndex        =   18
      Top             =   1710
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CheckBox chkCost 
         Caption         =   "Other"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   1140
         Width           =   1185
      End
      Begin VB.CheckBox chkCost 
         Caption         =   "Labour"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   930
         Width           =   1185
      End
      Begin VB.CheckBox chkCost 
         Caption         =   "Subcontract"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1185
      End
      Begin VB.CheckBox chkCost 
         Caption         =   "Equipment"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   510
         Width           =   1185
      End
      Begin VB.CheckBox chkCost 
         Caption         =   "Material"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.OptionButton optBasis 
      Caption         =   "Lump Sum"
      Height          =   600
      Index           =   3
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   900
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.Frame frmRates 
      Caption         =   " Amount "
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
      Height          =   1485
      Left            =   180
      TabIndex        =   19
      Top             =   1710
      Visible         =   0   'False
      Width           =   2355
      Begin VB.OptionButton optRateType 
         Caption         =   "Margin"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtPercentage 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Text            =   "25%"
         Top             =   390
         Width           =   525
      End
      Begin VB.TextBox txtRate2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   1380
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   1170
         Width           =   795
      End
      Begin VB.TextBox txtRate1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   390
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1170
         Width           =   795
      End
      Begin VB.OptionButton optRateType 
         Caption         =   "Amount Per"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   930
         Width           =   1125
      End
      Begin VB.OptionButton optRateType 
         Caption         =   "Mark-up"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lblRatePer 
         AutoSize        =   -1  'True
         Caption         =   "/"
         Height          =   195
         Left            =   1230
         TabIndex        =   20
         Top             =   1170
         Width           =   75
      End
   End
   Begin VB.Frame frmLumpAmount 
      Caption         =   " Amount "
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
      Height          =   1485
      Left            =   180
      TabIndex        =   22
      Top             =   1710
      Width           =   2355
      Begin VB.TextBox txtLumpSumAmt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   390
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   510
         Width           =   1275
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Title"
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
      Left            =   180
      TabIndex        =   23
      Top             =   210
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cost Basis"
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
      Left            =   210
      TabIndex        =   21
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "FAddon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FAddon::"
Private mMode   As String
Private mCancel As Boolean
Private mGrid   As VSFlexGrid


Public Function CreateAddon(Grid As VSFlexGrid) As Boolean
    Dim r As Long
    mMode = "Create"
    Set mGrid = Grid
    Me.Show vbModal
    CreateAddon = Not mCancel
End Function

Public Function EditAddon(Grid As VSFlexGrid) As Boolean
    Dim r As Long
    mMode = "Edit"
    Set mGrid = Grid
    Me.Show vbModal
    EditAddon = Not mCancel
End Function

Private Sub Form_Load()
    Dim r As Long
    
    Call IniGetForm(Me)
    
    If mMode = "Create" Then
        r = Between(Val(IniGet(AppIni, "Options", "AddonType", 3)), 3, 1)
        optBasis(r).value = True
    Else
    With mGrid
        r = .Row
        txtTitle.Text = .TextMatrix(r, .ColIndex("Title"))
        
        
        Select Case .TextMatrix(r, .ColIndex("Basis"))
            Case "LumpSum":        optBasis(3).value = True:   frmCategories.Visible = False
            Case "Category":       optBasis(0).value = True:   frmCategories.Visible = True
            Case "LastSubTotal":   optBasis(1).value = True:   frmCategories.Visible = False
            Case "RunningTotal":   optBasis(2).value = True:   frmCategories.Visible = False
        End Select
        
        Select Case .TextMatrix(r, .ColIndex("RateType"))
            Case 2 'lump
                frmLumpAmount.Visible = True
                frmRates.Visible = False
            Case 1 'amt per
                optRateType(2).value = True
                frmLumpAmount.Visible = False
                frmRates.Visible = True
            Case 3 'margin
                optRateType(3).value = True
                frmLumpAmount.Visible = False
                frmRates.Visible = True
            Case Else 'mark-up
                optRateType(1).value = True
                frmLumpAmount.Visible = False
                frmRates.Visible = True
        End Select
        
        txtPercentage.Text = Val(.TextMatrix(r, .ColIndex("Percentage"))) & "%"
        txtRate1.Text = format(.TextMatrix(r, .ColIndex("Rate1")), "0.00")
        txtRate2.Text = format(.TextMatrix(r, .ColIndex("Rate2")), "0.00")
        txtLumpSumAmt.Text = format(.TextMatrix(r, .ColIndex("LumpSumAmt")), "0.00")
        
        chkCost(0).value = IIf(.TextMatrix(r, .ColIndex("Mat")) = "True", vbChecked, vbUnchecked)
        chkCost(1).value = IIf(.TextMatrix(r, .ColIndex("Eq")) = "True", vbChecked, vbUnchecked)
        chkCost(2).value = IIf(.TextMatrix(r, .ColIndex("Sub")) = "True", vbChecked, vbUnchecked)
        chkCost(3).value = IIf(.TextMatrix(r, .ColIndex("Lab")) = "True", vbChecked, vbUnchecked)
        chkCost(4).value = IIf(.TextMatrix(r, .ColIndex("Oth")) = "True", vbChecked, vbUnchecked)
        
    End With
    End If
    
    
End Sub


Private Sub cmdNav_Click(Index As Integer)
    Dim r As Long
    
    If Index = 0 Then
        mCancel = False
        With mGrid
            r = .Row
            If r < 1 Then r = 1
            If r > .Rows Then r = .Rows
            
            If mMode = "Create" Then
                .AddItem "", r
            End If
            
            .TextMatrix(r, .ColIndex("Title")) = txtTitle.Text
            Select Case True
                Case optBasis(3).value:            .TextMatrix(r, .ColIndex("Basis")) = "LumpSum"
                Case optBasis(0).value:            .TextMatrix(r, .ColIndex("Basis")) = "Category"
                Case optBasis(1).value:            .TextMatrix(r, .ColIndex("Basis")) = "LastSubTotal"
                Case optBasis(2).value:            .TextMatrix(r, .ColIndex("Basis")) = "RunningTotal"
            End Select
            
            Select Case True
                Case frmLumpAmount.Visible:        .TextMatrix(r, .ColIndex("RateType")) = 2 'lump
                Case optRateType(2).value:         .TextMatrix(r, .ColIndex("RateType")) = 1 'amt per
                Case optRateType(3).value:         .TextMatrix(r, .ColIndex("RateType")) = 3 'margin
                Case Else:                         .TextMatrix(r, .ColIndex("RateType")) = 0 'mark-up
            End Select
            
            .TextMatrix(r, .ColIndex("Percentage")) = Val(Replace(txtPercentage.Text, "%", ""))
            .TextMatrix(r, .ColIndex("Rate1")) = Val(txtRate1.Text)
            .TextMatrix(r, .ColIndex("Rate2")) = Val(txtRate2.Text)
            .TextMatrix(r, .ColIndex("LumpSumAmt")) = Val(txtLumpSumAmt.Text)
            .TextMatrix(r, .ColIndex("Mat")) = chkCost(0).value = vbChecked
            .TextMatrix(r, .ColIndex("Eq")) = chkCost(1).value = vbChecked
            .TextMatrix(r, .ColIndex("Sub")) = chkCost(2).value = vbChecked
            .TextMatrix(r, .ColIndex("Lab")) = chkCost(3).value = vbChecked
            .TextMatrix(r, .ColIndex("Oth")) = chkCost(4).value = vbChecked
            
        End With
    Else
        mCancel = True
    End If
    
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    Call IniPutForm(Me)
    
    If optBasis(0).value Then i = 0
    If optBasis(1).value Then i = 1
    If optBasis(2).value Then i = 2
    If optBasis(3).value Then i = 3
    Call IniPut(AppIni, "Options", "AddonType", i)
End Sub

Private Sub optBasis_Click(Index As Integer)
    Select Case Index
        Case 0 'cost
            frmRates.Visible = True
            frmLumpAmount.Visible = False
            frmCategories.Visible = True
        Case 1, 2 'previous or running total
            frmRates.Visible = True
            frmLumpAmount.Visible = False
            frmCategories.Visible = False
        Case 3 'lump sum
            frmRates.Visible = False
            frmLumpAmount.Visible = True
            frmCategories.Visible = False
    End Select
End Sub

Private Sub optRateType_Click(Index As Integer)
    txtPercentage.Enabled = optRateType(1).value
    txtRate1.Enabled = optRateType(2).value
    txtRate2.Enabled = optRateType(2).value
End Sub


Private Sub txtLumpSumAmt_GotFocus()
    SelectAll txtLumpSumAmt
End Sub

Private Sub txtLumpSumAmt_Validate(Cancel As Boolean)
    txtLumpSumAmt.Text = format(Val(txtLumpSumAmt.Text), "0.00")
End Sub
Private Sub txtPercentage_Change()
    Me.optRateType(1).Enabled = True
End Sub

Private Sub txtPercentage_GotFocus()
    SelectAll txtPercentage
End Sub

Private Sub txtPercentage_Validate(Cancel As Boolean)
    txtPercentage.Text = Val(txtPercentage.Text) & "%"
End Sub

Private Sub txtRate1_Change()
    Me.optRateType(2).Enabled = True
End Sub

Private Sub txtRate1_GotFocus()
    SelectAll txtRate1
End Sub

Private Sub txtRate1_Validate(Cancel As Boolean)
    txtRate1.Text = format(Val(txtRate1.Text), "0.00")
End Sub

Private Sub txtRate2_Change()
    Me.optRateType(2).Enabled = True
End Sub

Private Sub txtRate2_GotFocus()
    SelectAll txtRate2
End Sub

Private Sub txtRate2_Validate(Cancel As Boolean)
    txtRate2.Text = format(Val(txtRate2.Text), "0.00")
End Sub

Private Sub txtTitle_GotFocus()
    SelectAll txtTitle
End Sub
