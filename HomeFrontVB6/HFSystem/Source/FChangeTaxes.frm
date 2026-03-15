VERSION 5.00
Begin VB.Form FChangeTaxes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tax Change Wizard"
   ClientHeight    =   5220
   ClientLeft      =   3405
   ClientTop       =   1710
   ClientWidth     =   6495
   Icon            =   "FChangeTaxes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin HFSystem.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1588
      Caption         =   "Tax Changes"
      Description     =   "Change tax groups on estimated items"
      Icon            =   "FChangeTaxes.frx":000C
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   5340
      TabIndex        =   4
      Top             =   4740
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   4140
      TabIndex        =   3
      Top             =   4740
      Width           =   1095
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Change one tax group to another"
      Height          =   315
      Index           =   1
      Left            =   900
      TabIndex        =   2
      Top             =   3180
      Width           =   2715
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Tax rate has changed. Recalculate tax amounts."
      Height          =   315
      Index           =   0
      Left            =   900
      TabIndex        =   0
      Top             =   2160
      Value           =   -1  'True
      Width           =   3795
   End
   Begin HFSystem.VBCombo cboTaxGroup 
      Height          =   240
      Index           =   0
      Left            =   2040
      TabIndex        =   8
      Top             =   2520
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   423
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
   Begin HFSystem.VBCombo cboTaxGroup 
      Height          =   240
      Index           =   1
      Left            =   2040
      TabIndex        =   9
      Top             =   3555
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   423
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
   Begin HFSystem.VBCombo cboTaxGroup 
      Height          =   240
      Index           =   2
      Left            =   2040
      TabIndex        =   10
      Top             =   3810
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   423
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
   Begin VB.Label lblTaxGroup 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tax group"
      Height          =   195
      Index           =   0
      Left            =   1260
      TabIndex        =   7
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label lblTaxGroup 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "New tax group"
      Height          =   195
      Index           =   2
      Left            =   945
      TabIndex        =   6
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label lblTaxGroup 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Old tax group"
      Height          =   195
      Index           =   1
      Left            =   1035
      TabIndex        =   5
      Top             =   3540
      Width           =   945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   -240
      X2              =   26240
      Y1              =   4635
      Y2              =   4635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   -180
      X2              =   26300
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Label Label1 
      Caption         =   $"FChangeTaxes.frx":08E6
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   5655
   End
End
Attribute VB_Name = "FChangeTaxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Function SaveData() As Boolean
    Dim s As String
    
    Dim oldgrp As String
    Dim newgrp As String
    
    If optMode(0) Then
        oldgrp = cboTaxGroup(0).Text
        newgrp = cboTaxGroup(0).Text
    
        If vbCancel = MsgBox("Are you sure you want to reapply tax group " & newgrp & "?", vbOKCancel + vbQuestion, App.ProductName) Then Exit Function
    
    Else
        oldgrp = cboTaxGroup(1).Text
        newgrp = cboTaxGroup(2).Text
        
        If vbCancel = MsgBox("Are you sure you want to replace tax group " & oldgrp & " with " & newgrp & "?", vbOKCancel + vbQuestion, App.ProductName) Then Exit Function
        
    End If
    
    
    s = ""
    s = s & "update estimateitems" & vbCrLf
    s = s & "set budgettaxgroup=t.taxgroup" & vbCrLf
    s = s & "   ,budgetjctax=round(e.budgetpretax * t.jcrate/100,2)" & vbCrLf
    s = s & "   ,budgetjctaxrate=t.jcrate    " & vbCrLf
    s = s & "   ,budgetnjctax=round(e.budgetpretax * t.njcrate/100,2)" & vbCrLf
    s = s & "   ,budgetnjctaxrate=t.njcrate    " & vbCrLf
    s = s & "from estimateitems e" & vbCrLf
    s = s & "    join taxgroups t on (e.DivisionID = t.DivisionID and t.taxgroup=" & DbQuote(Str, newgrp) & ")" & vbCrLf
    s = s & "where e.DivisionID = " & HFApp.DivisionID & " and budgetgenerated=0" & vbCrLf
    s = s & "  and budgettaxgroup=" & DbQuote(Str, oldgrp) & "" & vbCrLf
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "update estimateitems" & vbCrLf
    s = s & "set potaxgroup=t.taxgroup" & vbCrLf
    s = s & "   ,pojctax=round(e.popretax * t.jcrate/100,2) " & vbCrLf
    s = s & "   ,pojctaxrate=t.jcrate    " & vbCrLf
    s = s & "   ,ponjctax=round(e.popretax * t.njcrate/100,2)" & vbCrLf
    s = s & "   ,ponjctaxrate=t.njcrate" & vbCrLf
    s = s & "from estimateitems e" & vbCrLf
    s = s & "    join taxgroups t on (e.DivisionID = t.DivisionID and t.taxgroup=" & DbQuote(Str, newgrp) & ")" & vbCrLf
    s = s & "where e.DivisionID = " & HFApp.DivisionID & " and pogenbatch=0" & vbCrLf
    s = s & "  and potaxgroup=" & DbQuote(Str, oldgrp) & "" & vbCrLf
    Call HFApp.SqlExec(s)
    SaveData = True
End Function

Private Sub LoadData()
    Dim i As Long
Dim rs As Recordset
    'load tax groups into combo boxes
    Set rs = HFApp.SqlExec("SELECT '' TaxGroup, '' Description UNION ALL SELECT TaxGroup, Description FROM TaxGroups where DivisionID = " & HFApp.DivisionID)
    For i = 0 To 2
        cboTaxGroup(i).Clear
    Next
    While Not rs.EOF
        For i = 0 To 2
            cboTaxGroup(i).AddItem "" & rs(0)
        Next
        rs.MoveNext
    Wend
    Call SetListIndex(cboTaxGroup(0), , HFApp.Options(TaxGroupMaterial))
    Call SetListIndex(cboTaxGroup(1), , HFApp.Options(TaxGroupMaterial))
    Call SetListIndex(cboTaxGroup(2), , HFApp.Options(TaxGroupMaterial))

End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then
        If SaveData Then Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call optMode_Click(0)
    Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub optMode_Click(Index As Integer)

    cboTaxGroup(0).Enabled = optMode(0).Value
    cboTaxGroup(1).Enabled = Not optMode(0).Value
    cboTaxGroup(2).Enabled = Not optMode(0).Value
    lblTaxGroup(0).Enabled = cboTaxGroup(0).Enabled
    lblTaxGroup(1).Enabled = cboTaxGroup(1).Enabled
    lblTaxGroup(2).Enabled = cboTaxGroup(2).Enabled
    
End Sub
