VERSION 5.00
Begin VB.Form FAddProperty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Property"
   ClientHeight    =   6075
   ClientLeft      =   2295
   ClientTop       =   2790
   ClientWidth     =   5550
   Icon            =   "FAddProperty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLength 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3420
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "20"
      Top             =   3180
      Width           =   435
   End
   Begin VB.OptionButton optType 
      Caption         =   "Selection List"
      Height          =   300
      Index           =   5
      Left            =   1440
      TabIndex        =   9
      Top             =   3420
      Width           =   1275
   End
   Begin VB.OptionButton optType 
      Caption         =   "Text"
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   7
      Top             =   3180
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.OptionButton optType 
      Caption         =   "Yes/No"
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   2940
      Width           =   2055
   End
   Begin VB.OptionButton optType 
      Caption         =   "Date"
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Top             =   2700
      Width           =   2055
   End
   Begin VB.OptionButton optType 
      Caption         =   "Currency"
      Height          =   315
      Index           =   6
      Left            =   1440
      TabIndex        =   4
      Top             =   2460
      Width           =   2055
   End
   Begin VB.TextBox txtCategory 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1260
      MaxLength       =   200
      TabIndex        =   1
      Top             =   1260
      Width           =   3375
   End
   Begin VB.TextBox txtPickList 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Text            =   "FAddProperty.frx":000C
      Top             =   3720
      Width           =   2985
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   11
      Top             =   5580
      Width           =   915
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Top             =   5580
      Width           =   915
   End
   Begin VB.OptionButton optType 
      Caption         =   "Numeric (decimal)"
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   2220
      Width           =   2055
   End
   Begin VB.OptionButton optType 
      Caption         =   "Numeric (integer)"
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   1980
      Width           =   2055
   End
   Begin VB.TextBox txtFieldName 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1260
      MaxLength       =   200
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Index           =   1
      Left            =   1260
      TabIndex        =   16
      Top             =   1020
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "FAddProperty.frx":002F
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblLength 
      AutoSize        =   -1  'True
      Caption         =   "Length"
      Height          =   195
      Left            =   2880
      TabIndex        =   15
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Data Type"
      Height          =   195
      Index           =   2
      Left            =   1260
      TabIndex        =   14
      Top             =   1740
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property Name"
      Height          =   195
      Index           =   0
      Left            =   1260
      TabIndex        =   13
      Top             =   360
      Width           =   1050
   End
End
Attribute VB_Name = "FAddProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMode As String
Private mFieldName As String

Public Sub AddField()
    mMode = "add"
    FAddProperty.Show vbModal
End Sub

Public Sub EditField(FieldName As String)
    mMode = "edit"
    mFieldName = FieldName
    FAddProperty.Show vbModal
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    Dim s As String
    Dim d As String
    Dim c As Connection
    
    Dim i As Long
    Dim SIZE As Long
    Dim list As String
    
    Select Case True
        Case Index = 0 And mMode = "edit" 'ok
            
            If optType(5) Then
                list = ""
                SIZE = Val(txtLength.Text)
                For i = 1 To Parse(txtPickList.Text, , vbCrLf)
                    s = left(Trim(Parse(txtPickList.Text, i, vbCrLf)), SIZE)
                    If s <> "" Then
                        list = list & "|" & s
                    End If
                Next
                list = Mid(list, 2)
            End If
            
            s = ""
            s = s & "update JobCustomFieldDefs set"
            s = s & " picklist=" & DbQuote(Str, list)
            s = s & ",category=" & DbQuote(Str, txtCategory.Text)
            s = s & " where name=" & DbQuote(Str, txtFieldName.Text, , True)
            Call HFApp.SqlExec(s)
            Unload Me
        
        
        Case Index = 0 And mMode = "add" 'ok
            If Trim(txtFieldName.Text) = "" Then Exit Sub
            If Trim(Me.txtCategory.Text) = "" Then
                MsgBox "Category is required", vbInformation, App.ProductName
                Exit Sub
            End If
            If optType(4).Value And Val(txtLength.Text) <= 0 Then Exit Sub
            If InStr(1, txtFieldName, vbQuote) Then Call MsgBox("The property name may not contain the quote character ("").", vbInformation, App.ProductName): Exit Sub
            If InStr(1, txtPickList, "|") Then Call MsgBox("The selection list may not contain the pipe character (|).", vbInformation, App.ProductName): Exit Sub
            
            If optType(5) Then
                list = ""
                SIZE = -1
                For i = 1 To Parse(txtPickList.Text, , vbCrLf)
                    s = Trim(Parse(txtPickList.Text, i, vbCrLf))
                    If s <> "" Then
                        SIZE = Max(SIZE, Len(s))
                        list = list & "|" & s
                    End If
                Next
                list = Mid(list, 2)
            End If
            
            If optType(0) Then s = " integer"
            If optType(1) Then s = " float"
            If optType(2) Then s = " datetime"
            If optType(3) Then s = " bit"
            If optType(4) Then s = " varchar(" & Val(txtLength.Text) & ")"
            If optType(5) Then s = " varchar(" & IIf(SIZE = -1, 25, SIZE) & ")"
            If optType(6) Then s = " money"
            
            Set c = New Connection
            c.Open HFApp.ConnectionString(dbHomefront) & ";App=HFDBUpgradeWiz"
            Call c.Execute("ALTER TABLE dbo.JobCustomFields ADD " & vbQuote & Trim(txtFieldName.Text) & vbQuote & s)
            Call c.Execute("ALTER TABLE dbo.WorkticketCustomFlds ADD " & vbQuote & Trim(txtFieldName.Text) & vbQuote & s)
            c.Close
            Set c = Nothing
        
            Call HFApp.SqlExec("INSERT INTO JobCustomFieldDefs(Name,Category,PickList) VALUES(" & DbQuote(Str, txtFieldName.Text, , True) & "," & DbQuote(Str, txtCategory.Text) & "," & DbQuote(Str, list) & ")")
            Unload Me
            
        Case Index = 1 'cancel
            Unload Me
            
    End Select
    Exit Sub
eh: MsgBox Parse(Err.Description, Parse(Err.Description, , "]"), "]"), vbExclamation, "Unable to create field"
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    If mMode = "edit" Then Call LoadField
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub optType_Click(Index As Integer)
    
    txtLength.Enabled = Index = 4
    txtLength.BackColor = IIf(Index = 4, vbWindowBackground, vbButtonFace)
    
    txtPickList.Enabled = Index = 5
    txtPickList.BackColor = IIf(Index = 5, vbWindowBackground, vbButtonFace)
    txtPickList.Text = IIf(Index = 5, "", "enter each item on its own line.")
    
End Sub

Private Sub txtLength_Validate(Cancel As Boolean)
    txtLength.Text = Val(txtLength.Text)
    Cancel = Val(txtLength.Text) < 1
End Sub

Private Sub txtFieldName_Change()
    cmdNav(0).Enabled = Trim(txtFieldName.Text) <> ""
End Sub

Private Sub txtFieldName_GotFocus()
    SelectAll txtFieldName
End Sub
Private Sub txtLength_GotFocus()
    SelectAll txtLength


End Sub

Private Sub LoadField()
    Dim s As String
    Dim rs As Recordset
    
    
    optType(5).Value = True
    txtPickList.Enabled = True
    txtFieldName.Enabled = False
    txtLength.Enabled = False
    optType(0).Enabled = False
    optType(1).Enabled = False
    optType(2).Enabled = False
    optType(3).Enabled = False
    optType(4).Enabled = False
    optType(5).Enabled = False
    optType(6).Enabled = False
    
    s = ""
    s = s & "select d.*,c.max_length Size" & vbCrLf
    s = s & "from JobCustomFieldDefs d" & vbCrLf
    s = s & "left outer join sys.columns c on c.object_id=object_id('JobCustomFields') and c.name=d.name" & vbCrLf
    s = s & "where d.name=" & DbQuote(Str, mFieldName)
    Set rs = HFApp.SqlExec(s)
    
    txtFieldName.Text = "" & rs("name")
    txtCategory.Text = "" & rs("category")
    txtLength.Text = "" & rs("size")
    txtPickList.Text = Replace("" & rs("picklist"), "|", vbCrLf)
    
    
End Sub
