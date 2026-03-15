VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSelectedList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HomeFront - Item Selection List"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   Icon            =   "frmSelectedList.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   3000
      Width           =   4575
   End
   Begin MSAdodcLib.Adodc curMainList 
      Height          =   330
      Left            =   3480
      Top             =   2520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "curMainList"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Height          =   375
      Left            =   2880
      Picture         =   "frmSelectedList.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "&Exit"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      ToolTipText     =   "Remove All"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Add All"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      ToolTipText     =   "Remove"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "Add"
      Top             =   120
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3000
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
   Begin MSAdodcLib.Adodc curAddFromList 
      Height          =   330
      Left            =   480
      Top             =   2520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "curAddFromList"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSelectedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mField1 As String
Dim mField2 As String
Dim MyMainTableName As String
Dim MyAddFromTableName As String
Dim gcall As Integer
Public Sub ShowForm(loadtype As Integer)
gcall = loadtype
Me.Show vbModal
End Sub

Private Sub Command1_Click()
On Error Resume Next
If List1.ListIndex >= 0 Then
    List2.AddItem List1.list(List1.ListIndex)
    AddToDatabase (MyLookUp(MyAddFromTableName, mField1, mField2, "" & List1.list(List1.ListIndex), "S"))
    List1.RemoveItem (List1.ListIndex)
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If List2.ListIndex >= 0 Then
    List1.AddItem List2.list(List2.ListIndex)
    DeleteFromDatabase (MyLookUp(MyAddFromTableName, mField1, mField2, "" & List2.list(List2.ListIndex), "S"))
    List2.RemoveItem (List2.ListIndex)
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim i, j As Integer
j = List1.ListCount
For i = 0 To j - 1
    List2.AddItem List1.list(0)
    AddToDatabase (MyLookUp(MyAddFromTableName, mField1, mField2, "" & List1.list(0), "S"))
    List1.RemoveItem (0)
Next i
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim i, j As Integer
j = List2.ListCount
For i = 0 To j - 1
    List1.AddItem List2.list(0)
    DeleteFromDatabase (MyLookUp(MyAddFromTableName, mField1, mField2, "" & List2.list(0), "S"))
    List2.RemoveItem (0)
Next i
End Sub

Private Sub Command5_Click()
Unload Me
Set frmSelectedList = Nothing
End Sub
Private Sub Command6_Click()
On Error Resume Next
Dim i, j As Integer
j = List2.ListCount
With curMainList.Recordset
    For i = 0 To j - 1
        List2.ListIndex = i
        .filter = adFilterNone
        .filter = "Area='" & MyLookUp(MyAddFromTableName, mField1, mField2, "" & List2.Text, "S") & "'"
        If .EOF Then
           .AddNew
           !Area = MyLookUp(MyAddFromTableName, mField1, mField2, "" & List2.Text, "S")
           !Sales_Person_ID = frmuser.txtUserId.Text
           .Update
        End If
    Next i
End With
End Sub

Private Sub Form_Load()

If gcall = 2 Then
    MyAddFromTableName = "tblLocality"
    MyMainTableName = "tblSalesPersonArea"
    mField1 = "Area"
    mField2 = "Description"
Else
    MyAddFromTableName = "Divisions"
    MyMainTableName = "DivisionUsers"
    mField1 = "DivisionID"
    mField2 = "DivisionName"

End If
curMainList.ConnectionString = HFApp.ConnectionString(dbHomeFront)
curMainList.CommandType = adCmdText
If gcall = 2 Then
    curMainList.RecordSource = "select * from tblSalesPersonArea where Sales_Person_ID='" & frmuser.txtUserId.Text & "'"
Else
    curMainList.RecordSource = "select * from  DivisionUsers  where userID=" & DbQuote(str, frmuser.txtUserId.Text)
End If
curMainList.Refresh

curAddFromList.ConnectionString = HFApp.ConnectionString(dbHomeFront)
curAddFromList.CommandType = adCmdText
curAddFromList.RecordSource = "select * from " & MyAddFromTableName
curAddFromList.Refresh

Call LoadList
Call RemoveFromList
If gcall = 2 Then
    Command6.Caption = "Assign Selected Communities to " & frmuser.txtUserId.Text
Else
    Command6.Caption = "Assign Selected Divisions to " & frmuser.txtUserId.Text
End If
End Sub

Public Sub LoadList()
On Error Resume Next
With curAddFromList.Recordset
    If Not .BOF Then .MoveFirst
    Do Until .EOF
        If gcall = 2 Then
            List1.AddItem "" & !Description
        Else
            List1.AddItem "" & !DivisionName
        End If
       .MoveNext
    Loop
End With

With curMainList.Recordset
    If Not .BOF Then .MoveFirst
    Do Until .EOF
       If gcall = 2 Then
            List2.AddItem MyLookUp("tblLocality", "Description", "Area", "" & !Area, "S")
       Else
            List2.AddItem MyLookUp("Divisions", "DivisionName", "DivisionID", "" & !DivisionID, "S")
       End If
       .MoveNext
    Loop
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Select Case gcall
Case 1
    frmSalesPersonSetup.WindowState = 0
Case 2
    frmuser.WindowState = 0
Case Else
End Select
End Sub

Public Sub RemoveFromList()
On Error Resume Next
Dim i, j As Integer
j = List1.ListCount
With curMainList.Recordset
    If Not .BOF Then .MoveFirst
    Do Until .EOF
       For i = 0 To j - 1
           List1.ListIndex = i
           If gcall = 2 Then
                If List1.Text = MyLookUp("tblLocality", "Description", "Area", "" & !Area, "S") Then
                    List1.RemoveItem (List1.ListIndex)
                End If
           Else
                If List1.Text = MyLookUp("Divisions", "DivisionName", "DivisionID", "" & !DivisionID, "N") Then
                    List1.RemoveItem (List1.ListIndex)
                End If
           End If
       Next i
       .MoveNext
    Loop
End With
End Sub
Public Function AddToDatabase(vArea As String)
With curMainList.Recordset
    .filter = adFilterNone
    If gcall = 2 Then
        .filter = "Area='" & vArea & "'"
    Else
        .filter = "DivisionID='" & vArea & "'"
    End If
    If .EOF Then
       .AddNew
        If gcall = 2 Then
            !Area = vArea
            !Sales_Person_ID = frmuser.txtUserId.Text
        Else
            !DivisionID = vArea
            !UserID = frmuser.txtUserId.Text
        End If
       .Update
    End If
End With
End Function
Public Function DeleteFromDatabase(vArea As String)
With curMainList.Recordset
    .filter = adFilterNone
    If gcall = 2 Then
        .filter = "Area='" & vArea & "'"
    Else
        .filter = "DivisionID='" & vArea & "'"
    End If
    If Not .EOF Then
       .Delete
    End If
End With
End Function

Private Sub List1_DblClick()
Call Command1_Click
End Sub

Private Sub List2_Click()
'Call Command2_Click
End Sub
