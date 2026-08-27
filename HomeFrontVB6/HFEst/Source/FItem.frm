VERSION 5.00
Begin VB.Form FItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item"
   ClientHeight    =   2700
   ClientLeft      =   2205
   ClientTop       =   4575
   ClientWidth     =   5340
   Icon            =   "FItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboCostType 
      Height          =   315
      ItemData        =   "FItem.frx":000C
      Left            =   1230
      List            =   "FItem.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2310
      Width           =   1605
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   4980
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1350
      Width           =   285
   End
   Begin VB.TextBox txtPhase 
      Height          =   285
      Left            =   1230
      MaxLength       =   100
      TabIndex        =   0
      Top             =   1350
      Width           =   3735
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Left            =   1230
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1650
      Width           =   4005
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1230
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1980
      Width           =   4005
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   4290
      TabIndex        =   5
      Top             =   150
      Width           =   945
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   1
      Left            =   4290
      TabIndex        =   6
      Top             =   570
      Width           =   945
   End
   Begin VB.Label lblCostType 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cost Type"
      Height          =   195
      Left            =   465
      TabIndex        =   11
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phase Number"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   10
      Top             =   1380
      Width           =   1050
   End
   Begin VB.Label lblCaption 
      Caption         =   "Enter a phase, item number and description for the new item."
      Height          =   585
      Left            =   210
      TabIndex        =   9
      Top             =   540
      Width           =   3555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Item Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   8
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   390
      TabIndex        =   7
      Top             =   2010
      Width           =   795
   End
End
Attribute VB_Name = "FItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FItem::"
Private mPhase       As String
Private mItemNumber  As String
Private mCostType    As String
Private mDescription As String
Private mCancel      As Boolean
Private mMode        As String


Public Function Add(Optional Phase As String, Optional Item As String, Optional Description As String) As Boolean
On Error Resume Next
    Load Me
    mMode = "Add"
    mCancel = False
    
    mPhase = Phase
    mItemNumber = Left(Item, Len(Item) - 1)
    mDescription = Description
    txtPhase.Text = mPhase
    txtItem.Text = mItemNumber
    txtDescription.Text = mDescription
    cboCostType.ListIndex = 0
    Me.Caption = "Add Item"
    lblCaption = "Enter a phase, item number and description for the new item."
    Me.Show vbModal
    
    Phase = mPhase
    Item = mItemNumber & mCostType
    Description = mDescription
    Add = Not mCancel
End Function

Public Function Renumber(Phase As String, Item As String) As Boolean
On Error Resume Next
    Load Me
    mMode = "Renumber"
    mCancel = False
    
    mPhase = Phase
    mItemNumber = Left(Item, Len(Item) - 1)
    
    lblDescription.Visible = False
    txtDescription.Visible = False
    lblCostType.Visible = False
    cboCostType.Visible = False
    Me.Height = 2505
    
    txtPhase.Text = mPhase
    txtItem.Text = mItemNumber
    
    Me.Caption = "Renumber Item"
    lblCaption = "Enter a new phase and item number for this item."
    Me.Show vbModal
    Renumber = Not mCancel
End Function

Public Function Duplicate(Phase As String, Item As String, Description As String) As Boolean
On Error Resume Next
    Load Me
    mMode = "Duplicate"
    mCancel = False
    mPhase = Phase
    mItemNumber = Item
    txtPhase.Text = ""
    txtItem.Text = ""
    txtDescription.Text = Description
    Me.Caption = "Duplicate Item"
    lblCaption = "Enter a new phase and item number for this item."
    Me.Show vbModal
    Duplicate = Not mCancel
End Function



Private Sub txtPhase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click
End Sub


Private Sub cmdBrowse_Click()
    txtPhase.SetFocus
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Phase", "select phase,description from tblEstPhases where DivisionID = " & HFApp.DivisionID & " and groupphase=0", txtPhase.Text) Then
        txtPhase.Text = FPickList.SelectedItem(1)
    End If
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0
            If SaveData Then
                mPhase = txtPhase.Text
                mItemNumber = txtItem.Text
                mDescription = txtDescription.Text
                mCancel = False
                Unload Me
            End If
        Case 1
            mCancel = True
            Unload Me
    End Select
End Sub


Private Sub Form_Load()
    Call IniGetForm(Me)
    txtPhase.MaxLength = HFApp.EstFieldSize("Phase")
    txtItem.MaxLength = HFApp.EstFieldSize("Item")
    txtDescription.MaxLength = HFApp.EstFieldSize("ItemDesc")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub txtItem_GotFocus()
    SelectAll txtItem
End Sub
Private Sub txtPhase_GotFocus()
    SelectAll txtPhase
End Sub
Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub

Public Function SaveData() As Boolean
    Select Case mMode
        Case "Add":       SaveData = Save_Add()
        Case "Renumber":  SaveData = Save_Renumber()
        Case "Duplicate": SaveData = Save_Duplicate()
    End Select
End Function


Public Function Save_Add() As Boolean
On Error GoTo eh:

    Dim rc As Long
    Dim s  As String
    Dim i  As Long
    Dim DontUseParts As Boolean
    
    DontUseParts = HFApp.Options(AccountingSystem) <> asMasterBuilder Or HFApp.Options.ValueByName("DontUseMBParts") = "True"
   
   
    If Trim(txtPhase.Text) = "" Or Trim(txtItem.Text) = "" Or Trim(txtDescription.Text) = "" Then
        MsgBox "Phase, item and description are all required", vbExclamation, App.ProductName
        Exit Function
    End If
            
    If HFApp.SqlExec("select * from tblEstPhases where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, txtPhase.Text, , True), dbHomefront).EOF Then
        If vbCancel = MsgBox("Phase " & Trim(txtPhase.Text) & " doesn't exist. Do you want to add it?", vbQuestion + vbOKCancel, App.ProductName) Then Exit Function
        Screen.MousePointer = vbHourglass
        Call HFApp.SqlExec("insert into tblestphases(DivisionID,phase,description,groupphase,sortorder) values(" & HFApp.DivisionID & "," & DbQuote(Str, txtPhase.Text, , True) & ",'untitled',0," & DbQuote(Num, txtPhase.Text, , True) & ")", dbHomefront)
        Call FixGroupPhaseValue
    End If
    Screen.MousePointer = vbHourglass
    
    
    s = ""
    s = s & "insert into tblPhaseItem(DivisionID,Phase,Item,ItemNumber,CostCategory,PartNumber,Description,PhaseSortOrder,ItemSortOrder,orderuom,takeoffuom,ConversionFactor,WastePercent,RoundDir,Roundto)" & vbCrLf
    s = s & "values(" & HFApp.DivisionID & "," & DbQuote(Str, txtPhase.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtItem.Text & Left(Me.cboCostType.Text, 1)) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtItem.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, Left(Me.cboCostType.Text, 1)) & vbCrLf
    s = s & "      ," & DbQuote(Str, IIf(DontUseParts, "", txtItem.Text)) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtPhase.Text) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtItem.Text) & ",'EA','EA',1,0,0,0)"
    Call HFApp.SqlExec(s)
    
    
    
    Save_Add = True
    Screen.MousePointer = vbDefault
    Exit Function
    
eh: Select Case True
    Case InStr(1, Err.Description, "duplicate") > 0:
        Screen.MousePointer = vbDefault
        MsgBox "Precision Builder already has an item numbered " & Trim(txtPhase.Text) & "/" & Trim(txtItem.Text) & ". Please specify a different value.", vbInformation, App.ProductName
    Case Else:  Call errHandler(SRCFILE & "SaveData", s)
    End Select
End Function


Public Function Save_Renumber() As Boolean
On Error GoTo eh:


    Dim oldPhase As String
    Dim newPhase As String
    Dim oldItem As String
    Dim newItem As String
    Dim rs As Recordset
    Dim rc As Long
    Dim s As String
    Dim i As Long
    
    oldPhase = mPhase
    oldItem = mItemNumber
    newPhase = Trim(txtPhase.Text)
    newItem = Trim(txtItem.Text)
    If newPhase = "" Or newItem = "" Then
        MsgBox "Phase and item are both required", vbExclamation, App.ProductName
        Exit Function
    End If
    
    If HFApp.SqlExec("select * from tblEstPhases where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, txtPhase.Text, , True), dbHomefront).EOF Then
        If vbCancel = MsgBox("Phase " & Trim(txtPhase.Text) & " doesn't exist. Do you want to add it?", vbQuestion + vbOKCancel, App.ProductName) Then Exit Function
        Screen.MousePointer = vbHourglass
        Call HFApp.SqlExec("insert into tblestphases(DivisionID,phase,description,groupphase,sortorder) values(" & HFApp.DivisionID & "," & DbQuote(Str, txtPhase.Text, , True) & ",'untitled',0," & DbQuote(Num, txtPhase.Text, , True) & ")", dbHomefront)
                    
        Call FixGroupPhaseValue
    End If
    Screen.MousePointer = vbHourglass
    
    
    s = ""
    s = s & "UPDATE tblPhaseItem" & vbCrLf
    s = s & "   SET ItemNumber=" & DbQuote(Str, newItem) & vbCrLf
    s = s & "     , Item=" & DbQuote(Str, newItem) & "+CostCategory" & vbCrLf
    s = s & "     , Phase=" & DbQuote(Str, newPhase) & vbCrLf
    s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, oldPhase) & vbCrLf
    s = s & "   AND ItemNumber=" & DbQuote(Str, oldItem) & vbCrLf
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "update tblDBAssemblyDetails" & vbCrLf
    s = s & "   set Phase=" & DbQuote(Str, newPhase) & vbCrLf
    s = s & "      ,Item=" & DbQuote(Str, newItem) & "+right(item,1)" & vbCrLf
    s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, oldPhase) & vbCrLf
    s = s & "   AND Item like " & DbQuote(Str, oldItem & "_") & vbCrLf
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "update tblSalesSheetCosts" & vbCrLf
    s = s & "   set Phase=" & DbQuote(Str, newPhase) & vbCrLf
    s = s & "      ,Item=" & DbQuote(Str, newItem) & "+right(item,1)" & vbCrLf
    s = s & " WHERE Phase=" & DbQuote(Str, oldPhase) & vbCrLf
    s = s & "   AND Item like " & DbQuote(Str, oldItem & "_") & vbCrLf
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "update tblvendorcost" & vbCrLf
    s = s & "   set Phase=" & DbQuote(Str, newPhase) & vbCrLf
    s = s & "      ,Item=" & DbQuote(Str, newItem) & "+right(item,1)" & vbCrLf
    s = s & " WHERE Phase=" & DbQuote(Str, oldPhase) & vbCrLf
    s = s & "   AND Item like " & DbQuote(Str, oldItem & "_") & vbCrLf
    s = s & "   AND DivisionID = " & HFApp.DivisionID
    Call HFApp.SqlExec(s)
    
    
    Save_Renumber = True

EXITSUB:
    Screen.MousePointer = vbDefault
    Exit Function
    
eh: Select Case True
    Case InStr(1, Err.Description, "duplicate") > 0:
        Screen.MousePointer = vbDefault
        MsgBox "Precision Builder already has an item numbered " & newPhase & "/" & newItem & ". Please specify a different value.", vbInformation, App.ProductName
    Case Else:  Call errHandler(SRCFILE & "SaveData", s)
    End Select
End Function

Public Function Save_Duplicate() As Boolean
On Error GoTo eh:

    Dim oldPhase As String
    Dim newPhase As String
    Dim oldItem As String
    Dim newItem As String
    Dim newDesc As String
    Dim rs As Recordset
    Dim rc As Long
    Dim s As String
    Dim i As Long
    
    oldPhase = mPhase
    oldItem = newItem
    newPhase = Trim(txtPhase.Text)
    newItem = Trim(txtItem.Text)
    newDesc = Trim(txtDescription.Text)
    If newPhase = "" Or newItem = "" Or newDesc = "" Then
        MsgBox "Phase, item and description are all required", vbExclamation, App.ProductName
        Exit Function
    End If
    
    If HFApp.SqlExec("select * from tblEstPhases where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, oldPhase, , True), dbHomefront).EOF Then
        If vbCancel = MsgBox("Phase " & Trim(oldPhase) & " doesn't exist. Do you want to add it?", vbQuestion + vbOKCancel, App.ProductName) Then Exit Function
        Screen.MousePointer = vbHourglass
        Call HFApp.SqlExec("insert into tblestphases(DivisionID,phase,description,groupphase,sortorder) values(" & HFApp.DivisionID & "," & DbQuote(Str, oldPhase, , True) & ",'untitled',0," & DbQuote(Num, oldPhase, , True) & ")", dbHomefront)
        Call FixGroupPhaseValue
    End If
    Screen.MousePointer = vbHourglass

    
    
    s = ""
    s = s & "insert into tblphaseitem(DivisionID,Phase,Item,pricelink," & DbQuote(Str, newDesc) & ",notes,poindex,jccostcode,jccategory,orderuom,takeoffuom,conversionfactor,price,taxgroup,wastepercent,rounddir,roundto)" & vbCrLf
    s = s & "select (" & HFApp.DivisionID & "," & DbQuote(Str, newPhase) & "," & DbQuote(Str, newItem) & ",pricelink,description,notes,poindex,jccostcode,jccategory,orderuom,takeoffuom,conversionfactor,price,taxgroup,wastepercent,rounddir,roundto)" & vbCrLf
    s = s & "  from tblphaseitem" & vbCrLf
    s = s & " where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, oldPhase) & vbCrLf
    s = s & "   and item=" & DbQuote(Str, oldItem) & vbCrLf
    Call HFApp.SqlExec(s)
    
    

    
    
    On Error GoTo ehMB:
    If HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asMasterBuilder And IsNumeric(oldItem) And IsNumeric(oldPhase) Then
        s = HFApp.XmlMbStart(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
        s = s & "<PartAddRq requestID=""1"">" & vbCrLf
        s = s & HFApp.XmlMBAdd(n, 15, "ObjectRef", oldItem)
        s = s & HFApp.XmlMBAdd(c, 75, "Desc", txtDescription.Text)
        s = s & HFApp.XmlMBAdd(c, 10, "Unit", "ea")
        s = s & HFApp.XmlMBAdd(n, 10, "PartClassRef", oldPhase)
        s = s & "</PartAddRq>" & vbCrLf
        s = s & HFApp.XmlMBEnd()
        Call HFApp.XmlMbSubmit(s, HFApp.Options(MasterBuilderPWD))
    End If
    Save_Duplicate = True




EXITSUB:
    Screen.MousePointer = vbDefault
    Exit Function
    
eh: Select Case True
    Case InStr(1, Err.Description, "duplicate") > 0:
        Screen.MousePointer = vbDefault
        MsgBox "Precision Builder already has an item numbered " & Trim(oldPhase) & "/" & Trim(oldItem) & ". Please specify a different value.", vbInformation, App.ProductName
    Case Else:  Call errHandler(SRCFILE & "SaveData", s)
    End Select
    Exit Function
    
ehMB: Select Case True
    Case InStr(1, Err.Description, "duplicate") > 0:
        Screen.MousePointer = vbDefault
        MsgBox "Master Builder already has an item numbered " & Trim(oldItem) & ". This item could not be added to Master Builder.", vbInformation, App.ProductName
        Resume Next
    Case Else:  Call errHandler(SRCFILE & "SaveData", s)
    End Select
    Exit Function
    
ehTL: Select Case True
    Case InStr(1, Err.Description, "duplicate") > 0:
        Screen.MousePointer = vbDefault
        MsgBox "Timberline already has an item numbered " & Trim(oldPhase) & "/" & Trim(oldItem) & ". Please specify a different value.", vbInformation, App.ProductName
    Case Else:  Call errHandler(SRCFILE & "SaveData", s)
    End Select
    Exit Function
    
End Function

