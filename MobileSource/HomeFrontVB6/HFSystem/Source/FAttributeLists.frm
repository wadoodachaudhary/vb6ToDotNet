VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FAttributeLists 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option Attribute Lists"
   ClientHeight    =   7650
   ClientLeft      =   3135
   ClientTop       =   1950
   ClientWidth     =   11160
   Icon            =   "FAttributeLists.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11160
   Begin VB.CheckBox chkDeferred 
      Caption         =   "Deferred - Value selection does not have to happen immediately."
      Enabled         =   0   'False
      Height          =   315
      Left            =   4260
      TabIndex        =   3
      Top             =   975
      Width           =   7095
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   8760
      TabIndex        =   5
      Top             =   7170
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   9990
      TabIndex        =   4
      Top             =   7170
      Width           =   1095
   End
   Begin VB.CheckBox chkStrict 
      Caption         =   "Restricted - The agent must select from this list of values. No other values will be allowed."
      Enabled         =   0   'False
      Height          =   315
      Left            =   4260
      TabIndex        =   2
      Top             =   705
      Width           =   7095
   End
   Begin VB.CheckBox chkRequired 
      Caption         =   "Mandatory - Sales agent must enter a value when the option is selected."
      Enabled         =   0   'False
      Height          =   315
      Left            =   4260
      TabIndex        =   1
      Top             =   435
      Width           =   6465
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   6735
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      _cx             =   7011
      _cy             =   11880
      Appearance      =   1
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAttributeLists.frx":000C
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
   Begin VSFlex8Ctl.VSFlexGrid gValues 
      Height          =   5520
      Left            =   4200
      TabIndex        =   9
      Top             =   1575
      Width           =   6810
      _cx             =   12012
      _cy             =   9737
      Appearance      =   1
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAttributeLists.frx":00D6
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
      ExplorerBar     =   8
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
   Begin VB.Image cmdDelete 
      Height          =   240
      Left            =   3810
      Picture         =   "FAttributeLists.frx":019C
      Top             =   90
      Width           =   240
   End
   Begin VB.Image cmdNew 
      Height          =   240
      Left            =   3540
      Picture         =   "FAttributeLists.frx":0726
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "List Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   4170
      TabIndex        =   8
      Top             =   225
      Width           =   4365
   End
   Begin VB.Label Label2 
      Caption         =   "Attribute Lists"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   90
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "Attribute Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   4200
      TabIndex        =   6
      Top             =   1320
      Width           =   4365
   End
End
Attribute VB_Name = "FAttributeLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FAttributeLists::"

Private mInitialListID As Long
Private mDirty As Boolean
Private mLoading As Boolean 'stupid flags



Public Sub ShowForm(SelectedListID As Long)
    mInitialListID = SelectedListID
    Me.Show vbModal
End Sub

Private Sub chkDeferred_Click()
    Call SaveAttributeList
End Sub

Private Sub cmdDelete_Click()
    With gData
        If .Row > -1 Then
            .RowHidden(.Row) = True
            .Row = GridNextVisibleRow(gData, .Row)
            
            If .RowHidden(.Row) Then
                gValues.Enabled = False
                gValues.Rows = 1
                chkRequired.Enabled = False
                chkStrict.Enabled = False
            End If
            
            mDirty = True
            
        End If
    End With
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0:  If SaveData(False) Then Unload Me
        Case 1:  Unload Me
    End Select
End Sub

Private Sub cmdNew_Click()
    With gData
        .AddItem ""
        .TextMatrix(.Rows - 1, .ColIndex("name")) = "New Attribute List"
        .TextMatrix(.Rows - 1, .ColIndex("required")) = "False"
        .TextMatrix(.Rows - 1, .ColIndex("strict")) = "False"
        .TextMatrix(.Rows - 1, .ColIndex("dirty")) = "True"
        mDirty = True
        .SetFocus
        Call .Select(.Rows - 1, .ColIndex("name"))
        Call .EditCell
    End With
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadData
End Sub

Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim listid As Long
    Dim selectedrow As Long
    
    Dim cl As String
    
    s = ""
    s = s & "select l.listid,l.name,l.required,l.strictlist strict,isnull(l.deferred,0) deferred,v.valueid,v.value,v.upcharge,v.otherlistid" & vbCrLf
    s = s & "from attributelists l" & vbCrLf
    s = s & "left outer join attributelistvalues v on l.listid=v.listid" & vbCrLf
    s = s & "where IsNull(l.inactive, 0) = 0" & vbCrLf
    s = s & "and IsNull(v.inactive, 0) = 0" & vbCrLf
    s = s & "order by l.name,l.listid,v.sortorder" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    listid = -999
    With gData
        .Rows = 0
        cl = "|#0;"
        While Not rs.EOF
        
            If listid = Val("" & rs("listid")) Then
                listid = Val("" & rs("listid"))
                If listid = mInitialListID Then selectedrow = .Rows - 1
            Else
                .AddItem ""
                listid = Val("" & rs("listid"))
                cl = cl & "|#" & rs("listid") & ";" & rs("name")
                .TextMatrix(.Rows - 1, .ColIndex("listid")) = "" & rs("listid")
                .TextMatrix(.Rows - 1, .ColIndex("name")) = "" & rs("name")
                .TextMatrix(.Rows - 1, .ColIndex("required")) = "" & rs("required")
                .TextMatrix(.Rows - 1, .ColIndex("strict")) = "" & rs("strict")
                .TextMatrix(.Rows - 1, .ColIndex("deferred")) = "" & rs("deferred")
            End If
            'If Trim("" & rs("value")) <> "" Then
                .TextMatrix(.Rows - 1, .ColIndex("list")) = .TextMatrix(.Rows - 1, .ColIndex("list")) & Trim("" & rs("valueid")) & vbTab & Trim("" & rs("value")) & vbTab & Trim("" & rs("upcharge")) & vbTab & Trim("" & rs("Otherlistid")) & vbCr
            'End If
        
            rs.MoveNext
        Wend
        gValues.ColComboList(4) = cl
        On Error Resume Next
        Call .Select(selectedrow, .ColIndex("name"))
    End With
    mDirty = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim section As String
    Dim s As String
    
    Dim ValueID As Long
    Dim Description As String
    Dim UpCharge As String
    Dim OtherListID As String
    
    Dim r As Long
    Dim lov As String
    Dim i As Long
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    
    If prompt Then
        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?", vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    Screen.MousePointer = vbHourglass
    
    With gData
        For r = .Rows - 1 To 0 Step -1
            If .RowHidden(r) Then
                s = "update attributelists set inactive=1 where listid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ListID")))
                HFApp.SqlExec s
                Call .RemoveItem(r)
            End If
        Next
        
        For r = 0 To .Rows - 1
            If .TextMatrix(r, .ColIndex("Dirty")) = "True" Then
                
                section = "List"
                If .ValueMatrix(r, .ColIndex("listid")) = 0 Then
                    s = ""
                    s = s & "insert into AttributeLists(name,required,strictlist,deferred)" & vbCrLf
                    s = s & "values(" & DbQuote(Str, .TextMatrix(r, .ColIndex("Name"))) & vbCrLf
                    s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Required"))) & vbCrLf
                    s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Strict"))) & vbCrLf
                    s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("deferred"))) & vbCrLf
                    s = s & ")"
                    HFApp.SqlExec s
                    .TextMatrix(r, .ColIndex("ListID")) = HFApp.SqlIdentity("AttributeLists")
                Else
                    s = ""
                    s = s & "update AttributeLists" & vbCrLf
                    s = s & "set name=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Name"))) & vbCrLf
                    s = s & "   ,required=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("Required"))) & vbCrLf
                    s = s & "   ,strictlist=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("Strict"))) & vbCrLf
                    s = s & "   ,deferred=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("deferred"))) & vbCrLf
                    s = s & "where listid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ListID"))) & vbCrLf
                    HFApp.SqlExec s
                End If
                
                section = "Values"
                s = "update attributelistvalues set inactive=1 where listid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ListID")))
                HFApp.SqlExec s
                lov = .TextMatrix(r, .ColIndex("List"))
                For i = 1 To Parse(lov, , vbCr)
                    s = Trim(Parse(lov, i, vbCr))
                    
                    ValueID = Val(Parse(s, 1, vbTab))
                    Description = Parse(s, 2, vbTab)
                    UpCharge = Parse(s, 3, vbTab)
                    OtherListID = Parse(s, 4, vbTab)
                    
                    If Description <> "" Or UpCharge <> "" Then
                        If ValueID = 0 Then
                            s = ""
                            s = s & "insert into attributelistvalues(listid,value,upcharge,otherlistid,sortorder)" & vbCrLf
                            s = s & "values(" & DbQuote(Num, .TextMatrix(r, .ColIndex("ListID"))) & vbCrLf
                            s = s & "      ," & DbQuote(Str, Description) & vbCrLf
                            s = s & "      ," & DbQuote(Num, UpCharge) & vbCrLf
                            s = s & "      ," & DbQuote(Num, OtherListID) & vbCrLf
                            s = s & "      ," & DbQuote(Num, i) & ")"
                        Else
                            s = ""
                            s = s & "update attributelistvalues set" & vbCrLf
                            s = s & " inactive=0" & vbCrLf
                            s = s & ",value=" & DbQuote(Str, Description) & vbCrLf
                            s = s & ",upcharge=" & DbQuote(Num, UpCharge) & vbCrLf
                            s = s & ",otherlistid=" & DbQuote(Num, OtherListID) & vbCrLf
                            s = s & ",sortorder=" & DbQuote(Num, i) & vbCrLf
                            s = s & "where valueid=" & DbQuote(Num, ValueID) & vbCrLf
                        End If
                        HFApp.SqlExec s
                    End If
                Next
                
                .TextMatrix(r, .ColIndex("Dirty")) = ""
            End If
        Next
    End With
       
    SaveData = True
    mDirty = False
    Screen.MousePointer = vbDefault
    
Exit Function
eh:
If Err.Description Like "*duplicate*" Then
    If section = "list" Then
        Screen.MousePointer = vbDefault
        Call gData.Select(r, gData.ColIndex("name"))
        MsgBox "You already have a list named """ & gData.Text & """. Please use a different name.", vbInformation, App.ProductName
        Call gData.EditCell
    Else
        Resume Next
    End If
Else
    Call errHandler(SRCFILE & "SaveData", s)
End If
End Function




Private Sub chkRequired_Click()
    Call SaveAttributeList
End Sub
Private Sub chkStrict_Click()
    Call SaveAttributeList
End Sub


Private Sub gValues_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call SaveAttributeList
End Sub

Private Sub gValues_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    Dim r1 As Long
    Dim r2 As Long
    Dim i As Long
    
    
    With gValues
        
        Select Case True
        'delete rows  CTRL+X or CTRL+DEL
        Case (Shift = vbCtrlMask And KeyCode = vbKeyX) Or (Shift = vbCtrlMask And KeyCode = vbKeyDelete)
            r1 = Min(.Row, .RowSel)
            r2 = Max(.Row, .RowSel)
            For i = r1 To r2 Step -1
                If i > 0 And i < .Rows - 1 Then .RemoveItem i
            Next
        
    
        'delete cells
        Case KeyCode = vbKeyDelete
            .Clip = ""
        
    
        'copy
        Case Shift = vbCtrlMask And KeyCode = vbKeyC
            s = .Clip
            
            'trim trailing blank row and ensure consistent vbCrlf
            s = Replace(s, vbCrLf, vbCr)
            s = Replace(s, vbLf, vbCr)
            If Right(s, 1) = vbCr Then s = Mid(s, 1, Len(s) - 1)
            If Right(s, 1) = vbTab Then s = Mid(s, 1, Len(s) - 1)
            If Right(s, 1) = vbCr Then s = Mid(s, 1, Len(s) - 1)
            s = Replace(s, vbCr, vbCrLf)
            
            Clipboard.Clear
            Clipboard.SetText s
        
        'paste
        Case Shift = vbCtrlMask And KeyCode = vbKeyV
            s = Clipboard.GetText
            
            'ensure consistant use of vbCr
            s = Replace(s, vbCrLf, vbCr)
            s = Replace(s, vbLf, vbCr)
            
            i = Parse(s, , vbCr)
            .Rows = Max(.Rows, .Row + i)
            .RowSel = .Row + i - 1
            If .Col = 1 Then .ColSel = 2
            .Clip = s
            If .RowSel = .Rows - 1 Then .AddItem ""
            .Row = .Row
        
        End Select
            
    End With
    Call gValues_AfterEdit(gValues.Row, gValues.Col)
    
End Sub

Private Sub gValues_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gValues
    Select Case .ColKey(Col)
    Case "upcharge"
        .EditText = Val(.EditText)
    Case Else
    End Select
    If Row = .Rows - 1 Then .AddItem ""
    End With
End Sub





Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF And Shift <> 0 Then
        Call FFind.ShowForm(gData, False)
    End If
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mDirty = True
    gData.TextMatrix(Row, gData.ColIndex("Dirty")) = "True"
End Sub

Private Sub gData_RowColChange()
    Call LoadAttributeList
End Sub


Private Sub LoadAttributeList()

    Dim s As String
    Dim i As Long

    mLoading = True
    With gData
    If .Row > -1 Then
        s = .TextMatrix(.Row, .ColIndex("list"))
        i = Parse(s, , vbCr)
        
        gValues.Rows = 1
        gValues.Rows = i + 1
        gValues.Cell(flexcpText, 1, 1, Max(1, i - 1), 4) = s
        
        chkRequired.Value = IIf(.TextMatrix(.Row, .ColIndex("required")) = "True", vbChecked, vbUnchecked)
        chkStrict.Value = IIf(.TextMatrix(.Row, .ColIndex("strict")) = "True", vbChecked, vbUnchecked)
        chkDeferred.Value = IIf(.TextMatrix(.Row, .ColIndex("deferred")) = "True", vbChecked, vbUnchecked)
    End If
    gValues.Enabled = .Row > -1
    chkRequired.Enabled = .Row > -1
    chkStrict.Enabled = .Row > -1
    chkDeferred.Enabled = .Row > -1
    End With
    mLoading = False
End Sub

Private Sub SaveAttributeList()
    Dim s As String
    With gData
    If .Row > -1 And Not mLoading Then
    
        s = gValues.Cell(flexcpText, 1, 1, gValues.Rows - 2, 3) & vbCr
        .TextMatrix(.Row, .ColIndex("list")) = s
        
        .TextMatrix(.Row, .ColIndex("required")) = chkRequired.Value = vbChecked
        .TextMatrix(.Row, .ColIndex("strict")) = chkStrict.Value = vbChecked
        .TextMatrix(.Row, .ColIndex("deferred")) = chkDeferred.Value = vbChecked
        .TextMatrix(.Row, .ColIndex("dirty")) = "True"
        mDirty = True
    End If
    End With
End Sub
