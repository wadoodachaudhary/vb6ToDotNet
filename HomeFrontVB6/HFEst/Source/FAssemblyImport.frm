VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FAssemblyImport 
   Caption         =   "Data Import Wizard"
   ClientHeight    =   3480
   ClientLeft      =   7740
   ClientTop       =   2220
   ClientWidth     =   5610
   ClipControls    =   0   'False
   Icon            =   "FAssemblyImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   5610
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1575
      Left            =   1125
      TabIndex        =   0
      Top             =   915
      Width           =   3210
      _cx             =   5662
      _cy             =   2778
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FAssemblyImport.frx":000C
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "FAssemblyImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FAssemblyImport::"

Private mSession        As String
Private mFileType As String       'eg: models, options, globals

Const mColumnMappings = "Option=OptionID"


Private Sub Form_Resize()
    gData.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub ImportAssemblies()
On Error GoTo eh
    
    
    Dim FileName As String
    Dim rs As Recordset
    Dim s As String
    Dim i As Long
    
    
    'choose file
    If Not VBGetOpenFileName(FileName, , , , , True, "Assembly Import Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv", , , , , FMain.hwnd) Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    'load data file into grid
    mSession = CreateGUID()
    Select Case FileExt(FileName)
        Case "xls", "xlsx"
            FileName = SaveCSV(FileName)
            Call LoadCSV(FileName, True)
            On Error Resume Next
            Kill FileName
            On Error GoTo eh
        Case Else
            Call LoadCSV(FileName, True)
    End Select
    
    
    With gData
    
        'change columnheading and columnkey from communitycode to community
        If .ColIndex("CommunityCode") <> -1 And .ColIndex("Community") = -1 Then
            .TextMatrix(0, .ColIndex("CommunityCode")) = "Community"
            .ColKey(.ColIndex("CommunityCode")) = "Community"
        End If
    
        'check filetype valid format
        'models  = community,model,assembly         -- assembly and model, no option
        'options = community,model,option,assembly  -- assembly and model and option
        'globals = community,option,assembly        -- assembly and option no model
        Select Case True
            Case .ColIndex("model") <> -1 And .ColIndex("optionid") = -1:  mFileType = "models"
            Case .ColIndex("model") <> -1 And .ColIndex("optionid") <> -1: mFileType = "options"
            Case .ColIndex("model") = -1 And .ColIndex("optionid") <> -1:  mFileType = "globals"
            Case Else:
                Screen.MousePointer = vbDefault
                Unload FProgress
                DoEvents
                MsgBox "Invalid file format.", vbExclamation, App.ProductName
                Exit Sub
        End Select
                  
        'insert data into tmptable
Call FProgress.Progress("Validating...", " ", 0, 0, FMain)
        Call SaveAssemblyData
        
        'validate data
        s = "exec Purch_ImportedAssemblies_Validate " & DbQuote(Num, HFApp.DivisionID) & ", " & DbQuote(Str, mSession) & ", " & DbQuote(Str, mFileType)
        Set rs = HFApp.SqlExec(s)
        If rs.EOF Then
            Screen.MousePointer = vbDefault
Call FProgress.Progress("Saving...", " ", 0, 0, FMain)
            s = "exec Purch_ImportedAssemblies_Commit " & DbQuote(Num, HFApp.DivisionID) & ", " & DbQuote(Str, mSession) & ", " & DbQuote(Str, mFileType)
            Call HFApp.SqlExec(s)
            MsgBox "Import Completed!", vbInformation, "Assembly Import"
            HFApp.SqlExec "delete importedassemblies where session=" & DbQuote(Str, mSession)
        Else
            s = ""
            i = 0
            While Not rs.EOF
                i = i + 1
                If i < 25 Then s = s & rs(0) & "        " & rs(1) & vbCrLf
                rs.MoveNext
            Wend
            
            If rs.EOF Then
                s = "This file cannot be imported." & vbCrLf & vbCrLf & "row   error" & vbCrLf & s
            Else
                s = "This file cannot be imported." & vbCrLf & i & " errors were found." & vbCrLf & vbCrLf & "row   error" & vbCrLf & s
            End If
            Screen.MousePointer = vbDefault
            Unload FProgress
            MsgBox s, vbExclamation, "Errors in import file"
            
        End If
        
        
    End With
    
    Unload FProgress
    Unload Me
    

    
        
        
Exit Sub
eh: Call errHandler(SRCFILE & "ImportAssemblies")
    HFApp.SqlExec "delete importedassemblies where session=" & DbQuote(Str, mSession)
    Screen.MousePointer = vbDefault
    Unload FProgress
End Sub




Public Sub ImportAssemblyTakeoffs()
On Error GoTo eh
    
    Dim FileName As String
    Dim rs As Recordset
    Dim s As String
    Dim i As Long
        
    'choose file
    If Not VBGetOpenFileName(FileName, , , , , True, "Assembly Import Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv", , , , , FMain.hwnd) Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    'load data file into grid
    mSession = CreateGUID()
    Select Case FileExt(FileName)
        Case "xls", "xlsx"
            FileName = SaveCSV(FileName)
            Call LoadCSV(FileName, False)
            On Error Resume Next
            Kill FileName
            On Error GoTo eh
        Case Else
            Call LoadCSV(FileName, False)
    End Select
    
           
    With gData
                  
        'insert data into tmptable
 Call FProgress.Progress("Writing...", " ", 0, 0, FMain)
        On Error Resume Next
        Call SaveAssemblyTakeoffData
        s = Err.Description
        On Error GoTo eh
        If s <> "" Then
            Screen.MousePointer = vbDefault
            Unload FProgress
            MsgBox "Invalid file format.", vbExclamation, App.ProductName
            Exit Sub
        End If
        
        
        'validate data
Call FProgress.Progress("Validating...", " ", 0, 0, FMain)
        s = "exec Purch_ImportedAssemblyTakeoffs_Validate " & DbQuote(Num, HFApp.DivisionID) & ", " & DbQuote(Str, mSession)
        Set rs = HFApp.SqlExec(s)
        If rs.EOF Then
            Screen.MousePointer = vbDefault
Call FProgress.Progress("Saving...", " ", 0, 0, FMain)
            s = "exec Purch_ImportedAssemblyTakeoffs_Commit " & DbQuote(Num, HFApp.DivisionID) & ", " & DbQuote(Str, mSession)
            Call HFApp.SqlExec(s)
            MsgBox "Import Completed!", vbInformation, "Assembly Takeoffs Import"
        Else
            s = ""
            i = 0
            While Not rs.EOF
                i = i + 1
                If i < 25 Then s = s & rs(1) & vbCrLf
                rs.MoveNext
            Wend
            
            If rs.EOF Then
                s = "This file cannot be imported." & vbCrLf & vbCrLf & s
            Else
                s = "This file cannot be imported." & vbCrLf & i & " errors were found." & vbCrLf & vbCrLf & s
            End If
            Screen.MousePointer = vbDefault
            MsgBox s, vbExclamation, "Errors in import file"
            HFApp.SqlExec "delete importedassemblytakeoffs where session=" & DbQuote(Str, mSession)
            HFApp.SqlExec "delete importedassemblytakeoffitems where session=" & DbQuote(Str, mSession)

        End If
        
        
    End With
    
    Unload FProgress
    Unload Me
    

    
        
        
Exit Sub
eh: Call errHandler(SRCFILE & "ImportAssemblyTakeoffs", s)
    HFApp.SqlExec "delete importedassemblytakeoffs where session=" & DbQuote(Str, mSession)
    HFApp.SqlExec "delete importedassemblytakeoffitems where session=" & DbQuote(Str, mSession)
    Screen.MousePointer = vbDefault
    Unload FProgress
End Sub




Private Function SaveCSV(FileName As String) As String
    Dim s As String
    
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSH As Object
    
Call FProgress.Progress("Initializating...", "loading Excel", 0, 0, FMain)
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Open(FileName, , True)
    Set xlSH = xlWB.Worksheets(1)
        
    s = TempFile("csv")
    Call xlSH.SaveAs(s, 6, , , , , False)
    
    Set xlSH = Nothing
    Call xlWB.Close
    Set xlWB = Nothing
    Set xlApp = Nothing
    
    SaveCSV = s
End Function

Private Sub LoadCSV(FileName As String, HasColHeadings As Boolean)
On Error GoTo eh

    Dim i As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    Dim c1 As String
    Dim c2 As String
    
    Dim rowkey As String
    Dim missingkey As Boolean
    
    
Call FProgress.Progress("Initializating...", "reading data", 0, 0, FMain)
    
    Select Case FileExt(FileName)
        Case "csv": Call gData.LoadGrid(FileName, flexFileCommaText)
        Case Else:  Call gData.LoadGrid(FileName, flexFileTabText)
    End Select
    
    With gData
    
        If HasColHeadings Then
            'set column indexes
            For i = 1 To .Cols - 1
                .ColKey(i) = Replace(.TextMatrix(1, i), " ", "")
            Next
            Call .RemoveItem(0)
            .FixedRows = 1
            
            'translate column keys
            If mColumnMappings <> "" Then
                On Error Resume Next
                For i = 1 To Parse(mColumnMappings)
                    c1 = Trim(Parse(Parse(mColumnMappings, i), 1, "="))
                    c2 = Trim(Parse(Parse(mColumnMappings, i), 2, "="))
                    
                    If .ColIndex(c1) <> -1 Then
                        If c1 <> "" And c2 <> "" Then
                            .TextMatrix(0, .ColIndex(c1)) = c2
                            .ColKey(.ColIndex(c1)) = c2
                        End If
                        If c2 = "" Then
                            .ColHidden(.ColIndex(c1)) = True
                            .ColKey(.ColIndex(c1)) = ""
                        End If
                    End If
                Next
                On Error GoTo eh
            End If
        Else
            'write excel style column names to fixed cols
            For i = 1 To .Cols - 1
                .TextMatrix(0, i) = FormatExcelColRef(i)
            Next
            
        End If
        
        Call .AutoSize(0, .Cols - 1)
        
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "LoadCSV", s)
    Screen.MousePointer = vbDefault
    Unload FProgress
End Sub





Private Function SaveAssemblyData() As Boolean
On Error GoTo eh

    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim columns As String
    Dim values As String

    With gData
    
        columns = "session,row"
        For c = 1 To .Cols - 1
            If .ColKey(c) <> "" Then
                columns = columns & "," & .ColKey(c)
            End If
        Next
        
        values = ""
        For r = 1 To .Rows - 1
        
            'skip blank rows
            Call .Select(r, 0, r, .Cols - 1)
            s = Replace(Replace(.clip, vbTab, ""), " ", "")
            If s <> "" Then
                s = "'" & mSession & "'," & r + 1 & vbCrLf
                For c = 1 To .Cols - 1
                    If .ColKey(c) <> "" Then
                        s = s & "," & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
                    End If
                Next
                values = values & ",(" & s & ")" & vbCrLf
            End If
        Next
        values = Mid(values, 2)
        
        s = "insert ImportedAssemblies(" & columns & ") values " & vbCrLf & values
        Call HFApp.SqlExec(s, dbHomefront)
    End With

Exit Function
eh: Call errHandler(SRCFILE & "SaveAssemblyData", s)
    Screen.MousePointer = vbDefault
    Unload FProgress
End Function






Private Function SaveAssemblyTakeoffData() As Boolean

    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim values As String
    
    With gData
    
        'write headers to ImportedAssemblyTakeoffs
 Call FProgress.Progress("", "headers...", 0, 0, FMain)
        values = ""
        For c = 6 To .Cols - 1
            'skip blank assemblies
            Call .Select(1, c, 12, c)
            s = Replace(.clip, " ", "")
            s = Replace(s, vbCr, "")
            s = Replace(s, vbLf, "")
            If s <> "" Then
                'build values string
                s = ",'" & mSession & "','" & .TextMatrix(0, c) & "'" & vbCrLf
                s = s & "," & DbQuote(Str, Parse(.TextMatrix(1, c), 1, " - ")) & vbCrLf 'community
                s = s & "," & DbQuote(Str, .TextMatrix(2, c)) & vbCrLf 'model
                s = s & "," & DbQuote(Str, .TextMatrix(3, c)) & vbCrLf 'option
                s = s & "," & DbQuote(Str, .TextMatrix(4, c)) & vbCrLf 'assembly
                s = s & "," & DbQuote(Str, .TextMatrix(5, c)) & vbCrLf 'description
                s = s & "," & DbQuote(Str, .TextMatrix(6, c)) & vbCrLf 'comments
                s = s & "," & DbQuote(Str, Parse(.TextMatrix(7, c), 1, " - ")) & vbCrLf 'series
                s = s & "," & DbQuote(Str, .TextMatrix(8, c)) & vbCrLf 'scheduletemplate
                s = s & "," & DbQuote(Str, .TextMatrix(9, c)) & vbCrLf 'notes
                s = s & "," & DbQuote(Str, Parse(.TextMatrix(10, c), 1, " - ")) & vbCrLf 'category
                s = s & "," & DbQuote(Str, Parse(.TextMatrix(11, c), 1, " - ")) & vbCrLf 'constcutoff
                s = s & "," & DbQuote(Str, .TextMatrix(12, c)) & vbCrLf 'uom
                values = values & ",(" & Mid(s, 2) & ")"
            End If
        Next
        values = Mid(values, 2)
        s = "insert ImportedAssemblyTakeoffs(session,col,Community,Model,OptionID,Assembly,Description,Comments,Series,ScheduleTemplate,EstimatorNotes,SubCat,ConstCutoff,UOM) values" & vbCrLf & values
        Call HFApp.SqlExec(s, dbHomefront)
                    
        'unformat poindex/phase/item
        For r = 14 To .Rows - 1
            'strip description from poindex in column A:   format -> POIndex - description
            .TextMatrix(r, 1) = Parse(.TextMatrix(r, 1), 1, " - ")
            'parse phase and item from column B:   format -> phase\item - description (uom)
            s = Parse(.TextMatrix(r, 2), 1, " - ")
            .Cell(flexcpData, r, 2) = Parse(s, 1, "\") 'phase
            .Cell(flexcpText, r, 2) = Parse(s, 2, "\") 'item
        Next
        

        'write quantities to ImportedAssemblyTakeoffs
        For c = 6 To .Cols - 1
            
 Call FProgress.Progress("", "quantities...", c - 5, .Cols - 5, FMain)
            'skip blank assemblies
            Call .Select(1, c, 12, c)
            s = Replace(Replace(.clip, vbTab, ""), " ", "")
            If s <> "" Then
            
                'build values string
                values = ""
                For r = 14 To .Rows - 1
                    'skip blank quantities
                    If Trim(.TextMatrix(r, c)) <> "" Then
                        s = "'" & mSession & "'," & r & ",'" & .TextMatrix(0, c) & "'"
                        s = s & "," & DbQuote(Str, Parse(.TextMatrix(1, c), 1, " - ")) & vbCrLf 'community
                        s = s & "," & DbQuote(Str, .TextMatrix(2, c))  'model
                        s = s & "," & DbQuote(Str, .TextMatrix(3, c))  'option
                        s = s & "," & DbQuote(Str, .TextMatrix(4, c))  'assembly
                        s = s & "," & DbQuote(Str, .TextMatrix(r, 1))   'poindex
                        s = s & "," & DbQuote(Str, .Cell(flexcpData, r, 2))  'phase
                        s = s & "," & DbQuote(Str, .TextMatrix(r, 2))  'item
                        s = s & "," & DbQuote(Str, .TextMatrix(r, 3))  'invertable
                        s = s & "," & DbQuote(Str, .TextMatrix(r, 4))  'location
                        s = s & "," & DbQuote(Str, .TextMatrix(r, 5))  'notes
                        s = s & "," & DbQuote(Str, .TextMatrix(r, c))  'qty
                        values = values & ",(" & s & ")" & vbCrLf
                    End If
                Next
                values = Mid(values, 2)
                
                'write one assembly
                If values <> "" Then
                    s = "insert ImportedAssemblyTakeoffItems(session,row,col,   Community,Model,OptionID,Assembly,   POIndex,Phase,Item,Invertable,Location,Notes,Qty) values " & vbCrLf & values
                    Call HFApp.SqlExec(s, dbHomefront)
                End If
                
            End If
        Next
    End With

End Function





