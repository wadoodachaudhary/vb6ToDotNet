VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FBIMImport 
   Caption         =   "Import BIM Assemblies"
   ClientHeight    =   8205
   ClientLeft      =   3315
   ClientTop       =   975
   ClientWidth     =   11130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FBIMImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11130
   Begin VB.Frame frmFinish 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4305
      Left            =   450
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   7440
      Begin VB.Image Image1 
         Height          =   480
         Left            =   375
         Picture         =   "FBIMImport.frx":000C
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblReady 
         Caption         =   "Ready? "
         Height          =   195
         Left            =   1365
         TabIndex        =   6
         Top             =   555
         Width           =   600
      End
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   1588
      Caption         =   "Import Assemblies from BIM Pipeline"
      Description     =   "These are the assemblies you are importing. Please verify that they are correct."
      Icon            =   "FBIMImport.frx":08D6
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11130
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7620
      Width           =   11130
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   1
         Left            =   2970
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1860
      Left            =   105
      TabIndex        =   3
      Top             =   960
      Width           =   4845
      _cx             =   8546
      _cy             =   3281
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FBIMImport.frx":11B0
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   1
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
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Remove"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FBIMImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FBIMImport::"

Dim mSessionCode As String
Dim mCancel As Boolean
Dim mGridMode As String ' which step are we on =  models, options, buildingphases, items
Dim mIsModelFile As Boolean ' have imported a modelsmaster file as opposed to an optionsmaster+itemdetails file

Dim mEditItemDB As Boolean ' user permission - can edit?

Public Function ShowForm() As Boolean
On Error GoTo eh:
    
    Dim FileName As String
    Dim filter As String
    Dim r As Long
    Dim s As String
    
    Dim done As Boolean
    
    
    If Not HFApp.UserPermission("editassemblies") Then
        MsgBox "You don't have permissions to update assemblies. Contact an administrator to change your permissions.", vbInformation, App.ProductName
        Exit Function
    End If
    mEditItemDB = HFApp.UserPermission("edititemdb")
    
    mSessionCode = CreateGUID()
    
    'get and load file into staging area
    filter = "Excel Files (*.csv;*.xls;*.xlsx)|*.csv;*.xls;*.xlsx"
    If Not VBGetOpenFileName(FileName, , , , , True, filter, , , "BIM Pipeline Import", , Screen.ActiveForm.hwnd) Then Exit Function
    If FileExt(FileName) = "csv" Then
        s = FileName
    Else
        s = SaveToCSV(FileName, 2)
    End If
    Screen.MousePointer = vbHourglass
    Call LoadFile(s)
    Screen.MousePointer = vbDefault
    On Error Resume Next
    If FileExt(FileName) <> "csv" Then Kill s
    On Error GoTo eh
    
    With gData
    
        'Check for size problems
        WizHead1.Description = "This file contains data tha exceeds HomeFront field sizes. Please shrink the data before importing."
        s = "select SizeError from dbo.BIMImportSizeErrors where sessioncode=" & DbQuote(Str, mSessionCode)
        Call LoadSQL(s)
        If gData.Rows > 1 Then
            cmdNav(0).Visible = False
            Me.Show vbModal
            GoTo ExitFunction
        End If
        
    
        If Not mIsModelFile Then
            'show/edit included dependent conditions
            WizHead1.Description = "Which BIM Pipeline conditions would you like to import into HomeFront?"
            s = ""
            s = s & "Select distinct d.DependentCondition,cast(case when c.condition is null then 0 else 1 end as bit) Include" & vbCrLf
            s = s & "from BIMImportDetails d" & vbCrLf
            s = s & "left outer join BIMImportIncludedConditions c on d.DependentCondition=c.Condition" & vbCrLf
            s = s & "where sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
            s = s & "order by Include desc, DependentCondition" & vbCrLf
            Call LoadSQL(s)
            If gData.Rows > 1 Then
                mGridMode = "Conditions"
                Call IniGetGrid(Me, gData, , , mGridMode)
                Me.Show vbModal
                Call IniPutGrid(Me, gData, , mGridMode)
                If mCancel Then GoTo ExitFunction
                'save changes
                For r = 1 To .Rows - 1
                If .Cell(flexcpData, r, 0) = "dirty" Then
                    If .Cell(flexcpChecked, r, .ColIndex("include")) = flexChecked Then
                        s = "insert into BIMImportIncludedConditions(Condition) values(" & DbQuote(Str, .TextMatrix(r, .ColIndex("DependentCondition"))) & ")"
                    Else
                        s = "delete BIMImportIncludedConditions where Condition=" & DbQuote(Str, .TextMatrix(r, .ColIndex("DependentCondition")))
                    End If
                    On Error Resume Next
                    Call HFApp.SqlExec(s)
                    On Error GoTo eh
                End If
                Next
            End If
        End If
        
        
        'show/edit assembly headers
        WizHead1.Description = "Step 1 - These are the assemblies you are about to import. Please verify that they are correct."
        s = ""
        s = s & "select Model,OptionID,Description,Series,FloorArea,Bedrooms,Bathrooms,Style,AssemblyType" & vbCrLf
        s = s & "from BIMImportAssemblyMasters" & vbCrLf
        s = s & "where sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
        s = s & "order by model,optionid"
        Call LoadSQL(s)
        mGridMode = IIf(mIsModelFile, "Models", "Options")
        Call IniGetGrid(Me, gData, , , mGridMode)
        .ColHidden(.ColIndex("OptionID")) = mIsModelFile
        .ColHidden(.ColIndex("AssemblyType")) = True
        Me.Show vbModal
        Call IniPutGrid(Me, gData, , mGridMode)
        If mCancel Then GoTo ExitFunction
        'save changes
        For r = 1 To .Rows - 1
        If .Cell(flexcpData, r, 0) = "dirty" Then
            If .TextMatrix(r, .ColIndex("AssemblyType")) = "0" Then
                s = ""
                s = s & "update BIMImportModels" & vbCrLf
                s = s & "set description=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                s = s & "   ,style=" & DbQuote(Str, .TextMatrix(r, .ColIndex("style"))) & vbCrLf
                s = s & "   ,series=" & DbQuote(Str, .TextMatrix(r, .ColIndex("series"))) & vbCrLf
                s = s & "   ,totalsqf=" & DbQuote(Num, .TextMatrix(r, .ColIndex("floorarea"))) & vbCrLf
                s = s & "   ,bedrooms=" & DbQuote(Num, .TextMatrix(r, .ColIndex("bedrooms"))) & vbCrLf
                s = s & "   ,bathrooms=" & DbQuote(Num, .TextMatrix(r, .ColIndex("bathrooms"))) & vbCrLf
                s = s & "where sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
                s = s & "and plannumber=" & DbQuote(Str, .TextMatrix(r, .ColIndex("model"))) & vbCrLf
            Else
                s = ""
                s = s & "update BIMImportDetails" & vbCrLf
                s = s & "set optiondescription=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                s = s & "where sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
                s = s & "and plannumber=" & DbQuote(Str, .TextMatrix(r, .ColIndex("model"))) & vbCrLf
                s = s & "and optionnumber=" & DbQuote(Str, .TextMatrix(r, .ColIndex("optionid"))) & vbCrLf
            End If
            Call HFApp.SqlExec(s)
        End If
        Next
    
        'show/edit ignorable building phases
        WizHead1.Description = "Step 2 - Mark any BIM Pipeline building phases that you don't want to import into HomeFront"
        s = ""
        s = s & "Select distinct BuildingPhase,IgnoreBuildingPhase Ignore" & vbCrLf
        s = s & "from BIMImportAssemblyDetails" & vbCrLf
        s = s & "where sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
        s = s & "order by IgnoreBuildingPhase desc, BuildingPhase"
        Call LoadSQL(s)
        If gData.Rows > 1 Then
            mGridMode = "BuildingPhase"
            Call IniGetGrid(Me, gData, , , mGridMode)
            Me.Show vbModal
            Call IniPutGrid(Me, gData, , mGridMode)
            If mCancel Then GoTo ExitFunction
            'save changes
            For r = 1 To .Rows - 1
            If .Cell(flexcpData, r, 0) = "dirty" Then
                If .Cell(flexcpChecked, r, .ColIndex("ignore")) = flexChecked Then
                    s = "insert into BIMImortIgnoreBuildingPhases(BuildingPhase) values(" & DbQuote(Str, .TextMatrix(r, .ColIndex("BuildingPhase"))) & ")"
                Else
                    s = "delete BIMImortIgnoreBuildingPhases where BuildingPhase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("BuildingPhase")))
                End If
                On Error Resume Next
                Call HFApp.SqlExec(s)
                On Error GoTo eh
            End If
            Next
        End If
        
        
        'show/edit unknown items
        If mEditItemDB Then
        WizHead1.Description = "Step 3 - These items don't exist in your HomeFront item database. Skip them, create new items, or map to an existing item?"
        Else
        WizHead1.Description = "Step 3 - These items don't exist in your HomeFront item database. They cannot be imported. You must correct the import file or skip these items."
        End If
        done = False
        Do While Not done
            s = ""
            s = s & "select distinct Skip,[Create],Map,BuildingPhase,Product,POIndex,Phase,Item,Description,TakeoffUnit" & vbCrLf
            s = s & "from BIMImportAssemblyDetails" & vbCrLf
            s = s & "where UnknownItem=1 and ignorebuildingphase=0 and isnull(skip,0)=0" & vbCrLf
            s = s & "and sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
            s = s & "order by poindex,item"
            Call LoadSQL(s)
            
            done = gData.Rows = 1
            
            If Not done Then
                mGridMode = "items"
                Call IniGetGrid(Me, gData, , , mGridMode)
                .ColHidden(.ColIndex("Create")) = Not mEditItemDB
                .ColHidden(.ColIndex("Map")) = Not mEditItemDB
                
                Me.Show vbModal
                Call IniPutGrid(Me, gData, , mGridMode)
                If mCancel Then GoTo ExitFunction
            
                'save changes
                For r = 1 To .Rows - 1
                If .Cell(flexcpData, r, 0) = "dirty" Then
                    s = ""
                    s = s & "update BIMImportDetails" & vbCrLf
                    s = s & "set map=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("map")) = flexChecked) & vbCrLf
                    s = s & "   ,skip=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("Skip")) = flexChecked) & vbCrLf
                    s = s & "   ,[create]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex("create")) = flexChecked) & vbCrLf
                    s = s & "   ,poindex=" & DbQuote(Str, .TextMatrix(r, .ColIndex("POIndex"))) & vbCrLf
                    s = s & "   ,phase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase"))) & vbCrLf
                    s = s & "   ,item=" & DbQuote(Str, .TextMatrix(r, .ColIndex("item"))) & vbCrLf
                    s = s & "   ,description=" & DbQuote(Str, .TextMatrix(r, .ColIndex("description"))) & vbCrLf
                    s = s & "   ,unit=" & DbQuote(Str, .TextMatrix(r, .ColIndex("takeoffunit"))) & vbCrLf
                    s = s & "where sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
                    s = s & "and buildingphase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("buildingphase"))) & vbCrLf
                    s = s & "and product=" & DbQuote(Str, .TextMatrix(r, .ColIndex("product"))) & vbCrLf
                    Call HFApp.SqlExec(s)
                End If
                Next
                
                'map items
                s = ""
                s = s & "update tblphaseitem set partnumber=d.product" & vbCrLf
                s = s & "from bimimportassemblydetails d" & vbCrLf
                s = s & "join tblphaseitem i on d.phase=i.phase and d.item=i.item and i.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                s = s & "where d.sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
                s = s & "and d.unknownitem=1" & vbCrLf
                s = s & "and d.ignorebuildingphase=0" & vbCrLf
                s = s & "and d.map=1" & vbCrLf
                Call HFApp.SqlExec(s)
                
                'create items
                If mEditItemDB Then
                    s = ""
                    s = s & "insert into tblphaseitem(partnumber,poindex,Phase,ItemNumber,Item,CostCategory,Description,OrderUOM,TakeoffUOM,ConversionFactor,WastePercent,RoundDir,Roundto,IsQuote,DivisionID)" & vbCrLf
                    s = s & "select distinct d.product,d.poindex,d.phase,left(d.item,len(d.item)-1),d.item,right(rtrim(d.item),1),d.Description,d.TakeoffUnit,d.TakeoffUnit,1,0,0,0,0," & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                    s = s & "from bimimportassemblydetails d" & vbCrLf
                    s = s & "where d.sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
                    s = s & "and d.unknownitem=1" & vbCrLf
                    s = s & "and d.ignorebuildingphase=0" & vbCrLf
                    s = s & "and d.[create]=1" & vbCrLf
                    Call HFApp.SqlExec(s)
                End If
            End If
        Loop
        
        gData.Visible = False
        frmFinish.Visible = True
        WizHead1.Description = "Ready to proceed."
        
        r = HFApp.SqlExec("select count(*) from bimimportassemblymasters where sessioncode=" & DbQuote(Str, mSessionCode))(0)
        
        lblReady.Caption = "The BIM Pipeline import is ready to proceed." & vbCrLf & vbCrLf & "This process will import " & r & " assemblies into HomeFront's database. If they already exist they will be completely replaced with the BIM version. Once completed this import is irreversible and cannot be undone." & vbCrLf & vbCrLf & "Are you sure you would like to proceed?"
        cmdNav(0).Caption = "&Save"
        Me.Show vbModal
        If mCancel Then GoTo ExitFunction
    
        'move from staging area to tblDBAssemblyMaster & Details
        
        s = ""
        s = s & "-- UPDATE EXISTING MASTERS -------------------------" & vbCrLf
        s = s & "update tbldbassemblymaster" & vbCrLf
        s = s & "set description=case when x.Description='' then m.description else x.Description end" & vbCrLf
        s = s & "   ,series= case when x.Series='' then m.series else x.Series end" & vbCrLf
        s = s & "   ,floorarea= case when isnull(x.FloorArea,0)=0 then m.floorarea else isnull(x.FloorArea,0) end" & vbCrLf
        s = s & "   ,bedrooms= case when isnull(x.Bedrooms,0)=0 then m.bedrooms else isnull(x.Bedrooms,0) end" & vbCrLf
        s = s & "   ,bathrooms= case when isnull(x.Bathrooms,0)=0 then m.bathrooms else isnull(x.Bathrooms,0) end" & vbCrLf
        s = s & "   ,style= case when x.Style='' then m.style else x.Style end" & vbCrLf
        s = s & "from bimimportassemblymasters x" & vbCrLf
        s = s & "join tbldbassemblymaster m on m.community='' and m.assembly=x.assembly and m.model=x.model and m.optionid=x.optionid and m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "where x.sessioncode = " & DbQuote(Str, mSessionCode) & vbCrLf
        s = s & vbCrLf
        s = s & "-- CREATE NEW MASTERS -------------------------" & vbCrLf
        s = s & "insert into tblDBAssemblyMaster(Community,Assembly,Description,OptionID,Model,Series,AssemblyType,FloorArea,Bedrooms,Bathrooms,Style,DivisionID)" & vbCrLf
        s = s & "select '',x.Assembly,x.Description,x.OptionID,x.Model,x.Series,x.AssemblyType,isnull(x.FloorArea,0),isnull(x.Bedrooms,0),isnull(x.Bathrooms,0),x.Style," & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "from bimimportassemblymasters x" & vbCrLf
        s = s & "left outer join tbldbassemblymaster m on m.community='' and m.assembly=x.assembly and m.model=x.model and m.optionid=x.optionid and m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "where m.assembly is null" & vbCrLf
        s = s & "and sessioncode = " & DbQuote(Str, mSessionCode) & vbCrLf
        
        If Not mIsModelFile Then
            s = s & vbCrLf
            s = s & "-- REMOVE ALL DETAILS -------------------------" & vbCrLf
            s = s & "delete tbldbassemblydetails" & vbCrLf
            s = s & "from bimimportassemblymasters x" & vbCrLf
            s = s & "join tbldbassemblydetails m on m.community=''" & vbCrLf
            s = s & "     and m.assembly=x.assembly" & vbCrLf
            s = s & "     and m.model=case when substring(x.OptionID,1,4) = '9999' then '' else x.model end" & vbCrLf
            s = s & "     and m.optionid=x.optionid" & vbCrLf
            s = s & "     and m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "where x.sessioncode = " & DbQuote(Str, mSessionCode) & vbCrLf
            s = s & vbCrLf
            s = s & "-- CREATE NEW DETAILS -------------------------" & vbCrLf
            s = s & "insert into tbldbassemblydetails(Community,Assembly,Model,OptionID,Phase,Item,TakeoffQty,OrderQty,POIndex,Notes,DivisionID,AssemblyID)" & vbCrLf
            s = s & "select '',m.Assembly,m.Model,m.OptionID,d.Phase,d.Item,d.TakeoffQty,d.TakeoffQty,d.POIndex,d.Notes,m.divisionid,m.AssemblyID" & vbCrLf
            s = s & "from bimimportassemblydetails d" & vbCrLf
            s = s & "join tbldbassemblymaster m on m.community='' " & vbCrLf
            s = s & "     and m.assembly=d.assembly " & vbCrLf
            s = s & "     and m.model=case when substring(d.OptionID,1,4) = '9999' then '' else d.model end" & vbCrLf
            s = s & "     and m.optionid=d.optionid " & vbCrLf
            s = s & "     and m.divisionid=isnull(d.divisionid," & DbQuote(Num, HFApp.DivisionID) & ")" & vbCrLf
            s = s & "where d.ignorebuildingphase=0 and isnull(d.skip,0)=0" & vbCrLf
            s = s & "and d.sessioncode = " & DbQuote(Str, mSessionCode) & vbCrLf
        End If
        
        Call HFApp.SqlExec(s)
    
    End With
    
ExitFunction:
    Unload Me
    Exit Function
eh:
    If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        MsgBox "Unable to create a new item because the phase/item number has already been used." & vbCrLf & vbCrLf & "---ERROR----------------------------------------------------" & vbCrLf & Err.Description, vbCritical, App.ProductName
        Resume Next
    Else
        Call errHandler(SRCFILE & "ShowForm")
        Screen.MousePointer = vbDefault
        Unload Me
    End If
End Function

Private Sub LoadFile(FileName As String)
    Dim c As Long
    Dim r As Long
    Dim i As Long
    Dim s As String
    Dim TableName As String

    With gData
        
        'load into grid
        .FixedRows = 0
        Call .LoadGrid(FileName, flexFileCommaText)
        .FixedRows = 1
        
        'set column keys
        mIsModelFile = .TextMatrix(0, 0) = "Plan Name"
        For c = 0 To .Cols - 1
            'column key becomes description with space period and underscore removed. these are also the column names in the table
            'Plan Name in the model file has goofy characters (﻿) in it so remove them too.
            .ColKey(c) = Replace(Replace(Replace(Replace(.TextMatrix(0, c), " ", ""), ".", ""), "_", ""), "﻿", "")
            'model file uses PlanNumber, detail file uses PlanNo. Be consistent and use PlanNumber everywhere.
            If .ColKey(c) = "PlanNo" Then .ColKey(c) = "PlanNumber"
            
            
            'v3 changes -- Mar2018
            ' - "Qty" has become "Product Qty" and "Adjusted Qty" has become "Total Qty"
            ' - 2 new ignorable qty columns added
            If .ColKey(c) = "ProductQty" Then .ColKey(c) = "Qty"
            If .ColKey(c) = "TotalQty" Then .ColKey(c) = "AdjustedQty"
            If .ColKey(c) = "WasteQty" Then .ColKey(c) = ""
            If .ColKey(c) = "RoundingQty" Then .ColKey(c) = ""
            
            
            'set datatype and size in coldata
            Select Case .ColKey(c)
                Case "BasePrice", "BasementSQF", "FirstSQF", "SecondSQF", "HeatedSQF", "TotalSQF", "Stories", "Bedrooms", "Bathrooms", "Qty", "Waste", "Accuracy", "AdjustedQty"
                    .ColData(c) = "numeric"
                Case "Style"
                    .ColData(c) = "string:200"
                Case "Description"
                    If mIsModelFile Then
                        .ColData(c) = "string:200"
                    Else
                        .ColData(c) = "string:150"
                    End If
                Case Else
                    .ColData(c) = "string"
            End Select
        Next
    
        'now write to table
        TableName = IIf(mIsModelFile, "BIMImportModels", "BIMImportDetails")
        
        'do in batches of 150
        i = 0
        r = 1
        While r < .Rows
            If (i = 0 Or i = 150) Then
                If i > 0 Then
                    s = Left(s, Len(s) - 3)
                    Call HFApp.SqlExec(s)
                End If
                
                s = "insert into " & TableName & "(divisionid,sessioncode"
                For c = 0 To .Cols - 1
                If .ColKey(c) <> "" Then
                    s = s & ",[" & .ColKey(c) & "]"
                End If
                Next
                s = s & ") values" & vbCrLf
                i = 0
                
            End If
            
            
            If .TextMatrix(r, .ColIndex("PlanNumber")) <> "" Then
                i = i + 1
                s = s & "(" & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, mSessionCode)
                For c = 0 To .Cols - 1
                If .ColKey(c) <> "" Then
                    Select Case .ColData(c)
                    Case "numeric"
                        s = s & "," & DbQuote(Num, .TextMatrix(r, c))
                    Case "string:200"
                        s = s & "," & DbQuote(Str, .TextMatrix(r, c), , True, 200)
                    Case "string:150"
                        s = s & "," & DbQuote(Str, .TextMatrix(r, c), , True, 150)
                    Case Else
                        s = s & "," & DbQuote(Str, .TextMatrix(r, c), , True)
                    End Select
                End If
                Next
                s = s & ")," & vbCrLf
            End If
            
            r = r + 1
        Wend
        
        'remember to insert last batch
        If i > 0 Then
            s = Left(s, Len(s) - 3)
            Call HFApp.SqlExec(s)
        End If
        
        Dim BaseModel As String
        BaseModel = HFApp.Options.ValueByName("BIMBaseModelNumber")
        If BaseModel = "" Then BaseModel = "BASE0100"
                
        'extract models from details file.
        s = ""
        s = s & "insert into BIMImportModels(DivisionID,SessionCode, PlanName, PlanNumber, SalesHouseName,Description)" & vbCrLf
        s = s & "select top 1" & vbCrLf
        s = s & "  x.divisionid" & vbCrLf
        s = s & ", x.SessionCode" & vbCrLf
        s = s & ", m.Assembly PlanName" & vbCrLf
        s = s & ", m.Assembly PlanNumber" & vbCrLf
        s = s & ", m.Assembly SalesHouseName" & vbCrLf
        s = s & ", m.Description" & vbCrLf
        s = s & "from bimimportdetails x" & vbCrLf
        s = s & "join tbldbassemblymaster m on m.assembly=x.PlanNumber and AssemblyType=0" & vbCrLf
        s = s & "where x.sessioncode=" & DbQuote(Str, mSessionCode) & vbCrLf
        s = s & "and x.Optionnumber=" & DbQuote(Str, BaseModel) & vbCrLf
        Call HFApp.SqlExec(s)
    
    End With
    
End Sub

Private Sub LoadSQL(sql As String)
    Dim rs As Recordset
    Dim r As Long
    Dim c As Long
    
    With gData
        .Rows = 1
        .Cols = 0
        
        Set rs = HFApp.SqlExec(sql)
        .Cols = rs.fields.Count
        For c = 0 To .Cols - 1
            .TextMatrix(0, c) = rs.fields(c).Name
            .ColKey(c) = rs.fields(c).Name
            If rs.fields(c).Type = adBoolean Then
                .ColDataType(c) = flexDTBoolean
            End If
        Next
        
        r = 0
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            For c = 0 To .Cols - 1
                .TextMatrix(r, c) = "" & rs(c)
            Next
            rs.MoveNext
        Wend
    
        Call .AutoSize(0, .Cols - 1)
    End With

End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 And mGridMode = "items" Then
        If Not ValidateItems Then Exit Sub
    End If
    
    mCancel = Index = 1
    Me.Hide
    
End Sub


Private Function ValidateItems() As Boolean
On Error GoTo eh:
    'check that all unknown items have been dealt with.
    
    Dim r As Long
    Dim s As String
    
    With gData
    For r = 1 To .Rows - 1
        Select Case True
            Case .Cell(flexcpChecked, r, .ColIndex("Skip")) = flexChecked
            Case .Cell(flexcpChecked, r, .ColIndex("Create")) = flexChecked
                If .TextMatrix(r, .ColIndex("POIndex")) = "" Or .TextMatrix(r, .ColIndex("Phase")) = "" Or .TextMatrix(r, .ColIndex("Item")) = "" Then
                    MsgBox "POIndex, Phase, and Item are all required to create a new item.", vbInformation, App.ProductName
                    ValidateItems = False
                    Exit Function
                End If
            
            Case .Cell(flexcpChecked, r, .ColIndex("Map")) = flexChecked
                If .TextMatrix(r, .ColIndex("Phase")) = "" Or .TextMatrix(r, .ColIndex("Item")) = "" Then
                    MsgBox "Phase and Item are required to map an item.", vbInformation, App.ProductName
                    ValidateItems = False
                    Exit Function
                End If
        End Select
    Next
    End With
    
    ValidateItems = True
    
Exit Function
eh:
    If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "ValidateItems")
    End If
End Function

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60

    gData.Move margin, Me.WizHead1.Height + margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - WizHead1.Height - WizFoot.Height - 2 * margin
    frmFinish.Move gData.Left, gData.Top, gData.Width, gData.Height
    lblReady.Move lblReady.Left, lblReady.Top, gData.Width - 2 * lblReady.Left, gData.Height - lblReady.Top
 
    cmdNav(0).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin

End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim s As String
    
    s = ""
    s = s & "delete BIMImportModels where SessionCode=" & DbQuote(Str, mSessionCode) & vbCrLf
    s = s & "delete BIMImportDetails where SessionCode=" & DbQuote(Str, mSessionCode) & vbCrLf
    Call HFApp.SqlExec(s)
    
    Call IniPutForm(Me)
    
End Sub

Private Sub gData_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
On Error Resume Next
    gData.ColSel = gData.Col
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'models & options
'    Description string
'    series      picklist
'    FloorArea   numeric
'    Bedrooms    numeric
'    Bathrooms   numeric
'    style       picklist
'POIndex
'    Ignore      checkbox
'items
'    Skip        checkbox
'    Create      checkbox
'    map         checkbox
'    poindex     picklist
'    phase       picklist
'    item        picklist
    
    With gData
    .ComboList = ""
    .AutoSearch = flexSearchNone
    
    Select Case mGridMode
    Case "models", "options"
        Select Case .ColKey(Col)
        Case "Description"
            .EditMaxLength = 200
        Case "Style"
            .EditMaxLength = 200
            Cancel = .TextMatrix(Row, .ColIndex("AssemblyType")) <> "0"
        Case "FloorArea", "Bedrooms", "Bathrooms"
            Cancel = .TextMatrix(Row, .ColIndex("AssemblyType")) <> "0"
        Case "Series"
            Cancel = .TextMatrix(Row, .ColIndex("AssemblyType")) <> "0"
            .EditMaxLength = 30
            .ComboList = "..."
        Case Else
            .AutoSearch = flexSearchFromCursor
            Cancel = True
        End Select
    
    
    Case "buildingphase"
        Select Case .ColKey(Col)
        Case "Ignore"
        Case Else
            .AutoSearch = flexSearchFromCursor
            Cancel = True
        End Select
    
    
    Case "items"
        Select Case .ColKey(Col)
        
        Case "Skip"
            .EditMaxLength = 150
            
        Case "Description", "Create", "Map"
            .EditMaxLength = 150
            Cancel = Not mEditItemDB
            
        Case "Phase"
            .EditMaxLength = 20
            .ComboList = IIf(mEditItemDB, "|...", "...")
            
        Case "Item"
            .EditMaxLength = 16
            .ComboList = "|..."
            Cancel = Not mEditItemDB
            
        Case "poindex"
            .ComboList = "..."
        
        Case Else
            .AutoSearch = flexSearchFromCursor
            Cancel = True
        End Select
        
    End Select
    End With

End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim filter As String
    With gData
    Select Case .ColKey(Col)
        Case "Series":
            s = "SELECT series,description FROM tblSeries where divisionid = " & DbQuote(Num, HFApp.DivisionID)
            If FPickList.Choose(HFApp.Databases(dbHomefront), .ColKey(Col), s) Then
                .Cell(flexcpText, Row, Col, .RowSel, Col) = FPickList.SelectedItem(1)
                .Cell(flexcpData, Row, 0, .RowSel, 0) = "dirty"
            End If
        
        Case "POIndex":
            If .TextMatrix(Row, .ColIndex("poindex")) = "" Then filter = "poindex=" & Parse(Parse(.TextMatrix(Row, .ColIndex("buildingphase")), 1, " "), 1, "-")
            s = "select poindex,description from tblpoindex where divisionid = " & DbQuote(Num, HFApp.DivisionID)
            If FPickList.Choose(HFApp.Databases(dbHomefront), .ColKey(Col), s, .TextMatrix(Row, .ColIndex("poindex")), , , , , , , filter) Then
                .Cell(flexcpText, Row, Col, .RowSel, Col) = FPickList.SelectedItem(1)
                
                .Cell(flexcpData, .Row, 0, .RowSel, 0) = "dirty"
            End If
        
        Case "Phase"
            s = "select Phase,Description from estphases where GroupPhase=0 and divisionid = " & DbQuote(Num, HFApp.DivisionID)
            If FPickList.Choose(HFApp.Databases(dbHomefront), .ColKey(Col), s, .TextMatrix(Row, .ColIndex("phase")), , , , , , , filter) Then
                .Cell(flexcpText, Row, Col, .RowSel, Col) = FPickList.SelectedItem(1)
                
                .Cell(flexcpChecked, .Row, .ColIndex("map"), .RowSel, .ColIndex("map")) = flexUnchecked
                .Cell(flexcpChecked, .Row, .ColIndex("skip"), .RowSel, .ColIndex("skip")) = flexUnchecked
                .Cell(flexcpChecked, .Row, .ColIndex("create"), .RowSel, .ColIndex("create")) = flexChecked
                .Cell(flexcpData, .Row, 0, .RowSel, 0) = "dirty"
            End If
            
            
        Case "Item":
            filter = "description=" & Parse(.TextMatrix(Row, .ColIndex("description")), 1, " ")
            s = "select POIndex,Phase,Item,Description,PartNumber,TakeoffUOM from tblphaseitem where divisionid=" & DbQuote(Num, HFApp.DivisionID)
            If FPickList.Choose(HFApp.Databases(dbHomefront), .ColKey(Col), s, , , , , , , , filter) Then
                .Cell(flexcpText, .Row, .ColIndex("POIndex"), .RowSel, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                .Cell(flexcpText, .Row, .ColIndex("Phase"), .RowSel, .ColIndex("Phase")) = FPickList.SelectedItem("Phase")
                .Cell(flexcpText, .Row, .ColIndex("Item"), .RowSel, .ColIndex("Item")) = FPickList.SelectedItem("Item")
                .Cell(flexcpText, .Row, .ColIndex("Description"), .RowSel, .ColIndex("Description")) = FPickList.SelectedItem("Description")
                .Cell(flexcpText, .Row, .ColIndex("TakeoffUnit"), .RowSel, .ColIndex("TakeoffUnit")) = FPickList.SelectedItem("TakeoffUOM")
                
                .Cell(flexcpChecked, .Row, .ColIndex("map"), .RowSel, .ColIndex("map")) = flexChecked
                .Cell(flexcpChecked, .Row, .ColIndex("skip"), .RowSel, .ColIndex("skip")) = flexUnchecked
                .Cell(flexcpChecked, .Row, .ColIndex("create"), .RowSel, .ColIndex("create")) = flexUnchecked
                .Cell(flexcpData, .Row, 0, .RowSel, 0) = "dirty"
            End If
            
    End Select
    End With
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
    'validateedit happens before the selection changes, but we have to use the selection. fix it if its obviously wrong
    If Row <> gData.Row Then .Row = Row
    
    Select Case .ColKey(Col)
        Case "FloorArea", "Bedrooms", "Bathrooms"
            .EditText = Val(.EditText)
            
        Case "Map"
            If Val(.EditText) = flexChecked Then
                .Cell(flexcpChecked, .Row, .ColIndex("skip"), .RowSel, .ColIndex("skip")) = flexUnchecked
                .Cell(flexcpChecked, .Row, .ColIndex("create"), .RowSel, .ColIndex("create")) = flexUnchecked
            End If
        
        Case "Skip"
            If Val(.EditText) = flexChecked Then
                .Cell(flexcpChecked, .Row, .ColIndex("map"), .RowSel, .ColIndex("map")) = flexUnchecked
                .Cell(flexcpChecked, .Row, .ColIndex("create"), .RowSel, .ColIndex("create")) = flexUnchecked
            End If
            
        Case "Create"
            If Val(.EditText) = flexChecked Then
                .Cell(flexcpChecked, .Row, .ColIndex("map"), .RowSel, .ColIndex("map")) = flexUnchecked
                .Cell(flexcpChecked, .Row, .ColIndex("skip"), .RowSel, .ColIndex("skip")) = flexUnchecked
            End If
        
        Case "Phase"
            .Cell(flexcpChecked, .Row, .ColIndex("map"), .RowSel, .ColIndex("map")) = flexUnchecked
            .Cell(flexcpChecked, .Row, .ColIndex("create"), .RowSel, .ColIndex("create")) = flexChecked
            .Cell(flexcpChecked, .Row, .ColIndex("skip"), .RowSel, .ColIndex("skip")) = flexUnchecked
    
        Case "Item"
            .EditText = Trim(.EditText)
            'make sure item includes costtype
            If Not IsIn(Right(.EditText, 1), "L", "E", "M", "S", "O") Then .EditText = .EditText & "M"
            'make sure costtype is uppercase
            .EditText = Left(.EditText, Len(.EditText) - 1) & UCase(Right(.EditText, 1))
            
            If Len(.EditText) > 16 Then .EditText = Left(.EditText, 15) & Right(.EditText, 1)
            
            .Cell(flexcpChecked, .Row, .ColIndex("map"), .RowSel, .ColIndex("map")) = flexUnchecked
            .Cell(flexcpChecked, .Row, .ColIndex("create"), .RowSel, .ColIndex("create")) = flexChecked
            .Cell(flexcpChecked, .Row, .ColIndex("skip"), .RowSel, .ColIndex("skip")) = flexUnchecked
    
    End Select
    .Cell(flexcpData, .Row, 0, .RowSel, 0) = "dirty"
    End With
End Sub
