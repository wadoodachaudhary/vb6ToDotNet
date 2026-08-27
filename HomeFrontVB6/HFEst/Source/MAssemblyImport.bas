Attribute VB_Name = "MAssemblyExport"
Option Explicit
Option Compare Text
Const SRCFILE = "MAssemblyExport::"


Public Sub ExportAssemblies()
On Error GoTo eh

    Dim s As String
    Dim rs As Recordset
    
    Dim i As Long
    Dim Models As String
    Dim AssemblyIDs As String
    Dim AssemblyType As String
    
    'select upto 400 assemblies
    s = ""
    s = s & "Models" & Chr(1) & "select a.AssemblyID, c.area Community,c.description CommunityDesc,a.Model,a.Assembly,a.Description,a.Series,a.Style from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
    s = s & "Model Specific Options" & Chr(1)
        s = s & "select distinct m.Model,m.Description" & vbCrLf
        s = s & "from tbldbassemblymaster o" & vbCrLf
        s = s & "join DistinctModelsByDivision m on o.divisionid=m.divisionid and m.model=o.model" & vbCrLf
        s = s & "where o.assemblytype=2" & vbCrLf
        s = s & "and isnull(o.inactive,0)=0 and o.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf & Chr(0)
    s = s & "Global Options" & Chr(1) & "select a.AssemblyID, ct.Description SubCat,c.area Community,c.description CommunityDesc,a.OptionID [Option],a.Assembly,a.Description from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=3 order by 2,3,4,5,6,7" & Chr(0)
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Assemblies", s, , , , , "assemblyid", True) Then Exit Sub
    
    Select Case FPickList.SelectedView
    Case 1: AssemblyType = "models"
    Case 3: AssemblyType = "globals"
    Case 2:
        AssemblyType = "options"
        
        For i = 1 To FPickList.SelectedItems
            Models = Models & "," & DbQuote(Str, FPickList.SelectedItem("Model", i))
        Next
        Models = Mid(Models, 2)
        
        s = ""
        s = s & "select a.AssemblyID,c.area Community,c.description CommunityDesc ,a.Model,dm.description ModelDescription,a.Series,a.OptionID [Option],a.Description, ct.Description Category" & vbCrLf
        s = s & "from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid)" & vbCrLf
        s = s & "where a.DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & " and isnull(a.inactive,0)=0" & vbCrLf
        s = s & " and a.assemblytype=2" & vbCrLf
        s = s & " and a.model in(" & Models & ")" & vbCrLf
        s = s & " order by 2,4,6,7" & vbCrLf
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Assemblies", s, , , , , "assemblyid", True) Then Exit Sub
    End Select
    
    For i = 1 To FPickList.SelectedItems
        AssemblyIDs = AssemblyIDs & "," & FPickList.SelectedItem("AssemblyID", i)
    Next
    AssemblyIDs = Mid(AssemblyIDs, 2)
    If AssemblyIDs = "" Then Exit Sub
    
    
    
    Select Case AssemblyType
    Case "models"
        s = ""
        s = s & "select" & vbCrLf
        s = s & " Community CommunityCode,Model,Assembly" & vbCrLf
        s = s & ",case when isnull(Inactive,0)=0 then 'n' else 'y' end Inactive" & vbCrLf
        s = s & ",Description" & vbCrLf
        s = s & ",GraphicPath Image" & vbCrLf
        s = s & ",Comments" & vbCrLf
        s = s & ",Series" & vbCrLf
        s = s & ",Style" & vbCrLf
        s = s & ",ScheduleTemplate" & vbCrLf
        s = s & ",Bedrooms" & vbCrLf
        s = s & ",Bathrooms" & vbCrLf
        s = s & ",FloorArea" & vbCrLf
        s = s & ",MaxWidth" & vbCrLf
        s = s & ",MaxLength" & vbCrLf
        s = s & ",SpecDocument" & vbCrLf
        s = s & ",case when isnull(TakeoffRequired,0)=0 then 'n' else 'y' end TakeoffRequired" & vbCrLf
        s = s & ",case when isnull(UseNormalSalesQtyFactors,0)=0 then 'n' else 'y' end UseNormalSalesQtyFactors" & vbCrLf
        s = s & ",MainSF" & vbCrLf
        s = s & ",UpperSF" & vbCrLf
        s = s & ",LowerSF" & vbCrLf
        s = s & ",GarageSF" & vbCrLf
        s = s & "from tbldbassemblymaster " & vbCrLf
        s = s & "where AssemblyID in(" & AssemblyIDs & ")"
    
    Case "options"
        s = ""
        s = s & "select" & vbCrLf
        s = s & " a.Community CommunityCode,a.Model,a.OptionID [Option],a.Assembly" & vbCrLf
        s = s & ",case when isnull(a.Inactive,0)=0 then 'n' else 'y' end Inactive" & vbCrLf
        s = s & ",a.Description" & vbCrLf
        s = s & ",a.GraphicPath Image" & vbCrLf
        s = s & ",a.Comments" & vbCrLf
        s = s & ",a.Notes EstimatorNotes" & vbCrLf
        s = s & ",a.Series" & vbCrLf
        s = s & ",a.Category SubCat" & vbCrLf
        s = s & ",a.JCExtra" & vbCrLf
        s = s & ",a.Qty" & vbCrLf
        s = s & ",a.AssemblyUOM UOM" & vbCrLf
        s = s & ",a.ConstCutoff" & vbCrLf
        s = s & ",a.Location" & vbCrLf
        s = s & ",case when isnull(a.IncludedOption,0)=0 then 'n' else 'y' end IncludedOption" & vbCrLf
        s = s & ",case when isnull(a.DesignCenterSalesOnly,0)=0 then 'n' else 'y' end DesignCenterSalesOnly" & vbCrLf
        s = s & ",case when isnull(a.SelectByRoom,0)=0 then 'n' else 'y' end SelectByRoom" & vbCrLf
        s = s & ",case when isnull(a.DisplayTotalOnly,0)=0 then 'n' else 'y' end DisplayTotalOnly" & vbCrLf
        s = s & ",case when isnull(a.TakeoffRequired,0)=0 then 'n' else 'y' end TakeoffRequired" & vbCrLf
        s = s & ",case when isnull(a.UseNormalSalesQtyFactors,0)=0 then 'n' else 'y' end UseNormalSalesQtyFactors" & vbCrLf
        s = s & ",clist.name ColorListName" & vbCrLf
        s = s & ",slist.name StyleListName" & vbCrLf
        s = s & ",flist.name FinishListName" & vbCrLf
        s = s & ",olist.name OtherListName" & vbCrLf
        s = s & ",a.Color" & vbCrLf
        s = s & ",a.StyleValue" & vbCrLf
        s = s & ",a.FinishValue" & vbCrLf
        s = s & ",a.OtherValue" & vbCrLf
        s = s & "from tbldbassemblymaster a" & vbCrLf
        s = s & "left join AttributeLists clist on a.ColorListID=cList.listid" & vbCrLf
        s = s & "left join AttributeLists slist on a.StyleListID=sList.listid" & vbCrLf
        s = s & "left join AttributeLists flist on a.FinishListID=fList.listid" & vbCrLf
        s = s & "left join AttributeLists olist on a.OtherListID=oList.listid" & vbCrLf
        s = s & "where a.AssemblyID in(" & AssemblyIDs & ")"
        
    Case "globals"
        s = ""
        s = s & "select" & vbCrLf
        s = s & " a.Community CommunityCode,a.OptionID [Option],a.Assembly" & vbCrLf
        s = s & ",case when isnull(a.Inactive,0)=0 then 'n' else 'y' end Inactive" & vbCrLf
        s = s & ",a.Description" & vbCrLf
        s = s & ",a.GraphicPath Image" & vbCrLf
        s = s & ",a.Comments" & vbCrLf
        s = s & ",a.Notes EstimatorNotes" & vbCrLf
        s = s & ",a.Category SubCat" & vbCrLf
        s = s & ",a.JCExtra" & vbCrLf
        s = s & ",a.Qty" & vbCrLf
        s = s & ",a.AssemblyUOM UOM" & vbCrLf
        s = s & ",a.ConstCutoff" & vbCrLf
        s = s & ",a.Location" & vbCrLf
        s = s & ",case when isnull(a.IncludedOption,0)=0 then 'n' else 'y' end IncludedOption" & vbCrLf
        s = s & ",case when isnull(a.DesignCenterSalesOnly,0)=0 then 'n' else 'y' end DesignCenterSalesOnly" & vbCrLf
        s = s & ",case when isnull(a.SelectByRoom,0)=0 then 'n' else 'y' end SelectByRoom" & vbCrLf
        s = s & ",case when isnull(a.DisplayTotalOnly,0)=0 then 'n' else 'y' end DisplayTotalOnly" & vbCrLf
        s = s & ",case when isnull(a.TakeoffRequired,0)=0 then 'n' else 'y' end TakeoffRequired" & vbCrLf
        s = s & ",case when isnull(a.UseNormalSalesQtyFactors,0)=0 then 'n' else 'y' end UseNormalSalesQtyFactors" & vbCrLf
        s = s & ",clist.name ColorListName" & vbCrLf
        s = s & ",slist.name StyleListName" & vbCrLf
        s = s & ",flist.name FinishListName" & vbCrLf
        s = s & ",olist.name OtherListName" & vbCrLf
        s = s & ",a.Color" & vbCrLf
        s = s & ",a.StyleValue" & vbCrLf
        s = s & ",a.FinishValue" & vbCrLf
        s = s & ",a.OtherValue" & vbCrLf
        s = s & "from tbldbassemblymaster a" & vbCrLf
        s = s & "left join AttributeLists clist on a.ColorListID=cList.listid" & vbCrLf
        s = s & "left join AttributeLists slist on a.StyleListID=sList.listid" & vbCrLf
        s = s & "left join AttributeLists flist on a.FinishListID=fList.listid" & vbCrLf
        s = s & "left join AttributeLists olist on a.OtherListID=oList.listid" & vbCrLf
        s = s & "where a.AssemblyID in(" & AssemblyIDs & ")"
    End Select

    Call HFApp.ExportData(s, "Export Assemblies")

Exit Sub
eh: Call errHandler(SRCFILE & "ExportAssemblies")
End Sub

Public Sub ExportAssemblyTakeoffs()
On Error GoTo eh

    If FProgress.Visible Then Exit Sub

    Dim srcXLS As String
    Dim dstXLS As String
    Dim s As String
    Dim rs As Recordset
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSH As Object
    
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim rowCount As Long
    Dim AssemblyIDs As String
    Dim AssemblyCount As Long
    Dim PivotColumns As String
    
    
Const RowCommunity = 1
Const RowModel = 2
Const RowOption = 3
Const RowAssembly = 4
Const RowDescription = 5
Const RowComments = 6
Const RowSeries = 7
Const RowBPTemplate = 8
Const RowNotes = 9
Const RowCategory = 10
Const RowConstCutoff = 11
Const RowAssemblyUOM = 12
Const RowDataStart = 13

    
    'select upto 400 assemblies
    s = ""
    s = s & "select AssemblyTypeDesc AssemblyType,AssemblyID,Community,Model,OptionID,Assembly,Description" & vbCrLf
    s = s & "from tbldbassemblymaster " & vbCrLf
    s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
    s = s & "and isnull(AssemblyType,0)<>4" & vbCrLf 'exclude design center options
    s = s & "and isnull(isbaseassembly,0)=0" & vbCrLf
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Assemblies", s, , , , , "AssemblyID", True) Then Exit Sub
    AssemblyCount = FPickList.SelectedItems
    For i = 1 To AssemblyCount
        AssemblyIDs = AssemblyIDs & "," & FPickList.SelectedItem("AssemblyID", i)
    Next
    AssemblyIDs = Mid(AssemblyIDs, 2)
    If AssemblyIDs = "" Then Exit Sub
    
    
    'copy template to destination
    srcXLS = PathAppend(HFApp.SystemFolder, "Estimating\Templates\NewAssmImports.xlsx")
                          
    If Not VBGetSaveFileName(dstXLS, , , "Excel Files (*.xlsx)|*.xlsx", , , , 1, FMain.hwnd) Then Exit Sub
    Call FileCopy(srcXLS, dstXLS)
    
        
Call FProgress.Progress("Initializating...", "loading Excel", 0, 0, FMain)

    
    'write validation data to sheet 2
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Open(dstXLS)
    Set xlSH = xlWB.Worksheets(2)
    
    
'col 1. BPTemplates"
Call FProgress.Progress(, "writing BuildPro templates")
    s = Replace(HFApp.Options.ValueByName("ScheduleTemplates"), "|", ",")
    If s <> "" Then Call WriteColumn(xlSH, 1, , , s)
     

'col 2. Item as "phase/item - description (OrderUOM)"
Call FProgress.Progress("Preparing Template...", "writing item db")
    
    s = ""
    s = s & "select count(*)" & vbCrLf
    s = s & "from tblphaseitem" & vbCrLf
    s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
    rowCount = Val(HFApp.SqlExec(s)(0))
    
    s = ""
    s = s & "select phase + '\' + item + ' - ' + isnull(description,'') + isnull(' (' + orderuom + ')','')" & vbCrLf
    s = s & "from tblphaseitem" & vbCrLf
    s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
    s = s & "order by 1" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    Call WriteColumn(xlSH, 2, rs, rowCount)
    
'col 3. SubCat as "category - description"
Call FProgress.Progress(, "writing option categories")
    s = ""
    s = s & "select count(*)" & vbCrLf
    s = s & "from tblcategories" & vbCrLf
    rowCount = Val(HFApp.SqlExec(s)(0))
    
    s = ""
    s = s & "select category + ' - ' + isnull(description,'') " & vbCrLf
    s = s & "from tblcategories" & vbCrLf
    s = s & "order by 1" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    Call WriteColumn(xlSH, 3, rs, rowCount)
    
'col 4. JobStatus as "status - description
Call FProgress.Progress(, "writing job statuses")
    s = ""
    s = s & "select count(*)" & vbCrLf
    s = s & "from tbljobstatus" & vbCrLf
    rowCount = Val(HFApp.SqlExec(s)(0))
    
    s = ""
    s = s & "select cast(job_status as varchar(10)) + ' - ' + isnull(description,'') " & vbCrLf
    s = s & "from tbljobstatus" & vbCrLf
    s = s & "order by job_status" & vbCrLf
    rowCount = Val(HFApp.SqlExec("select count(*) from tblphaseitem where divisionid=" & HFApp.DivisionID)(0))
    Set rs = HFApp.SqlExec(s)
    Call WriteColumn(xlSH, 4, rs, rowCount)
    
'col 5. Series as "series - description"
Call FProgress.Progress(, "writing series")
    
    s = ""
    s = s & "select count(*)" & vbCrLf
    s = s & "from tblseries" & vbCrLf
    s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
    rowCount = Val(HFApp.SqlExec(s)(0))
    
    s = ""
    s = s & "select series + ' - ' + isnull(description,'') " & vbCrLf
    s = s & "from tblseries" & vbCrLf
    s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
    s = s & "order by 1" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    Call WriteColumn(xlSH, 5, rs, rowCount)
    
'col 6. Communities as "community - description"
Call FProgress.Progress(, "writing communities")
    
    s = ""
    s = s & "select count(*)" & vbCrLf
    s = s & "from tbllocality l" & vbCrLf
    s = s & "join DivisionCommunities d on l.area=d.community" & vbCrLf
    s = s & "where d.divisionid=" & HFApp.DivisionID & vbCrLf
    rowCount = Val(HFApp.SqlExec(s)(0))
    
    s = ""
    s = s & "select area + ' - ' + isnull(description,'') " & vbCrLf
    s = s & "from tbllocality l" & vbCrLf
    s = s & "join DivisionCommunities d on l.area=d.community" & vbCrLf
    s = s & "where d.divisionid=" & HFApp.DivisionID & vbCrLf
    s = s & "order by 1" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    Call WriteColumn(xlSH, 6, rs, rowCount)
    
'col 7. POIndex as "poindex - description"
Call FProgress.Progress(, "writing poindexes")
    
    s = ""
    s = s & "select count(*)" & vbCrLf
    s = s & "from tblpoindex" & vbCrLf
    s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
    rowCount = Val(HFApp.SqlExec(s)(0))
    
    s = ""
    s = s & "select poindex + ' - ' + isnull(description,'') " & vbCrLf
    s = s & "from tblpoindex" & vbCrLf
    s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
    s = s & "order by 1" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    Call WriteColumn(xlSH, 7, rs, rowCount)
Unload FProgress
       
'protect the data validation worksheet
'Call Sheet.Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns, AllowFormattingRows
'                 , AllowInsertingColumns, AllowInsertingRows, AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows, AllowSorting, AllowFiltering, AllowUsingPivotTables)
xlSH.columns("A:I").AutoFit
xlSH.Rows("2:55555").RowHeight = xlSH.Rows("1").RowHeight
Call xlSH.Protect("admin", True, True, True, True, True, True, True)

       
    
    'write headers to sheet 1
    Set xlSH = xlWB.Worksheets(1)
Call FProgress.Progress(, "writing assembly headers")
    s = ""
    s = s & "select top 500" & vbCrLf
    s = s & " l.area + ' - ' + isnull(l.description,'') Community" & vbCrLf
    s = s & ",a.Model" & vbCrLf
    s = s & ",a.OptionID" & vbCrLf
    s = s & ",a.Assembly" & vbCrLf
    s = s & ",a.AssemblyID" & vbCrLf
    s = s & ",a.Description" & vbCrLf
    s = s & ",a.Comments" & vbCrLf
    s = s & ",a.notes" & vbCrLf
    s = s & ",s.series + ' - ' + isnull(s.description,'') series" & vbCrLf
    s = s & ",a.scheduletemplate" & vbCrLf
    
    's = s & ",c.category + ' - ' + isnull(c.description,'') category" & vbCrLf
    's = s & ",cast(j.job_status as varchar(10)) + ' - ' + isnull(j.description,'') constcutoff" & vbCrLf
    's = s & ",a.assemblyuom" & vbCrLf
    
    s = s & ",case when a.assemblytype=0 then '' else c.category + ' - ' + isnull(c.description,'') end category" & vbCrLf
    s = s & ",case when a.assemblytype=0 then '' else cast(j.job_status as varchar(10)) + ' - ' + isnull(j.description,'') end constcutoff" & vbCrLf
    s = s & ",case when a.assemblytype=0 then '' else a.assemblyuom end assemblyuom" & vbCrLf
    
    s = s & "from tbldbassemblymaster a" & vbCrLf
    s = s & "left join tblseries s on a.divisionid=s.divisionid and a.series=s.series" & vbCrLf
    s = s & "left join tblcategories c on a.category=c.category" & vbCrLf
    s = s & "left join tbljobstatus j on a.constcutoff=j.job_status" & vbCrLf
    s = s & "left join tbllocality l on a.community=l.area" & vbCrLf
    s = s & "where a.AssemblyID in(" & AssemblyIDs & ")" & vbCrLf
    s = s & "order by 1,2,3,4" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    i = 0
    While Not rs.EOF
        i = i + 1
        c = 5 + i
Call FProgress.Progress(, , i, AssemblyCount)
        PivotColumns = PivotColumns & ",[" & rs("AssemblyID") & "]"
        
        If "" & rs("Community") <> "" Then xlSH.Cells(RowCommunity, c).value = "'" & rs("Community")
        If "" & rs("Model") <> "" Then xlSH.Cells(RowModel, c).value = "'" & rs("Model")
        If "" & rs("OptionID") <> "" Then xlSH.Cells(RowOption, c).value = "'" & rs("OptionId")
        If "" & rs("Assembly") <> "" Then xlSH.Cells(RowAssembly, c).value = "'" & rs("Assembly")
        If "" & rs("Description") <> "" Then xlSH.Cells(RowDescription, c).value = "'" & rs("Description")
        
        If "" & rs("Comments") <> "" Then xlSH.Cells(RowComments, c).value = "'" & rs("Comments")
        If "" & rs("Series") <> "" Then xlSH.Cells(RowSeries, c).value = "'" & rs("Series")
        If "" & rs("ScheduleTemplate") <> "" Then xlSH.Cells(RowBPTemplate, c).value = "'" & rs("ScheduleTemplate")
        
        If "" & rs("Notes") <> "" Then xlSH.Cells(RowNotes, c).value = "'" & rs("Notes")
        If "" & rs("Category") <> "" Then xlSH.Cells(RowCategory, c).value = "'" & rs("Category")
        If "" & rs("ConstCutoff") <> "" Then xlSH.Cells(RowConstCutoff, c).value = "'" & rs("ConstCutoff")
        If "" & rs("AssemblyUOM") <> "" Then xlSH.Cells(RowAssemblyUOM, c).value = "'" & rs("AssemblyUOM")
        rs.MoveNext
    Wend
    PivotColumns = Mid(PivotColumns, 2)
        
    
Call FProgress.Progress(, "writing assembly quantities")
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "from" & vbCrLf
    s = s & "(" & vbCrLf
    s = s & "    select " & vbCrLf
    s = s & "     p.poindex + ' - ' + isnull(p.description,'') poindex" & vbCrLf
    s = s & "    ,d.phase + '\' + d.item + ' - ' + isnull(i.description,'') + isnull(' (' + i.takeoffuom + ')','') item" & vbCrLf
    s = s & "    ,case when isnull(d.invertable,0)=1 then 'y' else '' end invertable" & vbCrLf
    s = s & "    ,d.location" & vbCrLf
    s = s & "    ,d.notes" & vbCrLf
    s = s & "    ,d.takeoffqty" & vbCrLf
    s = s & "    ,a.AssemblyID" & vbCrLf
    s = s & "    from tbldbassemblymaster a" & vbCrLf
    s = s & "    join tbldbassemblydetails d on a.assemblyid=d.assemblyid" & vbCrLf
    s = s & "    left join tblphaseitem i on a.divisionid=i.divisionid and d.phase=i.phase and d.item=i.item" & vbCrLf
    s = s & "    left join tblpoindex p on d.divisionid=p.divisionid and d.poindex=p.poindex" & vbCrLf
    s = s & "    where a.AssemblyID in(" & AssemblyIDs & ")" & vbCrLf
    s = s & ") p" & vbCrLf
    s = s & "pivot" & vbCrLf
    s = s & "(" & vbCrLf
    s = s & "    sum(takeoffqty)" & vbCrLf
    s = s & "    for assemblyid in(" & PivotColumns & ")" & vbCrLf
    s = s & ") as pvt" & vbCrLf
    s = s & "order by 2,1,3,4,5" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    r = RowDataStart
    While Not rs.EOF
        If "" & rs("item") <> "" Then
            r = r + 1
            If "" & rs("poindex") <> "" Then xlSH.Cells(r, 1).value = "'" & rs("poindex")
            xlSH.Cells(r, 2).value = "'" & rs("item")
            If "" & rs("Invertable") <> "" Then xlSH.Cells(r, 3).value = "" & rs("invertable")
            If "" & rs("Location") <> "" Then xlSH.Cells(r, 4).value = "'" & rs("location")
            If "" & rs("Notes") <> "" Then xlSH.Cells(r, 5).value = "'" & rs("notes")
            
            For i = 1 To AssemblyCount
                If "" & rs(4 + i) <> "" Then
                    c = 5 + i
                    xlSH.Cells(r, c).value = rs(4 + i)
                End If
            Next
        End If
        rs.MoveNext
    Wend
    
xlSH.Rows("2:55555").RowHeight = xlSH.Rows("1").RowHeight
    
    Unload FProgress
    xlWB.Save
    Set xlWB = Nothing
    xlApp.Visible = True
    Set xlApp = Nothing
    
Exit Sub
eh:
    If Err.Number = 70 Then
        MsgBox "Permission Denied." & vbCrLf & dstXLS & " is in use.", vbExclamation, App.ProductName
    Else
        Call errHandler(SRCFILE & "ExportAssemblyTakeoffs")
    End If
    Unload FProgress
    xlWB.Save
    xlWB.Close
    xlApp.Quit
    Set xlWB = Nothing
    Set xlApp = Nothing
End Sub

Private Sub WriteColumn(Worksheet As Object, Col As Long, Optional rs As Recordset, Optional rowCount As Long, Optional CsvString As String)
    Dim Row As Long
    Row = 1
    
    If CsvString <> "" Then
        rowCount = Parse(CsvString)
        For Row = 1 To rowCount
            Call FProgress.Progress(, , Row, rowCount)
            Worksheet.Cells(Row + 1, Col).value = Parse(CsvString, Row)
        Next
    Else
        While Not rs.EOF
            If rowCount <> 0 Then Call FProgress.Progress(, , Row, rowCount)
            Row = Row + 1
            Worksheet.Cells(Row, Col).value = "'" & rs(0)
            rs.MoveNext
        Wend
    End If
    
End Sub
