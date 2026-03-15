Attribute VB_Name = "MPipeline"
Option Explicit
Option Compare Text
Const SRCFILE = "MPipeline::"

Public Sub ImportPipelineProducts()
On Error GoTo eh

    Dim X As New PipelineWrapper.PipelineWrapper
    Dim xml As String
    Dim s As String
    Dim i As Integer
    Dim Count As Integer
    Dim session As String
    Dim rs As Recordset
    Dim page As Integer
    
    If MsgBox("Import Pipeline Products?" & vbCrLf & vbCrLf & "This is a large operation and can take several minutes. Your screen may not respond while the import is running. Please be patient.", vbInformation + vbOKCancel, App.ProductName) = vbCancel Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Call X.Connect(HFApp.Options.ValueByName("CGVisionsRootURI"), HFApp.Options.ValueByName("CGVisionsUID"), HFApp.Options.ValueByName("CGVisionsPWD"))
    Call X.ReadProducts(10000)
    
    session = CreateGUID()
    
    For page = 0 To X.PageCount - 1
        xml = X.GetPage(page)
        
        xml = CleanXML(xml)
    
        'log response
        i = FreeFile()
        s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & VB.App.EXEName & "\pipeline.products.xml")
        Call CreatePath("", FilePath(s))
        Open PathAppend(s) For Output As #i
        Print #i, xml
        Close #i
                    
        'stage
        s = "exec BIM_StageProducts " & DbQuote(Str, session) & "," & DbQuote(Str, HFApp.LoginID) & "," & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, xml)
        Call HFApp.SqlExec(s, dbHomefront)
    
    Next

    'validate uniqueness
    s = ""
    s = s & "select hf_phase,hf_item,productdescription" & vbCrLf
    s = s & "from importedbimproductsandphases" & vbCrLf
    s = s & "where sessionid=" & DbQuote(Str, session) & vbCrLf
    s = s & "and productname in(" & vbCrLf
    s = s & "select productname" & vbCrLf
    s = s & "from importedbimproductsandphases" & vbCrLf
    s = s & "where sessionid=" & DbQuote(Str, session) & vbCrLf
    s = s & "and hf_phase<>''" & vbCrLf
    s = s & "and hf_itemnumber<>''" & vbCrLf
    s = s & "group by productname" & vbCrLf
    s = s & "having count(*)>1" & vbCrLf
    s = s & ") order by 1,2,3" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    s = ""
    i = 0
    While Not rs.EOF And i < 50
        i = i + 1
        s = s & rs("hf_phase") & "/" & rs("hf_item") & "   " & rs("productdescription") & vbCrLf
        rs.MoveNext
    Wend
    If s <> "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Unable to import products. These product numbers are not unique. These products must be updated in Pipeline before importing." & vbCrLf & vbCrLf & s, vbExclamation, App.ProductName
        s = "delete importedbimproductsandphases where sessionid=" & DbQuote(Str, session)
        Call HFApp.SqlExec(s, dbHomefront)
        Exit Sub
    End If



    'preview
    s = ""
    s = s & "select i.Phase,i.Item,i.Description" & vbCrLf
    s = s & "from tblphaseItem i" & vbCrLf
    s = s & "left outer join importedbimproductsandphases x on " & vbCrLf
    s = s & "   x.hf_phase=i.phase and " & vbCrLf
    s = s & "   x.hf_item=i.item and " & vbCrLf
    s = s & "   x.divisionid=i.divisionid and " & vbCrLf
    s = s & "   x.SessionID=" & DbQuote(Str, session) & vbCrLf
    s = s & "" & vbCrLf
    s = s & "where i.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and x.divisionid is null" & vbCrLf
    s = s & "order by 2,1" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    s = ""
    Count = 0
    While Not rs.EOF
        Count = Count + 1
        s = s & rs("phase") & "/" & rs("item") & " " & rs("description") & vbCrLf
        rs.MoveNext
    Wend

    If Count = 0 Then
        s = "This will synchronize your item db with Pipelines product list."
        i = vbQuestion
    Else
        s = "This will synchronize your item db with Pipelines product list." & vbCrLf & vbCrLf & "WARNING! " & IIf(Count = 1, "This item", "These " & Count & " items") & " will be removed from HomeFronts ItemDB." & vbCrLf & vbCrLf & s
        i = vbExclamation
    End If

    Screen.MousePointer = vbNormal
    If MsgBox(s, vbOKCancel + i, App.ProductName) = vbCancel Then
        s = "delete importedbimproductsandphases where sessionid=" & DbQuote(Str, session)
        Call HFApp.SqlExec(s, dbHomefront)
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    s = "exec BIM_ImportProducts " & DbQuote(Str, session)
    Call HFApp.SqlExec(s, dbHomefront)

    Screen.MousePointer = vbNormal
    MsgBox "Import Complete", vbInformation, App.ProductName

Exit Sub
eh: Call errHandler(SRCFILE & "ImportPipelineProducts", s)
End Sub


Public Sub ImportPipelineGlobalOptions()
On Error GoTo eh

    Dim X As New PipelineWrapper.PipelineWrapper
    Dim xml As String
    Dim s As String
    Dim i As Integer
    Dim r As Long
    Dim Count As Integer
    Dim session As String
    Dim rs As Recordset
    Dim Divisions As String
    
    If MsgBox("Import Global Options from Pipeline?" & vbCrLf & vbCrLf & "This is a large operation and can take several minutes. Your screen may not respond while the import is running. Please be patient.", vbInformation + vbOKCancel, App.ProductName) = vbCancel Then Exit Sub
    
    'choose divisions
    If HFApp.Options.ValueByName("CorporateEstimating") = "True" Then
        s = "select  d.DivisionID,d.DivisionCode, d.DivisionName from Divisions d left outer join divisionusers u on u.DivisionID = d.DivisionID where u.userID = " & DbQuote(Str, HFApp.LoginID) & " or u.userid is null order by 2"
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Select Division", s, HFApp.DivisionID, True, False, , "DivisionID", True) Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        For r = 1 To FPickList.SelectedItems
            Divisions = Divisions & "," & FPickList.SelectedItem("DivisionID", r)
        Next
        Divisions = Mid(Divisions, 2)
    Else
        Divisions = HFApp.DivisionID
    End If
    
    
    Screen.MousePointer = vbHourglass
    Call X.Connect(HFApp.Options.ValueByName("CGVisionsRootURI"), HFApp.Options.ValueByName("CGVisionsUID"), HFApp.Options.ValueByName("CGVisionsPWD"))
    xml = X.GetPlans()
    
    'log response
    i = FreeFile()
    s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & VB.App.EXEName & "\pipeline.plans.xml")
    Call CreatePath("", FilePath(s))
    Open PathAppend(s) For Output As #i
    Print #i, xml
    Close #i
    
    
    'choose plans
    s = "exec BIM_Models " & DbQuote(Str, xml)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    If Not rs.EOF Then
        
        xml = X.GetGlobalProducts("0000")

        'log response
        i = FreeFile()
        s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & VB.App.EXEName & "\pipeline.planproducts.xml")
        Call CreatePath("", FilePath(s))
        Open PathAppend(s) For Output As #i
        Print #i, xml
        Close #i

        s = "exec BIM_ImportModelBOM"
        s = s & " " & DbQuote(Str, Divisions)
        s = s & "," & DbQuote(Num, 1)  'globals only
        s = s & "," & DbQuote(Str, "") 'community
        s = s & "," & DbQuote(Str, "") 'series
        s = s & "," & DbQuote(Num, 0)  'beds
        s = s & "," & DbQuote(Num, 0)  'baths
        s = s & "," & DbQuote(Num, 0)  'sqft
        s = s & "," & DbQuote(Str, HFApp.LoginID)
        s = s & "," & DbQuote(Str, xml)
        
        Call HFApp.SqlExec(s)
    
    End If

    Call ImportPipelineOptionCategories(Divisions)

    Screen.MousePointer = vbNormal
    MsgBox "Import Complete", vbInformation, App.ProductName

Exit Sub
eh: Call errHandler(SRCFILE & "ImportPipelineGlobalOptions", s)
End Sub



Public Sub ImportPipelineModelsAndOptions()
    If HFApp.Options.ValueByName("BIMCommunitySpecificPlansAndOptions") = "true" Then
        Call ImportPipelineModelsAndOptionsCOMMUNITY
    Else
        Call ImportPipelineModelsAndOptionsGLOBAL
    End If
End Sub


Public Sub ImportPipelineModelsAndOptionsCOMMUNITY()
 On Error GoTo eh


    Dim X As New PipelineWrapper.PipelineWrapper
    Dim xml As String
    Dim s As String
    Dim i As Integer
    Dim j As Long
    Dim r As Long
    Dim Count As Integer
    Dim session As String
    Dim rs As Recordset
    Dim community As String
    Dim Divisions As String
    
    
    'choose divisions
    If HFApp.Options.ValueByName("CorporateEstimating") = "True" Then
        s = "select  d.DivisionID,d.DivisionCode, d.DivisionName from Divisions d left outer join divisionusers u on u.DivisionID = d.DivisionID where u.userID = " & DbQuote(Str, HFApp.LoginID) & " or u.userid is null order by 2"
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Select Division", s, HFApp.DivisionID, True, False, , "DivisionID", True) Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        For r = 1 To FPickList.SelectedItems
            Divisions = Divisions & "," & FPickList.SelectedItem("DivisionID", r)
        Next
        Divisions = Mid(Divisions, 2)
    Else
        Divisions = HFApp.DivisionID
    End If

    
    'choose community
    s = ""
    s = s & "select c.Area Community,c.Description" & vbCrLf
    s = s & "from tbllocality c" & vbCrLf
    s = s & "join divisioncommunities d on c.area=d.community" & vbCrLf
    s = s & "where d.divisionid=" & HFApp.DivisionID & vbCrLf
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Community", s) Then Exit Sub
    community = FPickList.SelectedItem("Community")
    If community = "" Then Exit Sub
    
    
    'connect to pipeline and get plans
    If MsgBox("Import Models from Pipeline?" & vbCrLf & vbCrLf & "This is a large operation and can take a long time. Your screen may not respond while the import is running. Please be patient.", vbInformation + vbOKCancel, App.ProductName) = vbCancel Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Call X.Connect(HFApp.Options.ValueByName("CGVisionsRootURI"), HFApp.Options.ValueByName("CGVisionsUID"), HFApp.Options.ValueByName("CGVisionsPWD"))
    xml = X.GetCommunityPlans(community)
    'log response
    i = FreeFile()
    s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & VB.App.EXEName & "\pipeline.plans.xml")
    Call CreatePath("", FilePath(s))
    Open PathAppend(s) For Output As #i
    Print #i, xml
    Close #i
    
    
    'choose plans
    s = "exec BIM_CommunityModels " & DbQuote(Str, xml)
    Screen.MousePointer = vbNormal
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Plans", s, , , , , , True) Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    
Dim t As Single
Dim Y As String
t = Timer()
    For r = 1 To FPickList.SelectedItems
    
t = Timer()
        xml = X.GetPlanProducts(FPickList.SelectedItem("PlanNo", r), community)
Y = Y & "GetPlanProducts " & format(Timer() - t, "0.000") & vbCrLf
        
        'log response
        i = FreeFile()
        s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & VB.App.EXEName & "\pipeline.planproducts.xml")
        Call CreatePath("", FilePath(s))
        Open PathAppend(s) For Output As #i
        Print #i, xml
        Close #i

        s = "exec BIM_ImportModelBOM"
        s = s & " " & DbQuote(Str, HFApp.DivisionID)
        s = s & "," & DbQuote(Num, 0) 'globals only
        s = s & "," & DbQuote(Str, community)
        s = s & "," & DbQuote(Str, FPickList.SelectedItem("Series", r))
        s = s & "," & DbQuote(Num, FPickList.SelectedItem("Beds", r))
        s = s & "," & DbQuote(Num, FPickList.SelectedItem("Baths", r))
        s = s & "," & DbQuote(Num, FPickList.SelectedItem("SqFt", r))
        s = s & "," & DbQuote(Str, HFApp.LoginID)
        s = s & "," & DbQuote(Str, xml)
t = Timer()
        'Call HFApp.SqlExec(s)
        Set rs = HFApp.SqlExec(s)
Y = Y & "ImportPlanProducts " & format(Timer() - t, "0.000") & vbCrLf
'Y = Y & rs(0) & vbCrLf
        
    
    Next
    
t = Timer()
    Call ImportPipelineOptionCategories(Divisions)
Y = Y & "ImportOptionCategories " & format(Timer() - t, "0.000") & vbCrLf
'MsgBox Y, vbInformation, App.ProductName

    Screen.MousePointer = vbNormal
    MsgBox "Import Complete", vbInformation, App.ProductName

Exit Sub
eh: Call errHandler(SRCFILE & "ImportPipelineModelsAndOptionsCOMMUNITY", s)
End Sub




Public Sub ImportPipelineModelsAndOptionsGLOBAL()
 On Error GoTo eh


    Dim X As New PipelineWrapper.PipelineWrapper
    Dim xml As String
    Dim s As String
    Dim i As Integer
    Dim j As Long
    Dim r As Long
    Dim Count As Integer
    Dim session As String
    Dim rs As Recordset
    Dim Divisions As String
    
    If MsgBox("Import Models from Pipeline?" & vbCrLf & vbCrLf & "This is a large operation and can take a long time. Your screen may not respond while the import is running. Please be patient.", vbInformation + vbOKCancel, App.ProductName) = vbCancel Then Exit Sub
    
    'choose divisions
    If HFApp.Options.ValueByName("CorporateEstimating") = "True" Then
        s = "select  d.DivisionID,d.DivisionCode, d.DivisionName from Divisions d left outer join divisionusers u on u.DivisionID = d.DivisionID where u.userID = " & DbQuote(Str, HFApp.LoginID) & " or u.userid is null order by 2"
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Select Division", s, HFApp.DivisionID, True, False, , "DivisionID", True) Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        For r = 1 To FPickList.SelectedItems
            Divisions = Divisions & "," & FPickList.SelectedItem("DivisionID", r)
        Next
        Divisions = Mid(Divisions, 2)
    Else
        Divisions = HFApp.DivisionID
    End If

    
    Screen.MousePointer = vbHourglass
    Call X.Connect(HFApp.Options.ValueByName("CGVisionsRootURI"), HFApp.Options.ValueByName("CGVisionsUID"), HFApp.Options.ValueByName("CGVisionsPWD"))
    xml = X.GetPlans()
    
    'log response
    i = FreeFile()
    s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & VB.App.EXEName & "\pipeline.plans.xml")
    Call CreatePath("", FilePath(s))
    Open PathAppend(s) For Output As #i
    Print #i, xml
    Close #i
    
    
    'choose plans
    s = "exec BIM_Models " & DbQuote(Str, xml)
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Plans", s, , , , , , True) Then
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
    
    
Dim t As Single
Dim Y As String
t = Timer()
        For r = 1 To FPickList.SelectedItems
        
t = Timer()
            xml = X.GetPlanProducts(FPickList.SelectedItem("PlanNo", r), "0000")
Y = Y & "GetPlanProducts " & format(Timer() - t, "0.000") & vbCrLf
            
            'log response
            i = FreeFile()
            s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & VB.App.EXEName & "\pipeline.planproducts.xml")
            Call CreatePath("", FilePath(s))
            Open PathAppend(s) For Output As #i
            Print #i, xml
            Close #i


            s = "exec BIM_ImportModelBOM"
            s = s & " " & DbQuote(Str, Divisions)
            s = s & "," & DbQuote(Num, 0) 'globals only
            s = s & "," & DbQuote(Str, "") 'community
            s = s & "," & DbQuote(Str, FPickList.SelectedItem("Series", r))
            s = s & "," & DbQuote(Num, FPickList.SelectedItem("Beds", r))
            s = s & "," & DbQuote(Num, FPickList.SelectedItem("Baths", r))
            s = s & "," & DbQuote(Num, FPickList.SelectedItem("SqFt", r))
            s = s & "," & DbQuote(Str, HFApp.LoginID)
            s = s & "," & DbQuote(Str, xml)
t = Timer()
            'Call HFApp.SqlExec(s)
            Set rs = HFApp.SqlExec(s)
Y = Y & "ImportPlanProducts " & format(Timer() - t, "0.000") & vbCrLf
'Y = Y & rs(0) & vbCrLf
            
        
        Next
    End If
    
t = Timer()
    Call ImportPipelineOptionCategories(Divisions)
Y = Y & "ImportOptionCategories " & format(Timer() - t, "0.000") & vbCrLf
MsgBox Y, vbInformation, App.ProductName

    Screen.MousePointer = vbNormal
    MsgBox "Import Complete", vbInformation, App.ProductName

Exit Sub
eh: Call errHandler(SRCFILE & "ImportPipelineModelsAndOptionsGLOBAL", s)
End Sub



Public Sub ImportPipelineOptionCategories(Divisions As String)
On Error GoTo eh

    Dim X As New PipelineWrapper.PipelineWrapper
    Dim xml As String
    Dim s As String
    Dim i As Integer
    Dim j As Long
    Dim r As Long
    Dim Count As Integer
    Dim session As String
    Dim rs As Recordset

    
    Call X.Connect(HFApp.Options.ValueByName("CGVisionsRootURI"), HFApp.Options.ValueByName("CGVisionsUID"), HFApp.Options.ValueByName("CGVisionsPWD"))
    xml = X.GetOptions()

    'log response
    i = FreeFile()
    s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & VB.App.EXEName & "\pipeline.options.xml")
    Call CreatePath("", FilePath(s))
    Open PathAppend(s) For Output As #i
    Print #i, xml
    Close #i

    s = "exec BIM_ImportOptions"
    s = s & " " & DbQuote(Str, Divisions)
    s = s & "," & DbQuote(Str, HFApp.LoginID)
    s = s & "," & DbQuote(Str, xml)
    
    Call HFApp.SqlExec(s)

Exit Sub
eh: Call errHandler(SRCFILE & "ImportPipelineOptionCategories", s)
End Sub


