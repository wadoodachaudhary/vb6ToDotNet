VERSION 5.00
Object = "{39FDA063-61BA-11D2-AD84-00105A17B608}#1.0#0"; "DartFtp.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FBuilder1440 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transfering Data"
   ClientHeight    =   1575
   ClientLeft      =   4875
   ClientTop       =   2190
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFail 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   180
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
      Begin VB.Label lblMore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1080
         TabIndex        =   7
         Top             =   450
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Click"
         Height          =   240
         Left            =   690
         TabIndex        =   6
         Top             =   450
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Transfer failed."
         Height          =   240
         Left            =   690
         TabIndex        =   5
         Top             =   210
         Width           =   1560
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "FBuilder1440a.frx":0000
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "for more details..."
         Height          =   240
         Left            =   1410
         TabIndex        =   2
         Top             =   450
         Width           =   1365
      End
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   915
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   1614
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   275
      FullHeight      =   61
   End
   Begin DartFtpCtl.Ftp Ftp1 
      Left            =   3120
      OleObjectBlob   =   "FBuilder1440a.frx":08CA
      Top             =   180
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   210
      TabIndex        =   3
      Tag             =   "file123.zip from ftp://ftp.yoohoo.com"
      Top             =   1170
      Width           =   45
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing..."
      Height          =   240
      Left            =   210
      TabIndex        =   1
      Tag             =   "file123.zip from ftp://ftp.yoohoo.com"
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "FBuilder1440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FBuilder1440::"

        
Public Function ReadFromWeb(DownloadNew As Boolean) As Boolean
On Error GoTo eh
    Dim i As Integer
    Dim s As String
        
    If DownloadNew = False Then
        If vbCancel = MsgBox("Holding the CTRL key prevents Precision Builder from retrieving new data from the sales center. It will simply reprocess the data that was most recently downloaded. This is only beneficial during system setup and integration testing." & vbCrLf & vbCrLf & "Is this want you want to do?", vbQuestion + vbOKCancel, App.ProductName) Then
            Exit Function
        End If
    End If

    Dim SRCFILE As String
    Dim dstFile As String
    Dim trnFile As String
        
    Dim xslt As New MSXML2.XSLTemplate30
    Dim xslDoc As New MSXML2.FreeThreadedDOMDocument30
    Dim xmlDoc As New MSXML2.DOMDocument30
    Dim xslProc As IXSLProcessor
    
Call SetCaptions("Initializing...", "")
    
    Call CreatePath("", PathAppend(AppWorkingFolder, "1440"))
    SRCFILE = PathAppend(AppWorkingFolder, "1440", "1440Export.xml")
    dstFile = PathAppend(AppWorkingFolder, "1440", "HFImport.xml")
    trnFile = PathAppend(AppWorkingFolder, "1440", "transform.xsl")
    
    If FileExists(dstFile) Then Kill dstFile
    
    'copy transform to working folder for debuging'
    If FileExists(trnFile) Then Kill trnFile
    Call FileCopy(PathAppend(App.path, "1440transform.xsl"), trnFile)
    
    If DownloadNew Then
        
        Call ConnectFTP
Call SetCaptions("Downloading information...", "")
        
        'download src to working folder
        If FileExists(SRCFILE) Then Kill SRCFILE
                    
        Status "Directory = ""exports"""
        Ftp1.Directory = "exports"
        
        Status "GetFile(""BuilderExport.xml"", """ & SRCFILE & """)"
        Call Ftp1.Retrieve("BuilderExport.xml", SRCFILE)
        Status "DeleteFile(""BuilderExport.xml"")"
        Call Ftp1.Delete("BuilderExport.xml")
        
    End If
    
    
    If Not FileExists(SRCFILE) Then
        Unload Me
        Exit Function
    End If
Call SetCaptions("Processing...", "")
    
    
    
    'load source and transform.
    xslDoc.async = False
    xslDoc.Load trnFile
    If (xslDoc.parseError.errorCode <> 0) Then Err.Raise xslDoc.parseError.errorCode, , xslDoc.parseError.reason & "line: " & xslDoc.parseError.Line & ", Char: " & xslDoc.parseError.linepos & vbCrLf & Trim(xslDoc.parseError.srcText)
    Set xslt.stylesheet = xslDoc
    xmlDoc.async = False
    xmlDoc.Load SRCFILE
    If (xmlDoc.parseError.errorCode <> 0) Then Err.Raise xmlDoc.parseError.errorCode, , xslDoc.parseError.reason & "line: " & xslDoc.parseError.Line & ", Char: " & xslDoc.parseError.linepos & vbCrLf & Trim(xslDoc.parseError.srcText)
        
    'perform transform
    Set xslProc = xslt.createProcessor()
    xslProc.input = xmlDoc
    xslProc.Transform
    
    'save results to temp file for debug purposes
    i = FreeFile
    Open dstFile For Output As i
    s = xslProc.output
    Print #i, s
    Close i
    
    'send results to stored proc
    s = "exec importcustomer " & DbQuote(Str, s)
    HFApp.SqlExec s, dbHomeFront
 
    ReadFromWeb = True
    Unload Me
Exit Function
eh:
Select Case True
    Case Err.Description = "550 BuilderExport.xml: cannot find specified file" & vbCrLf
    ReadFromWeb = True
    
    Case Err.Description = "530 Invalid userid/password" & vbCrLf
        MsgBox "Invalid security context or password." & vbCrLf & vbCrLf & "Please check your system settings.", vbCritical, App.ProductName

    Case Else
        MsgBox Err.Description, vbCritical, App.ProductName
End Select
Unload Me
End Function



Public Sub WriteToWeb()
On Error GoTo eh
    Dim d As Long
    Dim r As Long
    
    Dim ld As Date
    Dim ud As Date
    Dim s  As String
    
    Dim iFile As Integer
    Dim sfile As String
    
    Dim bAppendUOM As Boolean
    Dim bUploadByCommunity As Boolean
        
    Dim divID   As Long
    Dim divCode As String
    Dim divName As String
    Dim divChgd As Boolean
    Dim comID   As String
    
    Dim rs As Recordset
    
    
'------------------------------------------------------------------------------------
' NOTES...
'   1. B1440's construction status must be greater than zero. When posting we will
'      increment ours by one.
'
'------------------------------------------------------------------------------------
    
    
    If Not ValidateData() Then
        Unload Me
        Exit Sub
    End If
    
    
    bAppendUOM = HFApp.Options.ValueByName("AppendUOMtoOptionDesc") = "True"
    bUploadByCommunity = HFApp.Options.ValueByName("WebUploadByCommunity") <> "False"

    If bUploadByCommunity Then
        s = ""
        s = s & "select d.divisionid" & vbCrLf
        s = s & ",d.modifieddate" & vbCrLf
        s = s & ",d.divisioncode Division" & vbCrLf
        s = s & ",d.divisionname DivisionDesc" & vbCrLf
        s = s & ",c.area Community" & vbCrLf
        s = s & ",c.description CommunityDesc" & vbCrLf
        s = s & " from divisions d" & vbCrLf
        s = s & " join divisioncommunities dc on(d.divisionid=dc.divisionid)" & vbCrLf
        s = s & " join tbllocality c on(c.area=dc.community)" & vbCrLf
        If Not FPickList.Choose(HFApp.Databases(dbHomeFront), "Community", s, , , , , "divisionid,modifieddate", True) Then Exit Sub
    Else
        s = ""
        s = s & "select d.divisionid" & vbCrLf
        s = s & ",d.modifieddate" & vbCrLf
        s = s & ",d.divisioncode Division" & vbCrLf
        s = s & ",d.divisionname DivisionDesc" & vbCrLf
        s = s & ",'' Community" & vbCrLf
        s = s & " from divisions d" & vbCrLf
        If Not FPickList.Choose(HFApp.Databases(dbHomeFront), "Division", s, , , , , "divisionid,modifieddate,Community", True) Then Exit Sub
    End If
    
    
    Call ConnectFTP
    
    For d = 1 To FPickList.SelectedItems
        
        ud = Now() 'get timestamp before we start loading files up
        divID = Val(FPickList.SelectedItem("DivisionID", d))
        divCode = FPickList.SelectedItem("Division", d)
        divName = FPickList.SelectedItem("DivisionDesc", d)
        comID = FPickList.SelectedItem("Community", d)
        ld = LastUpdate("1440:" & divID & ":" & comID)
        s = FPickList.SelectedItem("modifieddate", d)
        If IsDate(s) Then divChgd = DateValue(s) >= ld
        
        
        
        '------------------------------------------------
        ' open tmp file
        '------------------------------------------------
        Call CreatePath("", PathAppend(AppWorkingFolder, "1440"))
        sfile = PathAppend(AppWorkingFolder, "1440", CleanFileName(UCase(divCode & IIf(bUploadByCommunity, "_" & comID, "")) & ".xml"))
        iFile = FreeFile
        
        On Error Resume Next
        Kill sfile
        On Error GoTo eh
        Open sfile For Binary As iFile
        
        
Call SetCaptions("Preparing division " & divCode & IIf(bUploadByCommunity, ", community " & comID & "...", "..."), "")
    
    'A
        Put #iFile, , CleanXML("<?xml version='1.0' encoding=""UTF-8"" standalone=""no"" ?>")
        Put #iFile, , CleanXML("<!DOCTYPE hasp SYSTEM ""file:/N:/opt/hasp/xml/hasp.dtd"">")
        Put #iFile, , CleanXML("<hasp>")
        Put #iFile, , CleanXML("<builder action=""update""" & " buildercode=" & XmlQuote(HFApp.Options(BuilderCode), True) & " adminname=" & XmlQuote(HFApp.Options(WebAdminUser), True) & " country=" & XmlQuote(HFApp.Options(Country), True) & ">")
        Put #iFile, , CleanXML("<name>" & XmlQuote(HFApp.Options(CompanyName)) & "</name>")
    'B
        Put #iFile, , CleanXML(SqlXML("select address1,address2,city,upper(province) state,zip postalcode,isnull(nullif(upper(country),''),'US') country from system_setup address for xml auto, elements"))
    'C
        Put #iFile, , CleanXML(SqlXML("select contactperson name,phone,fax,email from system_setup contact for xml auto, elements"))
    
    'D
        Put #iFile, , CleanXML("<division action=""add"">")
        Put #iFile, , CleanXML("<divisioncode>" & XmlQuote("" & divCode) & "</divisioncode>")
        Put #iFile, , CleanXML("<name>" & XmlQuote("" & divName) & "</name>")
'EE
Call SetCaptions("", "Processing Option Types...")
        s = ""
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'       [optiontype!1!action]" & vbCrLf
        s = s & "      ,major_group [optiontype!1!optiontypecode!element]" & vbCrLf
        s = s & "      ,description [optiontype!1!description!element]" & vbCrLf
        s = s & "from tblmajorgroups" & vbCrLf
        If Not divChgd Then s = s & "where modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "for xml explicit" & vbCrLf
        Put #iFile, , CleanXML(SqlXML(s))
'E
Call SetCaptions("", "Processing Option Categories...")
        s = ""
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'       [optioncategory!1!action]" & vbCrLf
        s = s & "      ,category    [optioncategory!1!optioncategorycode!element]" & vbCrLf
        s = s & "      ,description [optioncategory!1!description!element]" & vbCrLf
        s = s & "from tblcategories" & vbCrLf
        If Not divChgd Then s = s & "where modifieddate>=" & DbQuote(DateTime, ld)
        s = s & "for xml explicit" & vbCrLf
        Put #iFile, , SqlXML(s)
'F
Call SetCaptions("", "Processing Room Names...")
        s = ""
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'       [room!1!action]" & vbCrLf
        s = s & "      ,area        [room!1!roomcode!element]" & vbCrLf
        s = s & "      ,description [room!1!roomname!element]" & vbCrLf
        s = s & "from tblareas" & vbCrLf
        If Not divChgd Then s = s & "where modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "for xml explicit" & vbCrLf
        Put #iFile, , SqlXML(s)
    
'G -- construction statuses. create a dummy record if there are none.
Call SetCaptions("", "Processing Construction Phases...")
        s = ""
        s = ""
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'     [phase!1!action]" & vbCrLf
        s = s & "      ,1         [phase!1!phasecode!element]" & vbCrLf
        s = s & "      ,'construction' [phase!1!description!element]" & vbCrLf
        s = s & "      ,1         [phase!1!phasesequence!element]" & vbCrLf
        s = s & "from system_setup left outer join tbljobstatus on(1=1) where tbljobstatus.job_status is null" & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'         [phase!1!action]" & vbCrLf
        s = s & "      ,job_status+1  [phase!1!phasecode!element]" & vbCrLf
        s = s & "      ,description [phase!1!description!element]" & vbCrLf
        s = s & "      ,(select count(*) from tbljobstatus b where b.job_status<=tbljobstatus.job_status) [phase!1!phasesequence!element]" & vbCrLf
        s = s & "from tbljobstatus" & vbCrLf
        If Not divChgd Then s = s & "where modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "for xml explicit" & vbCrLf
        Put #iFile, , SqlXML(s)

'H
Call SetCaptions("", "Processing Options...")
        s = ""
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'                         [option!1!action]" & vbCrLf
        s = s & "      ,opt                           [option!1!optioncode!element]" & vbCrLf
        s = s & "      ,major_group                   [option!1!optiontypecode!element]" & vbCrLf
        s = s & "      ,category                      [option!1!optioncategorycode!element]" & vbCrLf
        s = s & "      ,construction_cut_off+1        [option!1!phasecode!element]" & vbCrLf
                
        If bAppendUOM Then
            s = s & "      ,left(description,242)+isnull(' ('+nullif(uom,'')+')','')         [option!1!shortdescription!element]" & vbCrLf
        Else
            s = s & "      ,left(description,255)         [option!1!shortdescription!element]" & vbCrLf
        End If
        
        's = s & "      ,cast(comments as varchar)     [option!1!longdescription!element]" & vbCrLf
        s = s & "      ,''                            [option!1!longdescription!element]" & vbCrLf
        
        s = s & "      ,cast(max(price) as varchar)   [option!1!price!element]" & vbCrLf
        s = s & "      ,dbo.fn_Parse1440RoomsXML(location,';') [option!1!rooms!element]" & vbCrLf
        s = s & "from tblOptions" & vbCrLf
        If Not divChgd Then s = s & "where modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "group by opt,major_group,uom,category,construction_cut_off,description,cast(comments as varchar),dbo.fn_Parse1440RoomsXML(location,';')" & vbCrLf
        s = s & "union" & vbCrLf
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'                         [option!1!action]" & vbCrLf
        s = s & "      ,opt                           [option!1!optioncode!element]" & vbCrLf
        s = s & "      ,major_group                   [option!1!optiontypecode!element]" & vbCrLf
        s = s & "      ,category                      [option!1!optioncategorycode!element]" & vbCrLf
        s = s & "      ,construction_cut_off+1        [option!1!phasecode!element]" & vbCrLf
        If bAppendUOM Then
            s = s & "      ,left(description,242)+isnull(' ('+nullif(uom,'')+')','')         [option!1!shortdescription!element]" & vbCrLf
        Else
            s = s & "      ,left(description,255)         [option!1!shortdescription!element]" & vbCrLf
        End If
        's = s & "      ,cast(comments as varchar)     [option!1!longdescription!element]" & vbCrLf
        s = s & "      ,''                            [option!1!longdescription!element]" & vbCrLf
        s = s & "      ,cast(max(price) as varchar)   [option!1!price!element]" & vbCrLf
        s = s & "      ,dbo.fn_Parse1440RoomsXML(location,';') [option!1!rooms!element]" & vbCrLf
        s = s & "from tblglobalOptions" & vbCrLf
        If Not divChgd Then s = s & "where modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "group by opt,major_group,uom,category,construction_cut_off,description,cast(comments as varchar),dbo.fn_Parse1440RoomsXML(location,';')" & vbCrLf
        s = s & "union" & vbCrLf
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'                         [option!1!action]" & vbCrLf
        s = s & "      ,opt                           [option!1!optioncode!element]" & vbCrLf
        s = s & "      ,major_group                   [option!1!optiontypecode!element]" & vbCrLf
        s = s & "      ,category                      [option!1!optioncategorycode!element]" & vbCrLf
        s = s & "      ,1                             [option!1!phasecode!element]" & vbCrLf
        If bAppendUOM Then
            s = s & "      ,left(description,242)+isnull(' ('+nullif(uom,'')+')','')         [option!1!shortdescription!element]" & vbCrLf
        Else
            s = s & "      ,left(description,255)         [option!1!shortdescription!element]" & vbCrLf
        End If
        's = s & "      ,cast(comments as varchar)     [option!1!longdescription!element]" & vbCrLf
        s = s & "      ,''                            [option!1!longdescription!element]" & vbCrLf
        s = s & "      ,cast(max(item1) as varchar)   [option!1!price!element]" & vbCrLf
        s = s & "      ,dbo.fn_Parse1440RoomsXML(location,';') [option!1!rooms!element]" & vbCrLf
        s = s & "from tbldcOptions" & vbCrLf
        
        If Not divChgd Then s = s & "where modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        
        s = s & "group by opt,major_group,uom,category,description,cast(comments as varchar),dbo.fn_Parse1440RoomsXML(location,';')" & vbCrLf
        s = s & "for xml explicit" & vbCrLf
        Put #iFile, , SqlXML(s)
    
'I
Call SetCaptions("", "Processing Series...")
        s = ""
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'       [series!1!action]" & vbCrLf
        s = s & "      ,series      [series!1!seriescode!element]" & vbCrLf
        s = s & "      ,description [series!1!seriesname!element]" & vbCrLf
        s = s & "      ,description [series!1!description!element]" & vbCrLf
        s = s & "from tblseries" & vbCrLf
        If Not divChgd Then s = s & "where modifieddate>=" & DbQuote(DateTime, ld)
        s = s & "for xml explicit" & vbCrLf
        Put #iFile, , SqlXML(s)
    
'J
Call SetCaptions("", "Processing Models...")
        s = ""
        s = s & "select 1 tag" & vbCrLf
        s = s & "      ,0 parent" & vbCrLf
        s = s & "      ,'add'            [plan!1!action]" & vbCrLf
        s = s & "      ,m.model          [plan!1!plancode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(m.elevation,''),m.model)      [plan!1!elevationcode!element]" & vbCrLf
        s = s & "      ,m.series         [plan!1!seriescode!element]" & vbCrLf
        s = s & "      ,m.description    [plan!1!planelevationname!element]" & vbCrLf
        s = s & "      ,m.description    [plan!1!description!element]" & vbCrLf
        s = s & "      ,case when m.style in ('Single-Family','Condo','Townhouse') then m.style else 'Single-Family' end [plan!1!plantype!element]" & vbCrLf
        s = s & "      ,cast(m.modelsize as varchar) [plan!1!squarefootage!element]" & vbCrLf
        s = s & "      ,cast(m.base_house as varchar)     [plan!1!baseprice!element]" & vbCrLf
        s = s & "from tblmodels m" & vbCrLf
        s = s & "where isnull(m.inactive,0) = 0" & vbCrLf
        If Not divChgd Then
            s = s & "and (isnull(m.SentTo1440,0)=0 OR modifieddate>=" & DbQuote(DateTime, ld) & ")" & vbCrLf
        End If
        s = s & "for xml explicit" & vbCrLf
        Put #iFile, , SqlXML(s)
        
        
'K
        Put #iFile, , CleanXML("<addontype action=""add"">")
        Put #iFile, , CleanXML("<divisionaddontypecode>PRKING</divisionaddontypecode>")
        Put #iFile, , CleanXML("<name>Parking Space</name>")
        Put #iFile, , CleanXML("<description>Parking Space</description>")
        Put #iFile, , CleanXML("</addontype>")
        
        Put #iFile, , CleanXML("<addontype action=""add"">")
        Put #iFile, , CleanXML("<divisionaddontypecode>STOR</divisionaddontypecode>")
        Put #iFile, , CleanXML("<name>Storage Cabinet</name>")
        Put #iFile, , CleanXML("<description>Storage Cabinet</description>")
        Put #iFile, , CleanXML("</addontype>")
        
'L
Call SetCaptions("", "Querying Community...")
        s = ""
        s = s & "select 10            tag         --community" & vbCrLf
        s = s & "      ,0             parent" & vbCrLf
        s = s & "      ,1             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,'add'         [community!10!action]" & vbCrLf
        s = s & "      ,c.area        [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,c.description+" & DbQuote(Str, " (" & divCode & ")") & "[community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,''            [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllocality c" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 20            tag" & vbCrLf
        s = s & "      ,10            parent" & vbCrLf
        s = s & "      ,1             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,'add'         [community!10!action]" & vbCrLf
        s = s & "      ,c.area        [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,c.description [community!10!name!element]" & vbCrLf
        s = s & "      ,c.address1    [address!20!address1!element]" & vbCrLf
        s = s & "      ,c.address2    [address!20!address2!element]" & vbCrLf
        s = s & "      ,c.city        [address!20!city!element]" & vbCrLf
        s = s & "      ,upper(c.province)    [address!20!state!element]" & vbCrLf
        s = s & "      ,c.zip         [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,c.county      [address!20!county!element]" & vbCrLf
        s = s & "      ,c.country     [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,''            [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllocality c" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 30            tag" & vbCrLf
        s = s & "      ,10            parent" & vbCrLf
        s = s & "      ,1             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,c.area        [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,case when c.UsesPhases=1 then 'enable' else 'disable' end     [releasepricing!30!action]" & vbCrLf
        s = s & "      ,isnull(min(cp.communityphase),'N/A') [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,''            [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllocality c" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "     left outer join communityphase cp on(c.area=cp.community)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "group by c.area,c.UsesPhases" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 40            tag" & vbCrLf
        s = s & "      ,10            parent" & vbCrLf
        s = s & "      ,4             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,cp.community  [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,'add'         [release!40!action]" & vbCrLf
        s = s & "      ,isnull(cp.communityphase,'') [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,cp.description [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from communityphase cp" & vbCrLf
        s = s & "     join tbllocality c on cp.community = c.area" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and cp.community=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 50            tag" & vbCrLf
        s = s & "      ,10            parent" & vbCrLf
        s = s & "      ,5             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,isnull(c.area,m.area) [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(p.communityphase,m.communityphase)  [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,'add'         [communityplan!50!action]" & vbCrLf
        s = s & "      ,isnull(m.model,'')       [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(m.elevation,''),isnull(m.model,''))       [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,isnull(m.series,'')      [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,nullif(m.communityphase,'')  [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,cast(m.base_house as varchar)  [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,'add'         [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tblmodels m" & vbCrLf
        s = s & "     join tbllocality c on (m.area=c.area or isnull(m.area,'')='')" & vbCrLf
        s = s & "     left outer join communityphase p on (p.community=isnull(c.area,m.area) and (m.communityphase=p.communityphase or isnull(m.communityphase,'')=''))" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all ----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 60             tag" & vbCrLf
        s = s & "      ,50             parent" & vbCrLf
        s = s & "      ,5             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,isnull(c.area,m.area) [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(p.communityphase,m.communityphase)  [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,isnull(m.model,'')       [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(m.elevation,''),isnull(m.model,''))       [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,isnull(m.series,'')      [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,nullif(m.communityphase,'')  [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tblmodels m" & vbCrLf
        s = s & "     left join tbllocality c on (m.area=c.area or isnull(m.area,'')='')" & vbCrLf
        s = s & "     left outer join communityphase p on (p.community=isnull(c.area,m.area) and (m.communityphase=p.communityphase or isnull(m.communityphase,'')=''))" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(m.inactive,0) = 0" & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 65             tag         --communityplan" & vbCrLf
        s = s & "      ,50             parent" & vbCrLf
        s = s & "      ,5             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,isnull(c.area,m.area) [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(p.communityphase,m.communityphase)  [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,isnull(m.model,'')       [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(m.elevation,''),isnull(m.model,''))       [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,isnull(m.series,'')      [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,nullif(m.communityphase,'')              [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,cast(isnull(m.base_house,0) as varchar)  [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,'add'         [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tblmodels m" & vbCrLf
        s = s & "     join tbllocality c on (m.area=c.area or isnull(m.area,'')='')" & vbCrLf
        s = s & "     left outer join communityphase p on (p.community=isnull(c.area,m.area) and (m.communityphase=p.communityphase or isnull(m.communityphase,'')=''))" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 70             tag         --options" & vbCrLf
        s = s & "      ,50             parent" & vbCrLf
        s = s & "      ,5             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,c.area        [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(p.communityphase,'')          [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,isnull(o.model,'')       [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(m.elevation,''),isnull(m.model,''))       [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,o.series      [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,o.communityphase [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,'add'         [communityoption!70!action]" & vbCrLf
        s = s & "      ,isnull(o.opt,'')         [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,o.construction_cut_off+1 [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,cast(o.price as varchar)       [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,case isnull(o.IncludedOption,0) when 0 then 'N' else 'Y' end [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,case isnull(o.IncludedOption,0) when 0 then '0' else case o.qty when 0 then '1' else cast(o.qty as varchar) end end [communityoption!70!standardquantity!element]" & vbCrLf
        's = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        's = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,o.Comments    [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tblmodels m" & vbCrLf
        s = s & "     join tbllocality c on (m.area=c.area or isnull(m.area,'')='')" & vbCrLf
        s = s & "     left outer join communityphase p on (p.community=isnull(c.area,m.area) and (m.communityphase=p.communityphase or isnull(m.communityphase,'')=''))" & vbCrLf
        s = s & "     join tbloptions o on(m.model=o.model and m.series=o.series and (o.area=isnull(c.area,m.area) or isnull(o.area,'')='') and (o.communityphase=isnull(p.communityphase,m.communityphase) or isnull(o.communityphase,'')=''))" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        If Not divChgd Then s = s & "  and (isnull(m.SentTo1440,0)=0 OR o.modifieddate>=" & DbQuote(DateTime, ld) & ")" & vbCrLf
        s = s & "  and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 70             tag         --global options" & vbCrLf
        s = s & "      ,50             parent" & vbCrLf
        s = s & "      ,5             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,c.area        [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(p.communityphase,m.communityphase)          [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,isnull(m.model,'')       [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(m.elevation,''),isnull(m.model,''))       [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,m.series      [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,p.communityphase [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,'add'         [communityoption!70!action]" & vbCrLf
        s = s & "      ,isnull(o.opt,'')         [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,o.construction_cut_off+1 [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,cast(o.price as varchar)       [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,case isnull(o.IncludedOption,0) when 0 then 'N' else 'Y' end [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,case isnull(o.IncludedOption,0) when 0 then '0' else case o.qty when 0 then '1' else cast(o.qty as varchar) end end   [communityoption!70!standardquantity!element]" & vbCrLf
        's = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        's = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,o.comments    [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tblmodels m" & vbCrLf
        s = s & "     join tbllocality c on (m.area=c.area or isnull(m.area,'')='')" & vbCrLf
        s = s & "     left outer join communityphase p on (p.community=isnull(c.area,m.area) and (m.communityphase=p.communityphase or isnull(m.communityphase,'')=''))" & vbCrLf
        s = s & "     join tblglobaloptions o on((o.community=isnull(c.area,m.area) or isnull(o.community,'')='') and (o.communityphase=isnull(p.communityphase,m.communityphase) or isnull(o.communityphase,'')=''))" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(m.inactive,0) = 0" & vbCrLf
        If Not divChgd Then s = s & "  and (isnull(m.SentTo1440,0)=0 OR o.modifieddate>=" & DbQuote(DateTime, ld) & ")" & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 70             tag         --DC options" & vbCrLf
        s = s & "      ,50             parent" & vbCrLf
        s = s & "      ,5             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,c.area        [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(p.communityphase,m.communityphase)          [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,isnull(m.model,'')       [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(m.elevation,''),isnull(m.model,''))       [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,m.series      [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,p.communityphase [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,'add'         [communityoption!70!action]" & vbCrLf
        s = s & "      ,isnull(o.opt,'')         [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,1             [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,cast(o.item1 as varchar)       [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,case isnull(o.IncludedOption,0) when 0 then 'N' else 'Y' end [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,case isnull(o.IncludedOption,0) when 0 then '0' else case o.qty when 0 then '1' else cast(o.qty as varchar) end end   [communityoption!70!standardquantity!element]" & vbCrLf
        's = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        's = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,o.comments    [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tblmodels m" & vbCrLf
        s = s & "     join tbllocality c on (m.area=c.area or isnull(m.area,'')='')" & vbCrLf
        s = s & "     left outer join communityphase p on (p.community=isnull(c.area,m.area) and (m.communityphase=p.communityphase or isnull(m.communityphase,'')=''))" & vbCrLf
        s = s & "     join tbldcoptions o on((o.community=isnull(c.area,m.area) or isnull(o.community,'')='') and (o.communityphase=isnull(p.communityphase,m.communityphase) or isnull(o.communityphase,'')=''))" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        s = s & "  and isnull(m.inactive,0) = 0" & vbCrLf
        If Not divChgd Then s = s & "  and (isnull(m.SentTo1440,0)=0 OR o.modifieddate>=" & DbQuote(DateTime, ld) & ")" & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 80             tag         --lots" & vbCrLf
        s = s & "      ,10             parent" & vbCrLf
        s = s & "      ,6             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,l.community   [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(l.phase,'')       [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,''            [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,'add'         [lot!80!action]" & vbCrLf
        s = s & "      ,isnull(l.lot_no,'')      [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,l.block       [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,l.lotplan     [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,l.lot         [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,l.legaladdress  [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,cast(isnull(l.selling_price,0)+isnull(l.premiumvalue,0) as varchar) [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,cast(isnull(l.actuallength,0)*isnull(l.actualwidth,0) as varchar) [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,case l.status when 'Closed' then 'Settled' when 'Sold' then 'Settled' when 'Inventory' then 'Available' when 'Pending Sale' then 'Sale' when 'Pre Sold Home' then 'Sale' when 'Spec Home' then 'Spec' else 'Sale' end [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,l.ConstructionStatus+1          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,case upper(isnull(l.garage_orientation,'')) when 'R' then 'R' when 'RIGHT' then 'R' when 'L' then 'L' when 'LEFT' then 'L' else '' end [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,l.job_no      [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,l.phase       [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,l.comments    [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllotinventory l" & vbCrLf
        s = s & "     join tbllocality c on l.community=c.area" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        If Not divChgd Then s = s & "  and l.modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 85             tag" & vbCrLf
        s = s & "      ,80             parent" & vbCrLf
        s = s & "      ,6             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,l.community   [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(l.phase,'')       [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,'add'         [lot!80!action]" & vbCrLf
        s = s & "      ,isnull(l.lot_no,'')      [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,l.block       [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,l.lotplan     [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,l.lot         [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,l.legaladdress  [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,cast(isnull(l.selling_price,0)+isnull(l.premiumvalue,0) as varchar) [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,cast(isnull(l.actuallength,0)*isnull(l.actualwidth,0) as varchar) [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,case l.status when 'Closed' then 'Settled' when 'Sold' then 'Settled' when 'Inventory' then 'Available' when 'Pending Sale' then 'Sale' when 'Pre Sold Home' then 'Sale' when 'Spec Home' then 'Spec' else 'Sale' end [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,l.ConstructionStatus+1          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,case upper(isnull(l.garage_orientation,'')) when 'R' then 'R' when 'RIGHT' then 'R' when 'L' then 'L' when 'LEFT' then 'L' else '' end [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,l.job_no      [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,l.phase       [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,l.comments    [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllotinventory l" & vbCrLf
        s = s & "     join tbllocality c on l.community=c.area" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        If Not divChgd Then s = s & "  and l.modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 90             tag         --lot address" & vbCrLf
        s = s & "      ,80             parent" & vbCrLf
        s = s & "      ,6             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,l.community   [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(l.phase,'')       [release!40!releasecode!element]               " & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]          " & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,'add'         [lot!80!action]" & vbCrLf
        s = s & "      ,isnull(l.lot_no,'')      [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,l.block       [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,l.lotplan     [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,l.lot         [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,isnull(nullif(l.legaladdress,''),'not available')  [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,cast(isnull(l.selling_price,0)+isnull(l.premiumvalue,0) as varchar) [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,cast(isnull(l.actuallength,0)*isnull(l.actualwidth,0) as varchar) [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,case l.status when 'Closed' then 'Settled' when 'Sold' then 'Settled' when 'Inventory' then 'Available' when 'Pending Sale' then 'Sale' when 'Pre Sold Home' then 'Sale' when 'Spec Home' then 'Spec' else 'Sale' end [removeme!85!status!element]" & vbCrLf
        s = s & "      ,isnull(nullif(l.municipal_address,''),'not available')   [address!90!address1!element]" & vbCrLf
        s = s & "      ,isnull(nullif(l.city,''),c.city)         [address!90!city!element]" & vbCrLf
        s = s & "      ,isnull(nullif(l.province,''),c.province) [address!90!state!element]" & vbCrLf
        s = s & "      ,isnull(nullif(l.zip,''),c.zip)           [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(l.county,''),c.county)     [address!90!county!element]" & vbCrLf
        s = s & "      ,isnull(nullif(l.country,''),c.country)   [address!90!country!element]" & vbCrLf
        s = s & "      ,l.ConstructionStatus+1          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,case upper(isnull(l.garage_orientation,'')) when 'R' then 'R' when 'RIGHT' then 'R' when 'L' then 'L' when 'LEFT' then 'L' else '' end [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,l.job_no      [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,l.phase       [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,l.comments    [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllotinventory l " & vbCrLf
        s = s & "     join tbllocality c on(l.community=c.area)" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        If Not divChgd Then s = s & "  and l.modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 95             tag" & vbCrLf
        s = s & "      ,80             parent" & vbCrLf
        s = s & "      ,6             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,l.community   [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,isnull(l.phase,'')       [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,'add'         [lot!80!action]" & vbCrLf
        s = s & "      ,isnull(l.lot_no,'')      [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,l.block       [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,l.lotplan     [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,l.lot         [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,l.legaladdress  [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,cast(isnull(l.selling_price,0)+isnull(l.premiumvalue,0) as varchar) [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,cast(isnull(l.actuallength,0)*isnull(l.actualwidth,0) as varchar) [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,case l.status when 'Closed' then 'Settled' when 'Sold' then 'Settled' when 'Inventory' then 'Available' when 'Pending Sale' then 'Sale' when 'Pre Sold Home' then 'Sale' when 'Spec Home' then 'Spec' else 'Sale' end [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,l.ConstructionStatus+1          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,case upper(isnull(l.garage_orientation,'')) when 'R' then 'R' when 'RIGHT' then 'R' when 'L' then 'L' when 'LEFT' then 'L' else '' end [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,l.job_no      [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,isnull(nullif(l.phase,''),'N/A')       [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,l.comments    [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,case l.status when 'Inventory' then '3' when 'Show Home' then '3' when 'Spec Home' then '3' else '1' end [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,null          [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,''            [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllotinventory l" & vbCrLf
        s = s & "     join tbllocality c on l.community=c.area" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        If Not divChgd Then s = s & "  and l.modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 100            tag      -- communityaddontype" & vbCrLf
        s = s & "      ,10             parent" & vbCrLf
        s = s & "      ,7             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,l.area        [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,''            [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,'add'         [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,'STOR'        [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllocality l" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=l.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and l.area=" & DbQuote(Str, comID) & vbCrLf
        If Not divChgd Then s = s & "  and modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "  and isnull(l.inactive,0)=0" & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "select 100            tag" & vbCrLf
        s = s & "      ,10             parent" & vbCrLf
        s = s & "      ,7             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,l.area        [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,''            [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,'add'         [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,'PRKING'      [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,null          [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tbllocality l" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=l.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and l.area=" & DbQuote(Str, comID) & vbCrLf
        If Not divChgd Then s = s & "  and modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "  and isnull(l.inactive,0)=0" & vbCrLf
        s = s & "union all----------------------------------------------------------------------------" & vbCrLf
        s = s & "select 110           tag" & vbCrLf
        s = s & "      ,10            parent" & vbCrLf
        s = s & "      ,7             [community!10!Sort!hide]" & vbCrLf
        s = s & "      ,null          [community!10!action]" & vbCrLf
        s = s & "      ,p.community   [community!10!communitycode!element]" & vbCrLf
        s = s & "      ,null          [community!10!name!element]" & vbCrLf
        s = s & "      ,null          [address!20!address1!element]" & vbCrLf
        s = s & "      ,null          [address!20!address2!element]" & vbCrLf
        s = s & "      ,null          [address!20!city!element]" & vbCrLf
        s = s & "      ,null          [address!20!state!element]" & vbCrLf
        s = s & "      ,null          [address!20!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!20!county!element]" & vbCrLf
        s = s & "      ,null          [address!20!country!element]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!action]" & vbCrLf
        s = s & "      ,null          [releasepricing!30!defaultreleasecode!element]" & vbCrLf
        s = s & "      ,''            [release!40!action]" & vbCrLf
        s = s & "      ,''            [release!40!releasecode!element]" & vbCrLf
        s = s & "      ,null          [release!40!releasename!element]" & vbCrLf
        s = s & "      ,null          [communityplan!50!action]" & vbCrLf
        s = s & "      ,''            [divisionplankey!60!plancode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!elevationcode!element]" & vbCrLf
        s = s & "      ,null          [divisionplankey!60!seriescode!element]" & vbCrLf
        s = s & "      ,null          [removeme!65!releasecode!element] " & vbCrLf
        s = s & "      ,null          [removeme!65!baseprice!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!action]" & vbCrLf
        s = s & "      ,''            [communityoption!70!divisionoptioncode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!phasecode!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!price!element]" & vbCrLf
        s = s & "      ,'N'           [communityoption!70!standardoption!element]" & vbCrLf
        s = s & "      ,0             [communityoption!70!standardquantity!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!maxorderqty!element]" & vbCrLf
        s = s & "      ,null          [communityoption!70!extendeddescription!element]" & vbCrLf
        s = s & "      ,null          [lot!80!action]" & vbCrLf
        s = s & "      ,''            [removeme!85!communitylotcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotblockcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsectioncode!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotnumber!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!legaldescription!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotpremium!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!lotsize!element]" & vbCrLf
        s = s & "      ,null          [removeme!85!status!element]" & vbCrLf
        s = s & "      ,null          [address!90!address1!element]" & vbCrLf
        s = s & "      ,null          [address!90!city!element]" & vbCrLf
        s = s & "      ,null          [address!90!state!element]" & vbCrLf
        s = s & "      ,null          [address!90!postalcode!element]" & vbCrLf
        s = s & "      ,null          [address!90!county!element]" & vbCrLf
        s = s & "      ,null          [address!90!country!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!phasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!lotswingrestriction!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!jobcode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!releasecode!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!notes!element]" & vbCrLf
        s = s & "      ,null          [removeme!95!saleavailability!element]" & vbCrLf
        s = s & "      ,'add'         [communityaddontype!100!action]" & vbCrLf
        s = s & "      ,case p.isparking when 0 then 'STORE' else 'PRKING' end         [communityaddontype!100!divisionaddontypecode!element]" & vbCrLf
        s = s & "      ,'add'         [communityaddoninventory!110!action]" & vbCrLf
        s = s & "      ,parkingno     [communityaddoninventory!110!communityaddoninventorycode!element]" & vbCrLf
        s = s & "      ,p.description   [communityaddoninventory!110!description!element]" & vbCrLf
        s = s & "      ,cast(sellingprice as varchar) [communityaddoninventory!110!price!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!action]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandname!element]" & vbCrLf
        s = s & "      ,null          [brandleveldocument!120!brandmarketingname!element]" & vbCrLf
        s = s & "from tblParking p" & vbCrLf
        s = s & "     join tbllocality c on p.community=c.area" & vbCrLf
        s = s & "     join divisioncommunities div on(div.community=c.area)" & vbCrLf
        s = s & "where div.divisionid=" & DbQuote(Num, divID) & vbCrLf
        If bUploadByCommunity Then s = s & "  and c.area=" & DbQuote(Str, comID) & vbCrLf
        If Not divChgd Then s = s & "  and p.modifieddate>=" & DbQuote(DateTime, ld) & vbCrLf
        s = s & "  and isnull(c.inactive,0)=0" & vbCrLf
        s = s & "order by 5,3,17,20,21,22,26,34,55,1" & vbCrLf
        s = s & "for xml explicit" & vbCrLf
            
            
        Dim iSQL As Integer
        Dim sSQL As String
        iSQL = FreeFile
        sSQL = PathAppend(AppWorkingFolder, "1440", "HFExport.sql")
        Open sSQL For Output As #iSQL
        Print #iSQL, s
        Close iSQL
        
        
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        While Not rs.EOF
            r = r + 1
Call SetCaptions("", "generating row " & r)
            Put #iFile, , CleanXML("" & rs(0))
            rs.MoveNext
        Wend
        Put #iFile, , CleanXML("</division>")
        
    
        'close file
        Put #iFile, , CleanXML("</builder>")
        Put #iFile, , CleanXML("</hasp>")
        Close iFile
    
Call SetCaptions("", "closing file...")

        'STRIP removeme's from the file. we did it while writing the file but some may have slipped thru if they were right on the row boundary.
        Dim x As New Scripting.FileSystemObject
        Dim y As Scripting.TextStream
        Set y = x.OpenTextFile(sfile, ForReading, False)
        s = y.ReadAll
        Call y.Close
        Set y = Nothing
        s = Replace(s, "<removeme>", "")
        s = Replace(s, "</removeme>", "")

        On Error Resume Next
        Kill sfile
        On Error GoTo eh
        Set y = x.CreateTextFile(sfile, True, False)
        Call y.Write(s)
        y.Close
        Set y = Nothing
        Set x = Nothing
        
        
Call SetCaptions("", "Transfering Data...")
        Dim a As String
        Dim b As String
        a = sfile
        b = UCase(CleanFileName(divCode) & IIf(bUploadByCommunity, "_" & CleanFileName(comID), "") & ".xml")
        Status "PutFile(""" & a & """, """ & b & """)"
        Call Ftp1.Store(b, a)
                        
        LastUpdate("1440:" & divID & ":" & comID) = ud
    Next
    
    
    '------------------------------------------------
    ' update tblmodels sentto1440 flag
    '------------------------------------------------
    s = ""
    s = s & "update tblmodels set SentTo1440=1" & vbCrLf
    s = s & "where isnull(inactive,0) = 0" & vbCrLf
    s = s & "  and isnull(SentTo1440,0)=0" & vbCrLf
    Call HFApp.SqlExec(s, dbHomeFront)
    
    
    Unload Me
    
Exit Sub
eh:
        MsgBox Err.Description, vbCritical, App.ProductName
'Select Case True
'
'    Case Err.Description Like "*cannot find specified file*"
'        Unload Me
'
'    Case Err.Description Like "*Password not accepted*"
'        MsgBox Err.Description, vbCritical, App.ProductName
'        Unload Me
'
'    Case Err.Description Like "*access denied*"
'        MsgBox Err.Description, vbCritical, App.ProductName
'        Unload Me
'
'End Select
End Sub

Private Function SqlXML(SQL As String) As String
    Dim s As String
    Dim rs As Recordset
    Set rs = HFApp.SqlExec(SQL, dbHomeFront)
    s = ""
    While Not rs.EOF
        s = s & rs(0)
        rs.MoveNext
    Wend
    SqlXML = CleanXML(s)
End Function

Private Function XmlQuote(Text As String, Optional Quote As Boolean) As String
    Dim s As String
    s = Text
    s = Replace(s, Chr(0), "&#00;")
    s = Replace(s, vbTab, "&#09;")
    s = Replace(s, vbLf, "&#10;")
    s = Replace(s, vbCr, "&#13;")
    s = Replace(s, "!", "&#33;")
    s = Replace(s, vbQuote, "&#34;")
    s = Replace(s, "#", "&#35;")
    s = Replace(s, "$", "&#36;")
    s = Replace(s, "%", "&#37;")
    s = Replace(s, "&", "&#38;")
    s = Replace(s, "'", "&#39;")
    s = Replace(s, "*", "&#42;")
    s = Replace(s, "+", "&#43;")
    s = Replace(s, ":", "&#58;")
    s = Replace(s, ";", "&#59;")
    s = Replace(s, "<", "&#60;")
    s = Replace(s, ">", "&#62;")
    s = Replace(s, "?", "&#63;")
    s = Replace(s, "^", "&#94;")
    s = Replace(s, "`", "&#96;")
    s = Replace(s, "{", "&#123;")
    s = Replace(s, "}", "&#125;")
    s = Replace(s, "|", "&#124;")
    s = Replace(s, "~", "&#126;")
    If Quote Then s = vbQuote & s & vbQuote
    XmlQuote = s
End Function


Private Sub Form_Load()
    Call IniGetForm(Me)
    Me.Height = 1815
    Me.Width = 4965


    
    Animation1.Open App.path & "\filecopy.avi"
    Me.Show
    DoEvents
End Sub

Private Sub ConnectFTP()
    Dim ini As String
    Dim server As String
    Dim path As String
    Dim Port As Integer
    Dim s As String
    Dim i As Long
    
    
    'get server, path and port from settings
    s = Trim(HFApp.Options(WebUploadPath))
    If Left(s, 6) = "ftp://" Then s = Mid(s, 7)
    server = Parse(Parse(s, 1, "/"), 1, ":")
    path = Parse(Parse(s, 1, ":"), 2, server & "/")
    If Right(path, 1) <> "/" Then path = path & "/"
    If path = "/" Then path = ""
    Port = Val(Parse(s, 2, ":"))
    If Port < 1 Then Port = 21
    If Port > 32767 Then Port = 21
    
    
    
    
    ini = PathAppend(App.path, "ftp1.ini")

    Call Ftp1.Logout
    
    'timeout
    i = Val(IniGet(ini, "settings", "TimeOut", 0))
    If i > 0 Then
        Status "TimeOut = " & i
        Ftp1.TimeOut = i
    End If
    
    
    'passive
    i = Val(IniGet(ini, "settings", "Passive", 1))
    Status "Passive  = " & b
    Ftp1.Passive = i <> 0
    
    'proxytype
    i = Val(IniGet(ini, "settings", "ProxyType", Ftp1.ProxyType))
    Status "ProxyType = " & i
    
    'proxyusername
    s = IniGet(ini, "settings", "ProxyUsername", Ftp1.ProxyUsername)
    Status "ProxyUsername = """ & s & """"
    Ftp1.ProxyUsername = s
    
    'ProxyPassword
    s = IniGet(ini, "settings", "ProxyPassword", Ftp1.ProxyPassword)
    Status "ProxyPassword = """ & s & """"
    Ftp1.ProxyPassword = s
    
    'proxyhost
    s = IniGet(ini, "settings", "ProxyHost", Ftp1.ProxyHost)
    Status "ProxyHost = """ & s & """"
    Ftp1.ProxyHost = s
        
    'proxyport
    i = Val(IniGet(ini, "settings", "ProxyPort", Ftp1.ProxyPort))
    Status "ProxyPort = " & i
    Ftp1.ProxyPort = i
    
    'login
    Status "Login " & server & ":" & Port
    Ftp1.abort
    Ftp1.abort
    Call Ftp1.Login(server, HFApp.Options(WebUploadUser), HFApp.Options(WebUploadPswd), , Port)
    
    
    If path <> "" Then
        Status "Directory = """ & path & """"
        Ftp1.Directory = path
    End If
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub lblMore_Click()
    If Me.Height < 6735 Then Me.Height = 6735
End Sub


Private Function ValidateData() As Boolean
    Dim rs As Recordset
    Dim s As String

    '1440 wont allow 2 option categories to have the same description
    'this miraculus qry will add a number to the description of any duplcated categories
    s = ""
    s = s & "update tblcategories" & vbCrLf
    s = s & "set description = a.description + isnull((select ' '+cast(nullif(count(*),0) as varchar) from tblcategories b where a.description=b.description and b.category<a.category),'')" & vbCrLf
    s = s & "from tblcategories a" & vbCrLf
    Call HFApp.SqlExec(s, dbHomeFront)

    'check for data errors
    Set rs = HFApp.SqlExec("purch_getdataerrorssummary", dbHomeFront)
    s = ""
    While Not rs.EOF
        s = s & rs("description") & vbCrLf
        rs.MoveNext
    Wend
    If s = "" Then
        ValidateData = True
    Else
        MsgBox "Unable to post" & vbCrLf & vbCrLf & s, vbInformation, App.ProductName
    End If
    
    
End Function

Private Function CleanFileName(ByVal s As String) As String
    
    s = Replace(s, " ", "")
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, ":", "")
    s = Replace(s, "*", "")
    s = Replace(s, "?", "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "|", "")
    s = Replace(s, """", "")
    
    CleanFileName = s
    
End Function

Private Function CleanXML(ByVal s As String) As String
    
    s = Replace(s, "<removeme>", "")
    s = Replace(s, "</removeme>", "")
    
    'single quotes
    s = Replace(s, Chr(145), Chr(39)) 'left curly single quote
    s = Replace(s, Chr(146), Chr(39)) 'right curly single quote
    s = Replace(s, Chr(147), Chr(34)) 'left curly double quote
    s = Replace(s, Chr(148), Chr(34)) 'right curly double quote
    s = Replace(s, Chr(96), Chr(39))  'left accent
    s = Replace(s, Chr(180), Chr(39)) 'right accent
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    CleanXML = s
    
End Function


Private Function SetCaptions(Caption As String, Description As String)
    DoEvents
    If Caption <> "" Then
        lblCaption = Caption
        lblCaption.Refresh
    End If
    lblDesc = Description
    lblDesc.Refresh

End Function





Private Sub Status(s As String)
End Sub




