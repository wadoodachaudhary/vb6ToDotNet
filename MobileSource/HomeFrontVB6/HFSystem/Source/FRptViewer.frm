VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRptViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   7755
   ClientLeft      =   2580
   ClientTop       =   1425
   ClientWidth     =   10755
   Icon            =   "FRptViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10755
   Begin CRVIEWER9LibCtl.CRViewer9 CRView 
      Height          =   5745
      Left            =   1260
      TabIndex        =   1
      Top             =   1020
      Width           =   9045
      lastProp        =   600
      _cx             =   15954
      _cy             =   10134
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2460
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   735
      Left            =   -900
      TabIndex        =   0
      Top             =   900
      Width           =   735
   End
End
Attribute VB_Name = "FRptViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FRptViewer::"
Private mZoomLevel As Long
Private mIsPrintForm As Boolean
Private zoomed As Boolean

Public Report   As CRAXDRT.Report


Private Sub Command1_Click()
On Error Resume Next
    Unload Me
End Sub



Private Sub crView_PrintButtonClicked(UseDefault As Boolean)
'On Error GoTo eh
'
'    UseDefault = False
'    With CommonDialog
'        .Min = 1
'        .Max = 9999
'        .FromPage = 1
'        .Topage = 9999
'        .flags = cdlPDNoSelection
'        .CancelError = True
'        .ShowPrinter
'        Call Report.PrintOut(False, .Copies, (.flags And cdlPDCollate) = cdlPDCollate, .FromPage, .Topage)
'    End With
'
'    Exit Sub
'eh:
'    If Err.Number = cdlCancel Then
'        'do nothing
'    Else
'        Call ErrHandler(SRCFILE & "crView_PrintButtonClicked")
'    End If
End Sub

Private Sub CRView_ZoomLevelChanged(ByVal ZoomLevel As Integer)
On Error GoTo eh
    If zoomed Then
        mZoomLevel = ZoomLevel
    End If
    Exit Sub
eh: Call ErrHandler(SRCFILE & "crView_ZoomLevelChanged")
End Sub

Private Sub Form_Activate()
On Error GoTo eh
    If Not zoomed Then
        zoomed = True
        CRView.Zoom mZoomLevel
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
eh: Call ErrHandler(SRCFILE & "Form_Activate")
End Sub

Private Sub Form_Load()
On Error GoTo eh
    zoomed = False
    Call IniGetForm(Me)
    mZoomLevel = HFApp.Options(RptViewZoom)
    
'    If mIsPrintForm Then
'        CRView.EnableGroupTree = False
'        CRView.EnableSearchExpertButton = False
'        CRView.EnableSelectExpertButton = False
'        CRView.EnableRefreshButton = False
'        CRView.DisplayGroupTree = False
'    Else
'        CRView.EnableGroupTree = True
'        CRView.EnableSearchExpertButton = True
'        CRView.EnableSelectExpertButton = True
'        CRView.EnableRefreshButton = True
'        CRView.DisplayGroupTree = HFApp.Options(RptViewTree)
'    End If
    
    Exit Sub
eh: Call ErrHandler(SRCFILE & "Form_Load")
End Sub

Private Sub Form_Resize()
On Error GoTo eh
    CRView.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Exit Sub
eh: Call ErrHandler(SRCFILE & "Form_Resize")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo eh
    Call IniPutForm(Me)
    HFApp.Options.Value(RptViewZoom) = mZoomLevel
    
    If Not mIsPrintForm Then
        HFApp.Options.Value(RptViewTree) = CRView.DisplayGroupTree
    End If
    
    Exit Sub
eh: Call ErrHandler(SRCFILE & "Form_Unload")
End Sub



Public Sub ShowReport(File As String, Preview As Boolean, IsPrintForm As Boolean, Modal As Boolean, ParamArray parameters())
On Error GoTo eh
Const FILENOTFOUND = -2147206461
    Dim Crystal  As New CRAXDRT.Application
    Dim sections As sections
    Dim Section  As Section
    Dim rptObjs  As ReportObjects
    Dim rptObj   As Object
    Dim subRpt   As Report
    Dim tbl      As DatabaseTable
    Dim cprops   As ConnectionProperties
    Dim i As Long
    
    Dim Dsn As String
    Dim ddb As String
    Dim uid As String
    Dim pwd As String
    
    Dsn = Parse(Parse(HFApp.Databases(dbHomefront).ConnectionString, 2, "DSN="), 1, ";")
    ddb = HFApp.Databases(dbHomefront).DefaultDatabase
    uid = Parse(Parse(HFApp.Databases(dbHomefront).ConnectionString, 2, "UID="), 1, ";")
    pwd = Parse(Parse(HFApp.Databases(dbHomefront).ConnectionString, 2, "PWD="), 1, ";")
    
    
    'open the .rpt file
    Set Report = Crystal.OpenReport(File)
    Screen.MousePointer = vbHourglass
    
    'set connections
    For Each tbl In Report.Database.Tables
        tbl.SetLogOnInfo Dsn, ddb, uid, pwd
        tbl.Location = ddb & ".dbo." & tbl.Location
        
        Set cprops = tbl.ConnectionProperties
        cprops.DeleteAll
        cprops.Add "Provider", "SQLOLEDB"
        cprops.Add "Data Source", Dsn
        cprops.Add "Initial Catalog", ddb
        cprops.Add "User ID", uid
        cprops.Add "Password", pwd
        cprops.Add "Integrated Security", uid = ""
    Next
    
    For Each Section In Report.sections
        For Each rptObj In Section.ReportObjects
            If rptObj.Kind = crSubreportObject Then
                Set subRpt = rptObj.OpenSubreport
                For Each tbl In subRpt.Database.Tables
                    tbl.SetLogOnInfo Dsn, ddb, uid, pwd
                    tbl.Location = ddb & ".dbo." & tbl.Location
                    
                    Set cprops = tbl.ConnectionProperties
                    cprops.DeleteAll
                    cprops.Add "Provider", "SQLOLEDB"
                    cprops.Add "Data Source", Dsn
                    cprops.Add "Initial Catalog", ddb
                    cprops.Add "User ID", uid
                    cprops.Add "Password", pwd
                    cprops.Add "Integrated Security", uid = ""
                Next
            End If
        Next
    Next
    


    'set parameters
    On Error Resume Next
    For i = LBound(parameters) To UBound(parameters) Step 2
        Select Case Report.ParameterFields.GetItemByName(parameters(i)).ValueType
            Case crBooleanField: Report.ParameterFields.GetItemByName(parameters(i)).AddCurrentValue CBool(parameters(i + 1))
            Case Else:           Report.ParameterFields.GetItemByName(parameters(i)).AddCurrentValue (parameters(i + 1))
        End Select
    Next
    On Error GoTo eh
    
    If Preview Then
        mIsPrintForm = IsPrintForm
        CRView.ReportSource = Report
'        Me.Caption = "Print Preview - " & FileName(File)
        Me.Caption = "Print Preview - " & File
        CRView.ViewReport
        Me.Show IIf(Modal, 1, 0)
    Else
        Dim hdc      As Long
        Dim FromPage As Long
        Dim Topage   As Long
        Dim Copies   As Integer
        Dim Collate  As Boolean
        If VBPrintDlg(hdc, , , FromPage, Topage, True, Copies, , , , Collate) Then
            For i = 0 To VB.Printers.Count - 1
                If VB.Printers(i).hdc = hdc Then
                    Call Report.SelectPrinter(Printer.DriverName, Printer.DeviceName, Printer.Port)
                End If
            Next
            Screen.MousePointer = vbHourglass
            Call Report.PrintOut(False, Copies, Collate, FromPage, Topage)
            Screen.MousePointer = vbDefault
        End If
    End If
    
    
Exit Sub
eh: Select Case Err.Number
    Case FILENOTFOUND:  Err.Raise 438
    Case cdlCancel:     'ignore it
    Case Else
        Call ErrHandler(SRCFILE & "ShowReport")
        Screen.MousePointer = vbDefault
        Unload Me
    End Select
End Sub


Public Sub ShowSQLReport(File As String, Preview As Boolean, IsPrintForm As Boolean, Modal As Boolean, Recordset As ADODB.Recordset)
On Error GoTo eh
Const FILENOTFOUND = -2147206461

    Dim Crystal  As New CRAXDRT.Application
    Dim i As Long
    
    Screen.MousePointer = vbHourglass
        
    'open the .rpt file
    Set Report = Crystal.OpenReport(File)
    
    'reset datasource
    Call Report.Database.Tables(1).SetDataSource(Recordset)
    
    If Preview Then
        mIsPrintForm = IsPrintForm
        CRView.ReportSource = Report
        Me.Caption = "Print Preview" 'FileName(File)
        CRView.ViewReport
        Screen.MousePointer = vbDefault
        Me.Show IIf(Modal, 1, 0)
    Else
        Dim hdc      As Long
        Dim FromPage As Long
        Dim Topage   As Long
        Dim Copies   As Integer
        Dim Collate  As Boolean
        If VBPrintDlg(hdc, , , FromPage, Topage, True, Copies, False, , , Collate) Then
            For i = 0 To VB.Printers.Count - 1
                If VB.Printers(i).hdc = hdc Then
                    Call Report.SelectPrinter(Printer.DriverName, Printer.DeviceName, Printer.Port)
                End If
            Next
            Screen.MousePointer = vbHourglass
            Call Report.PrintOut(False, Copies, Collate, FromPage, Topage)
            Screen.MousePointer = vbDefault
        End If
    End If
    
    
Exit Sub
eh: Select Case Err.Number
    Case FILENOTFOUND:  Err.Raise 438
    Case cdlCancel:     'ignore it
    Case Else:          Call ErrHandler(SRCFILE & "ShowReport", File)
    End Select
    Screen.MousePointer = vbDefault
End Sub

