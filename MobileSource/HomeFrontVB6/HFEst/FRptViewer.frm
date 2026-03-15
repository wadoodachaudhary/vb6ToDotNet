VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRptViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   7752
   ClientLeft      =   2688
   ClientTop       =   1440
   ClientWidth     =   10752
   Icon            =   "FRptViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7752
   ScaleWidth      =   10752
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRView 
      Height          =   4185
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   9555
      _cx             =   16854
      _cy             =   7382
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
      EnableInteractiveParameterPrompting=   0   'False
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
eh: Call errHandler(SRCFILE & "crView_ZoomLevelChanged")
End Sub

Private Sub Form_Activate()
On Error GoTo eh
    If Not zoomed Then
        zoomed = True
        CRView.Zoom mZoomLevel
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
eh: Call errHandler(SRCFILE & "Form_Activate")
End Sub

Private Sub Form_Load()
On Error GoTo eh
    zoomed = False
    Call IniGetForm(Me)
    mZoomLevel = HFApp.Options(RptViewZoom)
    
    
    Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub

Private Sub Form_Resize()
On Error GoTo eh
    CRView.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Exit Sub
eh: Call errHandler(SRCFILE & "Form_Resize")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo eh
    Call IniPutForm(Me)
    HFApp.Options.value(RptViewZoom) = mZoomLevel
    Exit Sub
eh: Call errHandler(SRCFILE & "Form_Unload")
End Sub



Public Sub ShowReport(File As String, Preview As Boolean, Modal As Boolean, ParamArray parameters())
On Error GoTo eh
    Dim Crystal  As New CRAXDRT.Application
    Dim sections As sections
    Dim section  As section
    Dim rptObjs  As ReportObjects
    Dim rptObj   As Object
    Dim subRpt   As Report
    Dim tbl      As DatabaseTable
    Dim cprops   As ConnectionProperties
    Dim i As Long
    
    Dim Dsn As String
    Dim ddb As String
    Dim uId As String
    Dim pwd As String
    
    Dsn = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "DSN="), 1, ";")
    ddb = HFApp.Databases(dbHomefront).DefaultDatabase
    uId = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "UID="), 1, ";")
    pwd = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "PWD="), 1, ";")
    
    Me.Icon = FMain.Icon
    
    'open the .rpt file
    Set Report = Crystal.OpenReport(File)
    Screen.MousePointer = vbHourglass
    
    'set connections
    For Each tbl In Report.Database.Tables
        tbl.SetLogOnInfo Dsn, ddb, uId, pwd
        tbl.Location = ddb & ".dbo." & tbl.Location
    Next
    For Each section In Report.sections
        For Each rptObj In section.ReportObjects
            If rptObj.Kind = crSubreportObject Then
                Set subRpt = rptObj.OpenSubreport
                For Each tbl In subRpt.Database.Tables
                    tbl.SetLogOnInfo Dsn, ddb, uId, pwd
                    tbl.Location = ddb & ".dbo." & tbl.Location
                Next
            End If
        Next
    Next
    



    'set parameters
    On Error Resume Next
    Dim p As ParameterFieldDefinition
    Dim Name  As String
    Dim value As String
    Dim j As Long
    Dim s As String
    For i = LBound(parameters) To UBound(parameters) Step 2
        Name = parameters(i)
        value = parameters(i + 1)
        If value <> "" Then
        Set p = Nothing
        Set p = Report.ParameterFields.GetItemByName(Name)
        Select Case True
            Case p Is Nothing
                'skip
            Case p.Name = "{?PONumber}" And p.EnableMultipleValues = True
                For j = 1 To Parse(value, , "','")
                    s = Parse(value, j, "','")
                    If Left(s, 1) = "'" Then s = Mid(s, 2)
                    If Right(s, 1) = "'" Then s = Left(s, Len(s) - 1)
                    If s <> "" Then
                        Call p.AddCurrentValue(s)
                    End If
                Next
            Case p.Name = "{?DivisionID}"
                p.AddCurrentValue HFApp.DivisionID
            Case p.ValueType = crBooleanField:   p.AddCurrentValue CBool(value)
            Case p.ValueType = crNumberField:    p.AddCurrentValue Val(value)
            Case Else:                           p.AddCurrentValue value
        End Select
        End If
    Next
    
    
    Set p = Nothing
    Set p = Report.ParameterFields.GetItemByName("DivisionID")
    Select Case True
            Case p Is Nothing
            Case p.Name = "{?DivisionID}"
                p.AddCurrentValue Val(HFApp.DivisionID)
    End Select
    
    On Error GoTo eh
    If Preview Then
        CRView.ReportSource = Report
        Me.Caption = "Print Preview - " & File
        CRView.ViewReport
        Me.Show IIf(Modal, 1, 0)
    Else
        CommonDialog.CancelError = True
        CommonDialog.ShowPrinter
        Call Report.SelectPrinter(Printer.DriverName, Printer.DeviceName, Printer.Port)
        Call Report.PrintOut(False, CommonDialog.Copies)
        Screen.MousePointer = vbDefault
        
        
'        Dim hDC      As Long
'        Dim FromPage As Long
'        Dim Topage   As Long
'        Dim Copies   As Integer
'        Dim Collate  As Boolean
'        If VBPrintDlg(hDC, , , FromPage, Topage, True, Copies, , , , Collate) Then
'            For i = 0 To VB.Printers.Count - 1
'                If VB.Printers(i).hDC = hDC Then
'                    Call Report.SelectPrinter(Printer.DriverName, Printer.DeviceName, Printer.Port)
'                End If
'            Next
'            Screen.MousePointer = vbHourglass
'            Call Report.PrintOut(False, Copies, Collate, FromPage, Topage)
'            Screen.MousePointer = vbDefault
'        End If
    End If
    
    
Exit Sub
eh: Select Case Err.Number
    Case cdlCancel:     'ignore it
        Screen.MousePointer = vbDefault
    Case -2147206461
        MsgBox "The POIndex format report could not be found." & vbCrLf & vbCrLf & File, vbExclamation, App.ProductName
    Case Else
        Call errHandler(SRCFILE & "ShowReport")
        Screen.MousePointer = vbDefault
        Unload Me
    End Select
End Sub


Public Sub ShowSQLReport(File As String, Preview As Boolean, Modal As Boolean, Recordset As ADODB.Recordset)
On Error GoTo eh
Const FILENOTFOUND = -2147206461

    
    Dim Crystal  As New CRAXDRT.Application
    Dim i As Long

If ISDEBUG Then MsgBox "1"
    
    
    Screen.MousePointer = vbHourglass
    Me.Icon = FMain.Icon

If ISDEBUG Then MsgBox "2"
        
    'open the .rpt file
    Set Report = Crystal.OpenReport(File)
    
If ISDEBUG Then MsgBox "3"
    
    'reset datasource
    Call Report.Database.Tables(1).SetDataSource(Recordset)
    
If ISDEBUG Then MsgBox "4"
    If Preview Then
        CRView.ReportSource = Report
If ISDEBUG Then MsgBox "5"
        CRView.ViewReport
If ISDEBUG Then MsgBox "6"
        Screen.MousePointer = vbDefault
If ISDEBUG Then MsgBox "7"
        Me.Show IIf(Modal, 1, 0)
If ISDEBUG Then MsgBox "8"
        
    Else
        Dim hDC      As Long
        Dim FromPage As Long
        Dim Topage   As Long
        Dim Copies   As Integer
        Dim Collate  As Boolean
        If VBPrintDlg(hDC, , , FromPage, Topage, True, Copies, False, , , Collate) Then
            For i = 0 To VB.Printers.Count - 1
                If VB.Printers(i).hDC = hDC Then
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
    Case cdlCancel:     'ignore it
    Case Else:          Call errHandler(SRCFILE & "ShowReport", File)
    End Select
    Screen.MousePointer = vbDefault
End Sub

