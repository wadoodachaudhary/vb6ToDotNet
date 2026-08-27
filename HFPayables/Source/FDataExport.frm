VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FDataExport 
   Caption         =   "Pricelist Import Wizard"
   ClientHeight    =   6165
   ClientLeft      =   3585
   ClientTop       =   1785
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FDataExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7305
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   2655
      Left            =   390
      TabIndex        =   0
      Top             =   420
      Width           =   7155
      _cx             =   12621
      _cy             =   4683
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
      FormatString    =   $"FDataExport.frx":000C
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
Attribute VB_Name = "FDataExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FDataExport::"

Public Function ExportData(SelectStatement As String, Optional DialogTitle As String)
On Error GoTo eh
    Dim hwnd As Long
    Dim FileName As String
    Dim i As Long
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ra As Long
    
    i = IniGet(AppIni, "Options", "DataExportFileExt", 0)
    
    On Error Resume Next
    hwnd = Screen.ActiveForm.hwnd
    On Error GoTo eh
    Me.Caption = DialogTitle
    If VBGetSaveFileName(FileName, , , "Excel Files (*.xls)|*.xls|Text (Tab delimited)(*.txt)|*.txt|CSV (Comma delimited)(*.csv)|*.csv", i, , DialogTitle, "xls", hwnd) Then
        Screen.MousePointer = vbHourglass
        Call IniPut(AppIni, "Options", "DataExportFileExt", i)
        Load Me
        Set cn = New ADODB.Connection
        cn.ConnectionString = HFApp.ConnectionString(dbHomeFront)
        cn.Open
        
        Set rs = cn.Execute(SelectStatement, ra) ' HFApp.SqlExec(SelectStatement, dbHomeFront, ra)
        Set gData.DataSource = rs
        gData.FixedRows = 0
        
        
        'format these stupid things so excel doesnt break them
        For i = 0 To rs.Fields.Count - 1
            Select Case rs.Fields(i).Type
                'date/times
                Case adDate, adDBDate, adFileTime, adDBTime, adDBTimeStamp
                    gData.ColDataType(i + 1) = DataTypeSettings.flexDTDate
                'numeric
                Case adVarNumeric, adVarWChar, adCurrency, adDecimal, adBinary, adBoolean, adDouble, adWChar, adInteger, adBigInt, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    gData.ColDataType(i + 1) = DataTypeSettings.flexDTDouble
                'text
                Case Else
                    gData.ColDataType(i + 1) = DataTypeSettings.flexDTString
            End Select
        Next


        Select Case FileExt(FileName)
            Case "csv":   Call gData.SaveGrid(FileName, flexFileCommaText, False)
            Case "xls":   Call gData.SaveGrid(FileName, flexFileExcel)
            Case Else:    Call gData.SaveGrid(FileName, flexFileTabText, False)
        End Select
        rs.Close
        cn.Close
        Unload Me
        MsgBox "Export Complete", vbOKOnly + vbInformation, DialogTitle
        Screen.MousePointer = vbDefault
    End If
        
Exit Function
eh: Call ErrHandler(SRCFILE & "ExportData")
End Function


