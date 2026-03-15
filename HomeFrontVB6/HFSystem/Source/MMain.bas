Attribute VB_Name = "MMain"
Option Explicit

Public SimplyOdbcUtilities As SimplyWrapper.ODBCUtilities

Public mProductCode As String
Public mProductName As String
Public mModuleCode As String
Public mModuleName As String
Public mModuleVersion As String
Public mSystemVersion As String

Public mExePath As String

Public mTrustedUser As String
Public mTrustedConnection As String
Public mTrustedSkipAccounting As Boolean

Public IsProductionDatabase As Boolean
Public IsProductionVerified As Boolean
Public HFApp As HFSystem.Application

Public Closing As Boolean

Public tlFolder As String
Public StatusCtrl As Label

Public Const FILTERBARBACKCOLOR = &HFFEFE3   'vbInfoBackground
Public Const FILTERBARFORECOLOR = vbWindowText
Public Const FILTERBARFONTBOLD = True

Public Sub WriteLogFile(Filename As String, Text As String, Optional Append As Boolean = False)
    Dim i As Integer
    Dim p As String
    
    If Not InIde() Then
        Text = Replace(Text, "<password>1Hyphensolutions!</password>", "<password>####</password>")
    End If
    
    i = FreeFile()
    p = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront")
    Call CreatePath("", p)
    If Append Then
        Open PathAppend(p, Filename) For Append As #i
    Else
        Open PathAppend(p, Filename) For Output As #i
    End If
    Print #i, Text
    Close #i

End Sub

Public Property Get StatusLabel() As String
    StatusLabel = StatusCtrl.Caption
End Property


Public Property Let StatusLabel(RHS As String)
On Error Resume Next
    StatusCtrl.Caption = RHS
    StatusCtrl.Refresh
End Property

Public Function JobSimplicity() As Boolean
    Dim i As Long
    With HFApp.License
        i = .ModuleIndex("JobSimplicity")
        JobSimplicity = .ModuleDemo(i) Or .ModuleSeats(i) > 0
    End With
End Function


Public Function GetCustomDesc(ItemName As String) As String
On Error Resume Next
    GetCustomDesc = ItemName
    GetCustomDesc = HFApp.SqlExec("select isnull(nullif(custom_description,''),item) from customdescriptions where item=" & DbQuote(Str, ItemName), dbHomefront)(0)
End Function


Public Sub ApplyFilter(g As VSFlexGrid)
' apply filter to hide stuff that doesn't match
    Dim r As Long
    Dim c As Long

    Screen.MousePointer = vbHourglass
    With g
        .Redraw = flexRDNone
        For r = 2 To .Rows - 1
            If .RowData(r) = "delete" Then
                .RowHidden(r) = True
            Else
                .RowHidden(r) = False
                For c = 0 To .Cols - 1
                    If .ColComboList(c) <> "" And .TextMatrix(1, c) = "0" Then
                        .TextMatrix(1, c) = ""
                    End If
                
                
                    If .ColDataType(c) = flexDTBoolean Then
                        If .Cell(flexcpChecked, 1, c) = flexChecked And .Cell(flexcpChecked, r, c) = flexUnchecked Then
                            .RowHidden(r) = True
                            Exit For
                        End If
                    Else
                        If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                            .RowHidden(r) = True
                            Exit For
                        End If
                    End If
                Next
            End If
        Next
        .Redraw = flexRDBuffered
    
    End With
    Screen.MousePointer = vbDefault

End Sub

