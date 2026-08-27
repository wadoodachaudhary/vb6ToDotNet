VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CostCodes"
Attribute VB_Ext_KEY = "Member1" ,"Lists"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member2" ,"User"
Attribute VB_Ext_KEY = "Member3" ,"Options"
Option Explicit
Private Const SRCFILE = "CApp::"

Public Enum ApplicationConnections
    dbTLDDF
End Enum
Private mDatabases(0) As adodb.Connection

Public Admin              As Boolean
Public PostInvoices       As Boolean
Public EditInvoices       As Boolean
Public EditPostedInvoices As Boolean
Public ApproveInvoices    As Boolean
Public EditVendors        As Boolean
Public InvoiceMax         As Double
Public InvOverridePcnt    As Double
Public InvOverrideMax     As Double
Public POOveragePcnt      As Double
Public POOverageMax       As Double

Public Options         As New Options




Public Sub ConvertSQL(Sql As String, c As adodb.Connection)
On Error GoTo eh

    Dim i As Long
    Dim pieces() As String
    Dim s As String
    
    Select Case c.Properties("DBMS Name")
        
        Case "Pervasive.SQL"
          'do nothing
            
        Case "Microsoft SQL Server"
            pieces = split(Sql, "'")
            For i = LBound(pieces) To UBound(pieces) Step 2
                s = pieces(i)
                s = Replace(s, " DATE ", " DATETIME", , , vbTextCompare)
                s = Replace(s, "TIMESTAMPDIFF", "DATEDIFF", , , vbTextCompare)
                s = Replace(s, "SQL_TSI_MONTH", "M", , , vbTextCompare)
                s = Replace(s, "CURRENT_TIMESTAMP()", "GETDATE()", , , vbTextCompare)
                s = Replace(s, "IFNULL", "ISNULL", , , vbTextCompare)
                s = Replace(s, "CURDATE", "GETDATE", , , vbTextCompare)
                s = Replace(s, "NOW()", "GETDATE()", , , vbTextCompare)
                s = Replace(s, "LONGVARCHAR", "VARCHAR(2500)", , , vbTextCompare)
                s = Replace(s, "DOUBLE", "FLOAT", , , vbTextCompare)
                s = Replace(s, " IDENTITY", " INTEGER IDENTITY", , , vbTextCompare)
                s = Replace(s, " CASE", "", , , vbTextCompare)
                s = Replace(s, " DATETIMETIME", " DATETIME", , , vbTextCompare)
                s = Replace(s, " TIMESTAMP", " DATETIME", , , vbTextCompare)
                s = Replace(s, " TIME", " DATETIME", , , vbTextCompare)
                s = Replace(s, "UCASE(", "UPPER(", , , vbTextCompare)
                s = Replace(s, "LCASE(", "LOWER(", , , vbTextCompare)
                s = Replace(s, "CURTIME(", "GETDATE(", , , vbTextCompare)
                s = Replace(s, "CONVERT(", "CAST(", , , vbTextCompare)
                s = Replace(s, ", SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
                s = Replace(s, ",SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
                s = Replace(s, ", SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
                s = Replace(s, ",SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
                s = Replace(s, " MODIFY ", " ALTER COLUMN ", , , vbTextCompare)
                s = Replace(s, "CAST(VARCHAR", "CONVERT(VARCHAR", , , vbTextCompare)
                
                pieces(i) = s
            Next
            s = ""
            For i = LBound(pieces) To UBound(pieces)
                s = s & "'" & pieces(i)
            Next
            Sql = Mid(s, 2)
            
    End Select
    Exit Sub
eh:
    Select Case err.Number
    
        Case 91   'Object variable or With block variable not set
            err.Raise vbObjectError, "App::ConvertSQL()", "Database connection is nothing"
            
        Case 3265 'Item cannot be found in the collection corresponding to the requested name or ordinal.
            Exit Sub
            
        Case Else
            err.Raise err.Number
            
    End Select
End Sub

Public Function Login() As Boolean
On Error GoTo eh
    Dim rs As Recordset
    
    
    
    Set rs = HFApp.SqlExec("SELECT * FROM user_manager WHERE UCASE(User_ID)=UCASE(" & DbQuote(Str, HFApp.LoginID) & ")", dbHomeFront)
    If rs.EOF Then
        MsgBox "Login ID not found", vbCritical
        Login = False
        Exit Function
    Else
    
        Admin = "" & rs("CanEditPayablesSettings") = "True"
        EditInvoices = "" & rs("CanEditPayableInvoices") = "True"
        EditPostedInvoices = "" & rs("CanEditPostedPayableInvoices") = "True"
        ApproveInvoices = "" & rs("CanApprovePayableInvoices") = "True"
        PostInvoices = "" & rs("CanPostPayableInvoices") = "True"
        EditVendors = "" & rs("CanEditVendorSetup") = "True"
        
        InvoiceMax = Val("" & rs("MaxInvoiceAmount"))
        
        InvOverridePcnt = Val("" & rs("MaxInvoiceOverridePercent"))
        InvOverrideMax = Val("" & rs("MaxInvoiceOverrideAmount"))
        
        POOveragePcnt = Val("" & rs("MaxInvoicePOOveragePercent"))
        POOverageMax = Val("" & rs("MaxInvoicePOOverageAmount"))
        
        Call Options.DBRead
        
        
        On Error GoTo eh
        Set rs = HFApp.SqlExec("select description,grouprate from taxgroups where DivisionID =" & HFApp.divisionid & " and taxgroup=" & DbQuote(Str, App.Options(DefaultTaxGroup)), dbHomeFront)
        If Not rs.EOF Then
            App.Options.Value(DefaultTaxGroupDesc) = "" & rs(0)
            App.Options.Value(DefaultTaxGroupRate) = Val("" & rs(1))
        End If
        
        
        Login = True
        Call UpdatePOCompletions
    End If
    
    Exit Function
eh: Call ErrHandler("MMain::AppLogin")
End Function



















Public Property Get Title() As String
    Title = VB.App.Title
End Property

Public Property Get ProductName() As String
    ProductName = VB.App.ProductName
End Property
Public Property Get Major() As Integer
    Major = VB.App.Major
End Property
Public Property Get Minor() As Integer
    Minor = VB.App.Minor
End Property
Public Property Get Revision() As Integer
    Revision = VB.App.Revision
End Property
Public Property Get CompanyName() As String
    CompanyName = VB.App.CompanyName
End Property
Public Property Get LegalCopyright() As String
    LegalCopyright = VB.App.LegalCopyright
End Property
Public Property Get LegalTrademarks() As String
    LegalTrademarks = VB.App.LegalTrademarks
End Property
Public Property Get path() As String
    path = VB.App.path
End Property
Public Property Get ExeName() As String
    ExeName = VB.App.ExeName
End Property

