VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Options"
Option Explicit

Private mDataSourceID As String

Private Type OptionRec
    Name As String
    Value As Variant
    Dirty As Boolean
End Type
Private mOptions() As OptionRec


Public Enum SchedulingUsageEnum
    DoNotUse
    WarnIfIncomplete
    RestrictIfIncomplete
    HoldUntilCompleted
End Enum



Public Enum OptionsEnum
    
    [_AppOptions] '--------------------------
    SaveIncompleteInvoices          'allows users to save invoices that will be completed in approval portal.
    AutoApprovePOInvoices
    AutoReleaseHeldPOInvoices
    UseTimberlineAsPOSource
    DepartmentIsRequired
    EnableBackChargePOs
    BackChargeMarkup
    BackChargeVendorTypes
    EnableLineBackCharge
    BackChargeStatusItem            'which  misc check list item on the TL commitment is used to mark PO's as processed.
    
    AllowJobDrOverride              'let the user specify the debit account on job costed distributions
    
    JobBasedPONumbers               'po numbers begin with job number. if this is true we can improve performance on commitments picklist searches.
    
    WarnIfAcctDateNotCurrent
    DefaultDescription              'when entering invoice items default the description to invoice header description.
    AllowCrossPayingPOs             'show job field in invoice header and if entered, po list includes job po's not vendor po's
    AutoVarianceRemaining
    CalcDatesFromInvoiceDate
    AcctDateDefaultsToRecDate
    AcctDateDefaultsToInvDate
    InvDescDefaultsToJobAddr
    
    RecalcPOTaxAtCurrentRate        'added to help users deal with changed GST rate.
    AutomaticallyClosePO            'automatically close PO's when amount remaining reaches $0.00
    Caption_Vendor
    Caption_Commitment
    Caption_Job
    Caption_Extra
    Caption_Phase
    Caption_Category
    Caption_DebitAccount
    Caption_TaxGroup
    Caption_Retainage
    Caption_RetainagePct
    
    PayablesDisablePayPOsScreen

    InvoiceDateRequired
    ReceivedDateRequired
    PaymentDateRequired
    AccountingDateRequired
    VerifyBudgets
    
    ImportInvoiceFile
    ImportEstimateFile
    ImportMacroFile
    ImportHistoryFolder
    ImportPrintFile
    
    TimberlineDateFormat
    DefaultTaxGroup
    DefaultTaxGroupDesc
    DefaultTaxGroupRate
    DefaultGLAccount
    DefaultGLAccoundDesc
    DontSetPaymentDate
    
    InvoiceCode1Usage
    InvoiceCode2Usage
    InvoiceCode1Validate
    InvoiceCode2Validate
    InvoiceCode1Label
    InvoiceCode2Label
    
    SchedulingUsage

    ShowQtyAndUnitPrice
    AutoCalcHeaderTotals
    AllowScanning
    
    AutoEmail
    EmailSubject
    EmailMessage
    DocumentRoot
    
    ShowJobPM
    ShowJobTitle1
    ShowJobTitle2
    ShowJobTitle3
    ShowJobTitle4
    Caption_JobPM
    Caption_JobTitle1
    Caption_JobTitle2
    Caption_JobTitle3
    Caption_JobTitle4
    Caption_WrapInsurance
    
    
    [_UserOptions] '--------------------------
    ScannerIndex
    ScannerColorFormat
    ImageResolution
    RptViewZoom
    RptViewTree

    AutoPayApprovedPOs
    
End Enum

Public Property Get Value(Index As OptionsEnum)
Attribute Value.VB_UserMemId = 0
    Value = mOptions(Index).Value
End Property
Public Property Let Value(Index As OptionsEnum, RHS)
    mOptions(Index).Dirty = True 'mOptions(Index).Value <> RHS Or mOptions(Index).Dirty
    mOptions(Index).Value = RHS
End Property








Public Sub DBRead()
    Dim rs As ADODB.Recordset
    Dim Index As Long
    Dim s As String
    
    
    
    s = "SELECT OptionName,OptionValue FROM AppOptions WHERE DivisionID=" & DbQuote(num, HFApp.DivisionID) & " AND (ISNULL(UID,'')='' OR UID=" & DbQuote(Str, HFApp.LoginID, False) & ")"
    
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    While Not rs.EOF
            
        Index = ItemIndex(rs("OptionName"))
        
        mOptions(Index).Name = "" & rs("OptionName")
        mOptions(Index).Value = "" & rs("OptionValue")
        mOptions(Index).Dirty = False
        rs.MoveNext
    Wend
    rs.Close

    Value(ImportEstimateFile) = PathAppend(HFApp.Options.Value(Timberline_Data_Path), "Zybertech\PayablesDesk\Estimates.jce")
    Value(ImportHistoryFolder) = PathAppend(HFApp.Options.Value(Timberline_Data_Path), "Zybertech\PayablesDesk\History")
    Value(ImportInvoiceFile) = PathAppend(HFApp.Options.Value(Timberline_Data_Path), "Zybertech\PayablesDesk\Invoices.txt")
    Value(ImportMacroFile) = PathAppend(HFApp.Options.Value(Timberline_Data_Path), "Zybertech\PayablesDesk\Import.mac")
    Value(ImportPrintFile) = PathAppend(HFApp.Options.Value(Timberline_Data_Path), "Zybertech\PayablesDesk\Invoices_Print.prn")
        
End Sub

Private Function ItemIndex(Name As String) As Long
    Dim i As Long
    For i = LBound(mOptions) To UBound(mOptions)
        If UCase(mOptions(i).Name) = UCase(Name) Then
            ItemIndex = i
            Exit Function
        End If
    Next
End Function

Public Sub DBWrite()
    
    Dim i          As Long
    Dim Insert     As String
    Dim Update     As String
    Dim optValue   As String
        
    For i = 1 To UBound(mOptions)
        
        If mOptions(i).Dirty Then
        
            optValue = "" & mOptions(i).Value
            
            If i < [_UserOptions] Or UCase(mOptions(i).Name) = UCase("AutoPayApprovedPOs") Then
                'app options
                Insert = "INSERT INTO AppOptions(DivisionID,UID,OptionName,OptionValue) VALUES(" & DbQuote(num, HFApp.DivisionID) & ",''" & "," & DbQuote(Str, mOptions(i).Name) & "," & DbQuote(Str, optValue) & ")"
                Update = "UPDATE AppOptions SET OptionValue=" & DbQuote(Str, optValue) & " WHERE DivisionID=" & DbQuote(num, HFApp.DivisionID) & " and UID='' AND OptionName=" & DbQuote(Str, mOptions(i).Name)
            Else
                'user options
                Insert = "INSERT INTO AppOptions(DivisionID,UID,OptionName,OptionValue) VALUES(" & DbQuote(num, HFApp.DivisionID) & "," & DbQuote(Str, HFApp.LoginID, False) & "," & DbQuote(Str, mOptions(i).Name) & "," & DbQuote(Str, optValue) & ")"
                Update = "UPDATE AppOptions SET OptionValue=" & DbQuote(Str, optValue) & " WHERE DivisionID=" & DbQuote(num, HFApp.DivisionID) & " and UID=" & DbQuote(Str, HFApp.LoginID) & " AND OptionName=" & DbQuote(Str, mOptions(i).Name)
            End If
            On Error Resume Next
            
            Call HFApp.SqlExec(Insert, dbHomeFront)
            Call HFApp.SqlExec(Update, dbHomeFront)
            
            mOptions(i).Dirty = False
        End If
    Next
    Call DBRead
End Sub

Private Sub Class_Terminate()
    DBWrite
End Sub

Private Sub Class_Initialize()
    Dim i As Integer
    ReDim mOptions(82) As OptionRec
    
    '------------------------------------------------------------
    'App Options
    '------------------------------------------------------------
    i = i + 1: mOptions(i).Name = "SaveIncompleteInvoices":         mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AutoApprovePOInvoices":          mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AutoReleaseHeldPOInvoices":      mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "UseTimberlineAsPOSource":        mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "DepartmentIsRequired":           mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "EnableBackChargePOs":            mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "BackChargeMarkup":               mOptions(i).Value = 0
    i = i + 1: mOptions(i).Name = "BackChargeVendorTypes":          mOptions(i).Value = "Type 1,Type 2,Type 3,Type 4"
    i = i + 1: mOptions(i).Name = "EnableLineBackCharge":           mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "BackChargeStatusItem":           mOptions(i).Value = 1
    
    i = i + 1: mOptions(i).Name = "AllowJobDrOverride":             mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "JobBasedPONumbers":              mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "WarnIfAcctDateNotCurrent":       mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "DefaultDescription":             mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AllowCrossPayingPOs":            mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AutoVarianceRemaining":          mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "CalcDatesFromInvoiceDate":       mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "AcctdateDefaultsToRecDate":      mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AcctdateDefaultsToInvDate":      mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "InvDescDefaultsToJobAddr":       mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "RecalcPOTaxAtCurrentRate":       mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AutomaticallyClosePO":           mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "Caption_Vendor":                 mOptions(i).Value = "Vendor"
    i = i + 1: mOptions(i).Name = "Caption_Commitment":             mOptions(i).Value = "Commitment"
    i = i + 1: mOptions(i).Name = "Caption_Job":                    mOptions(i).Value = "Job"
    i = i + 1: mOptions(i).Name = "Caption_Extra":                  mOptions(i).Value = "Extra"
    i = i + 1: mOptions(i).Name = "Caption_Phase":                  mOptions(i).Value = "Phase"
    i = i + 1: mOptions(i).Name = "Caption_Category":               mOptions(i).Value = "Category"
    i = i + 1: mOptions(i).Name = "Caption_DebitAccount":           mOptions(i).Value = "Debit Account"
    i = i + 1: mOptions(i).Name = "Caption_TaxGroup":               mOptions(i).Value = "Tax Group"
    i = i + 1: mOptions(i).Name = "Caption_Retainage":              mOptions(i).Value = "Retainage"
    i = i + 1: mOptions(i).Name = "Caption_RetainagePct":           mOptions(i).Value = "%"
    
    i = i + 1: mOptions(i).Name = "PayablesDisablePayPOsScreen":    mOptions(i).Value = False
    
    
    i = i + 1: mOptions(i).Name = "InvoiceDateRequired":            mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "ReceivedDateRequired":           mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "PaymentDateRequired":            mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AccountingDateRequired":         mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "VerifyBudgets":                  mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "ImportInvoiceFile":              mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "ImportEstimateFile":             mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "ImportMacroFile":                mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "ImportHistoryFolder":            mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "ImportPrintFile":                mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "TimberlineDateFormat":           mOptions(i).Value = "mm-dd-yy"
    i = i + 1: mOptions(i).Name = "DefaultTaxGroup":                mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "DefaultTaxGroupDesc":            mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "DefaultTaxGroupRate":
    i = i + 1: mOptions(i).Name = "DefaultGLAccount":               mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "DefaultGLAccountDesc":           mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "DontSetPaymentDate":             mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "InvoiceCode1Usage":              mOptions(i).Value = "Hidden"
    i = i + 1: mOptions(i).Name = "InvoiceCode2Usage":              mOptions(i).Value = "Hidden"
    i = i + 1: mOptions(i).Name = "InvoiceCode1Validate":           mOptions(i).Value = "None"
    i = i + 1: mOptions(i).Name = "InvoiceCode2Validate":           mOptions(i).Value = "None"
    i = i + 1: mOptions(i).Name = "InvoiceCode1Label":              mOptions(i).Value = "Code 1"
    i = i + 1: mOptions(i).Name = "InvoiceCode2Label":              mOptions(i).Value = "Code 2"
    i = i + 1: mOptions(i).Name = "SchedulingUsage":                 mOptions(i).Value = SchedulingUsageEnum.DoNotUse
    i = i + 1: mOptions(i).Name = "ShowQtyAndUnitPrice":            mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AutoCalcHeaderTotals":           mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "AllowScanning":                  mOptions(i).Value = True
    
    i = i + 1: mOptions(i).Name = "AutoEmail":                      mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "EmailSubject":                   mOptions(i).Value = "Request for approval"
    i = i + 1: mOptions(i).Name = "EmailMessage":                   mOptions(i).Value = "An invoice has been placed on hold pending your approval. Please review the attached document(s) and advise me on course of action." & vbCrLf & vbCrLf & _
                                                                                        "Job: <%Job%>" & vbCrLf & _
                                                                                        "Vendor: <%Vendor%>" & vbCrLf & _
                                                                                        "Invoice: <%Invoice%>" & vbCrLf & _
                                                                                        "Reason: <%Reason%>" & vbCrLf & vbCrLf & _
                                                                                        "Thank you"
    
    
    i = i + 1: mOptions(i).Name = "DocumentRoot":                   mOptions(i).Value = ""
    
    i = i + 1: mOptions(i).Name = "ShowJobPM":                      mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ShowJobTitle1":                  mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ShowJobTitle2":                  mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ShowJobTitle3":                  mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ShowJobTitle4":                  mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "Caption_JobPM":                  mOptions(i).Value = "Project Manager"
    i = i + 1: mOptions(i).Name = "Caption_JobTitle1":              mOptions(i).Value = "Misc Title 1"
    i = i + 1: mOptions(i).Name = "Caption_JobTitle2":              mOptions(i).Value = "Misc Title 2"
    i = i + 1: mOptions(i).Name = "Caption_JobTitle3":              mOptions(i).Value = "Misc Title 3"
    i = i + 1: mOptions(i).Name = "Caption_JobTitle4":              mOptions(i).Value = "Misc Title 4"
    i = i + 1: mOptions(i).Name = "Caption_WrapInsurance":          mOptions(i).Value = "Wrap Insurance"

    
    '------------------------------------------------------------
    i = i + 1 'User Options
    '------------------------------------------------------------
    i = i + 1: mOptions(i).Name = "ScannerIndex":                   mOptions(i).Value = 0
    i = i + 1: mOptions(i).Name = "ScannerColorFormat":             mOptions(i).Value = 2   'full color
    i = i + 1: mOptions(i).Name = "ImageResolution":                mOptions(i).Value = 600
    i = i + 1: mOptions(i).Name = "RptViewZoom":                    mOptions(i).Value = 2
    i = i + 1: mOptions(i).Name = "RptViewTree":                    mOptions(i).Value = False
    'Another App Option
    i = i + 1: mOptions(i).Name = "AutopayApprovedPOs":             mOptions(i).Value = False


End Sub



