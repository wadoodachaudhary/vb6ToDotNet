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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Options"
Option Explicit
Option Compare Text
Private Type OptionRec
    Name As String
    Value As Variant
    Dirty As Boolean
End Type
Private mOptions() As OptionRec


'things to be changed on next break
'add option for WebAttachmentsFolder
'add option for UseChangeRequests
'add option for ChangeRequestStatuses



Public Enum OptionsEnum
    
    [_AppOptions] '--------------------------
        
        SendPO_CCPrjMgr
        SendPO_ViaMapi
        SendPO_SMTPServer
        SendPO_FromCurrentUsersEmail
        SQLDate_Format
        
        WebUploadByCommunity
        PurchasingRequiresSalesApproval
        Msg_RFP_Subject
        Msg_RFP_Body
        LetPurchaserChangeSaleQty
        FaxServer
        FaxCoverPage
        AccountingVersion
        AppendUOMtoOptionDesc
        
        UseCommunityStandards
        CommunityStandardsArePhaseSpecific
        
        RecordSalesSheetCosts
        VendorPriceChangeCategory
        UserControlOverBudgets
        IncludeLandCostsInSalesSheetMarkup
        UseJobAsMasterBuilderSubAccount
        
        
        POPrefix ' see foptions and generate po's in festimateitems
        ZeroQtyTakeoffs ' see ftakeoff
        SalesSystem
        
        AccountingSystem
        
        SimplyDataFile
        
        QuickBooksDataFile
        
        MasterBuilderDataFolder
        MasterBuilderCompany
        MasterBuilderUID
        MasterBuilderPWD
        
        FaxProvider
        FaxBillingCode
        Format_QtyRateEQZero_ForeColor
        Format_QtyRateEQZero_BackColor
        Format_QtyRateEQZero_FontStyle
        Format_QtyRateLTZero_ForeColor
        Format_QtyRateLTZero_BackColor
        Format_QtyRateLTZero_FontStyle
        Format_InvalidData_ForeColor
        Format_InvalidData_BackColor
        Format_InvalidData_FontStyle
        Format_NonSysRate_ForeColor
        Format_NonSysRate_BackColor
        Format_NonSysRate_FontStyle
        UsePricingFromEstimating
        ZeroRateOnRefreshCosts
        ZeroRateOnChangeVendor
        GroupPurchaseOrdersByVendor 'combine PO Indexes in the same group for the same vendor and the same format to a single purchase order.
        
        TimberlineEstimatingPath
        TimberlineEstimatingDictionaryPath
        EstimatingUser
        EstimatingPswd
        EstimatingVersion
        EstimatingMajorVersion
        EstimatingMinorVersion
        
        ImportDateARCustomers
        ImportDateStandardCostCodes
        ImportDateEstItems
        ImportDateTaxGroups
        ImportDateVendors
        ImportDateAccounts
        
        TaxGroupRequired
        TaxGroupLabour
        TaxGroupMaterial
        TaxGroupSubContract
        TaxGroupEquipment
        TaxGroupOverhead
        TaxGroupOther
        IncentivePhase
        IncentiveItem
        IncentiveItemDescription
        SelectAtTakeoffItem
        SelectAtTakeoffPhase
        SelectAtTakeoffDescription
        IncentiveCostPercent  'incentive is entered at retail. this portion of it is costed...
        LandPhase
        LandItem
        LandItemDescription
        GlobalPONumberSequence
        POFormat
        POFormatSegment1
        POFormatSegment2
        POFormatSegment3
        AdjustSellPriceWithCosts  'in estimator sales worksheet recalc selling price when costs change
        LockPostedSalesWorksheets
        AllowForecastsWhenRefreshingPOs 'when they refresh costs on a PO or EstItem should we let them choose a forecasted cost basis's
        WarnForecastsWhenRefreshingPOs
        PostSummarizedBudgets           'when posting budgets summarize to costcode/category rather than each individual item.
        PostSummarizedPOs               'when posting POs summarize to costcode/category rather than each individual item.
        IssueBudgetsByPhase             'show the estimate and PO screens by community phase
        UseEstimatingWorksheetsOnly     'hide marketing worksheet task. if false then it only publishes costs
        MarketingWorkSheetAutoRefreshCosts
        BarCodeDCRoomRequired
        CanAddContractItemsFromPurchasing
        
        Msg_CancelPO_Subject
        Msg_CancelPO_Body
        Msg_SendPO_Subject
        Msg_SendPO_Body
        Msg_SendPO_CC_PrjMgr
        
    [_SystemSettings] '-- read from system_setup -------------------------
        COEntryPriorToPurchase
        ShowAssignCustOnly
        RemoveDetailsonSpecConvert
        PromptForSpecOptions
        SendFaxUsingPurchasersInfo
        accounting_approve
        Acct_Receivable
        ACPrefix
        address1
        address2
        Adjustment_Credit
        allow_customprice
        AlternatingGridColor
        ApplyLotCostJob
        ApproveCOReqQuote
        AreaSpecSalesPerson
        Backup_Folder
        BalanceSheetMaxAcct
        Cardel
        city
        BuilderCode
        BuilderName
        WebAdminUser
        WebUploadPath
        WebUploadUser
        WebUploadPswd
        Country
        ContactPerson
        
        Comm_Category
        Comm_Cost_code
        Comm_Type
        Community_Cost_Code
        CompanyLogo
        CompanyName
        Contract_Document
        ContractDocInWord
        COPrefix
        CORevenue
        Cost_Code_Mask
        Cost_Code_section1
        Cost_Code_section2
        Cost_Code_section3
        Cost_Code_section4
        Cost_Code_Separator
        CostCodeSaleClosing
        CreateJobLot
        Cust_Section1
        Cust_Section2
        Cust_Section3
        Cust_Separator
        Customer_Mask
        customer_section1
        customer_section2
        customer_section3
        customer_Separator
        CustomOptionsTaxRate
        Cut_Off_Field_Name
        Cut_Off_Table_Name
        Database_Version
        DateFormat
        Days_Cutoff_Notice
        Days_Deposit_Notice
        DCOPTTaxRate
        DCPrefix
        Def_Law_Firm
        Def_Lending_Company
        Def_Rep_Path
        Def_Spec_Doc
        Def_Template_Path
        Default_Category
        DefaultCity
        DefaultFee
        DefaultProvince
        DefLotStatus
        dep_ar_cr_inv
        Deposit_credit
        Deposit_Debit
        Deposit_DSN
        DisableMiscApproved
        Display_CatDesc
        Document_Folder
        Draw
        Email
        EnableBankingLink
        EnableSelectionWeb
        Estimator_Email
        Export_Path
        ExportExcel_Path
        fax
        FileAttachmentBase
        fiscal_year
        french_contract
        GL_AR_Account
        GL_Cost_Dep_Account
        GL_Lot_Inv_Account
        GL_Lot_Pay_Account
        GL_Prefix
        GlobalJobIncr
        GLOPTTaxRate
        GST_Credit
        GST_No
        GST_Rate
        GST_Rebate_Credit
        GST_Rebate_Rate
        HFWebUDL
        HMSPWD
        HMSUID
        Holdback_Credit
        Holdback_Debit
        ID
        Import_Path
        includetax
        Interest_Cost_Code
        Interior_Viewable
        Job_Mask
        Job_Section1
        Job_Section2
        Job_Section3
        Job_Separator
        Job1Field
        Job1Incr
        Job1Ornt
        Job1Table
        Job2Field
        Job2Incr
        Job2Ornt
        Job2Table
        Job3Field
        Job3Incr
        Job3Ornt
        Job3Table
        Land_Category
        Land_Cost_Code
        LogFolder
        Lot_Inv_credit
        Lot_Inv_debit
        Lot_Inventory_Macro
        Lot_Inventory_Path
        Lot_LIAB_GL
        LotTaxRate
        Misc_Cost_Code
        ModelByArea
        ModelsByArea_Phase
        ModelSpecDocInWord
        ModelTaxRate
        Monitor_DT_TL
        Monitor_DT_TL_Field
        MonitorUserLog
        Mortgage_Credit
        Mortgage_Debit
        MSOPTTaxRate
        MultiFamily
        NETTAXRATE
        Next_Adjust_Seq
        Next_Batch
        Next_Comm_Seq
        Next_Customer_No
        Next_Dep_Seq
        Next_Job_Incr
        Next_Lot_No
        Next_Mort_Adv_Id
        Next_Mort_Seq
        NoImage
        OptionByArea
        OptionByAreaPhase
        GlobalOptionByArea
        GlobalOptionByAreaPhase
        DCOptionByArea
        DCOptionByAreaPhase
        phone
        PhoneMask
        Post_Adj_Macro
        Post_Adj_Path
        Post_Comm_Macro
        Post_Comm_Path
        Post_Deposit_Macro
        Post_Deposit_Path
        Post_Mortgage_Macro
        Post_Mortgage_Path
        POUPLOAD_DETAILS
        POUPLOAD_DSN
        POUPLOAD_FIELD
        POUPLOAD_FROM
        POUPLOAD_HEADER
        Process_Sales_Macro
        Process_Sales_Path
        province
        PST_Province
        PST_Rate
        PSTAcct
        PSTRebateAcct
        Req_Prefix_Job
        ReqCategory
        RevenueAdjGLAcct
        Sale_credit
        Sale_Debit
        Sales_Override
        SchRevenue
        ServiceModuleUDL
        Show_CDN_GST
        ShowMajorGroupDesc
        showMsgOnEachLoad
        ShowSpecDocument
        Tax_Group
        tax_id
        taxadjust_to_bldr
        Taxes_Cost_Code
        Tim_Category_Table
        Tim_Cost_Table
        Tim_Extra_Table
        Timberline_Data_Path
        Timberline_PWD
        Timberline_UID
        TL_AR_Data_Path
        TLJCACCTFORMAT
        TLMunicipal_Add
        UpdateDefLotStatus
        Use_Addition_Changes
        Use_french
        USE_POUPLOADWEB
        Use_Prefix
        Use_Timberline
        UseCM
        UseElevation
        UseHFWeb
        UseINCPrefix_Adj
        UseMajorGroup
        UsePO
        UsePst
        USER_DSN
        UseSchedule
        UseSelection
        UseSeqCustomer
        UseServiceModule
        UseSpecWordDoc
        UseUserLog
        Working_Progress
        zip


    
    [_UserOptions] '-------------------------
        SqlLogging
        SqlLogFile
        RptViewZoom
        RptViewTree
        POIndexView
        CloseAssemblyExportDialog
        ShowAssemblyExportReport
        ClosePricelistExportDialog
        PricelistExportPath
        PricelistExportSingleFile
        PricelistExportNextPrices
        PricelistImportSkipZeros
        PricelistImportDestination
        SalesWorksheetPhase
        
        ShowBudgetExportReport      'launch posting summary report when posting is complete
        ShowBudgetProcessingJournal 'launch processing journal when posting is complete
        ShowPOExportReport      'launch posting summary report when posting is complete
        ShowPOProcessingJournal 'launch processing journal when posting is complete
        
        RoundingFactor
        SalesPricesIncludeTax
        
    [_EndOfOptions]
    
End Enum

Private Sub Open3rdPartyDBs()
On Error Resume Next
    
    Dim s As String
    Dim rs As Recordset
    Dim userAccess As Long
    
    
    With HFApp
        s = ""
        
        Set .Databases(dbAccounting) = Nothing
        Set .Databases(dbAccountingCustom) = Nothing
        Set .Databases(dbAccountingDictionary) = Nothing
        Set .Databases(dbReceivables) = Nothing
        Set .Databases(dbReceivablesDictionary) = Nothing
        Set .Databases(dbEstimating) = Nothing
        Set .Databases(dbTimberlinePVdata) = Nothing
        .ConnectionString(dbAccounting) = ""
        .ConnectionString(dbAccountingCustom) = ""
        .ConnectionString(dbAccountingDictionary) = ""
        .ConnectionString(dbReceivables) = ""
        .ConnectionString(dbReceivablesDictionary) = ""
        .ConnectionString(dbEstimating) = ""
        .ConnectionString(dbTimberlinePVdata) = ""
        Set .Databases(dbAccounting) = New adodb.Connection
        Set .Databases(dbAccountingCustom) = New adodb.Connection
        Set .Databases(dbAccountingDictionary) = New adodb.Connection
        Set .Databases(dbReceivables) = New adodb.Connection
        Set .Databases(dbReceivablesDictionary) = New adodb.Connection
        Set .Databases(dbEstimating) = New adodb.Connection
        Set .Databases(dbTimberlinePVdata) = New adodb.Connection
        
        userAccess = Val("" & HFApp.SqlExec("SELECT User_Access FROM User_Manager WHERE User_Id=" & DbQuote(Str, HFApp.LoginID))(0))
        
        
        If IsIn(mModuleCode, "hfsched") Then
            Exit Sub
        End If
        
        Select Case Value(AccountingSystem)


            Case asTimberline
            
                .ConnectionString(dbAccounting) = TsConnectionString(Value(Timberline_Data_Path), Value(Timberline_UID), Value(Timberline_PWD), StandardNames, Accounting)
                .ConnectionString(dbAccountingCustom) = TsConnectionString(Value(Timberline_Data_Path), Value(Timberline_UID), Value(Timberline_PWD), CustomNames, Accounting)
                .ConnectionString(dbAccountingDictionary) = TsConnectionString(Value(Timberline_Data_Path), Value(Timberline_UID), Value(Timberline_PWD), DictionaryNames, Accounting)
                .ConnectionString(dbTimberlinePVdata) = "" & HFApp.Options.ValueByName("TLDSN")
                
                               
'                THIS DOESNT WORK...
'                If mProductCode = "hfpayables" Or HFApp.UserPermission("TaskPostDep") = True Or HFApp.UserPermission("CreateJob") = True Or HFApp.UserPermission("TaskUpdTotalJobCost") = True Or HFApp.UserPermission("PostBudgets") = True Or HFApp.UserPermission("PostCommitments") = True Or HFApp.UserPermission("SetupJobs") = True Then

                    StatusLabel = "opening Sage database"
                    Call .Databases(dbAccounting).Open(.ConnectionString(dbAccounting))
                    StatusLabel = "opening Sage custom names database"
                    Call .Databases(dbAccountingCustom).Open(.ConnectionString(dbAccountingCustom))
                    StatusLabel = "opening Sage dictionary"
                    Call .Databases(dbAccountingDictionary).Open(.ConnectionString(dbAccountingDictionary))
                    
                    If .Databases(dbAccounting).State <> adStateOpen Then
                        Select Case mModuleCode
                            Case "hfest", "hfpayables":                       s = s & "Sage Accounting:" & vbCrLf & Value(Timberline_Data_Path) & vbCrLf & Err.Description & vbCrLf & vbCrLf
                            Case "hfsales": If IsIn(userAccess, 1, 3, 7) Then s = s & "Sage Accounting:" & vbCrLf & Value(Timberline_Data_Path) & vbCrLf & Err.Description & vbCrLf & vbCrLf
                        End Select
                    End If
                    
                    If mModuleCode = "hfpayables" Then
                        If .ConnectionString(dbTimberlinePVdata) <> "" Then
                            If UCase(Mid(.ConnectionString(dbTimberlinePVdata), 1, 3)) <> "DSN" Then
                                .ConnectionString(dbTimberlinePVdata) = "DSN=" & .ConnectionString(dbTimberlinePVdata)
                            End If
                            Call .Databases(dbTimberlinePVdata).Open(.ConnectionString(dbTimberlinePVdata))
                        End If
                    End If
                    
                    If Trim(Value(TL_AR_Data_Path)) = "" Or Trim(Value(TL_AR_Data_Path)) = Trim(Value(Timberline_Data_Path)) Then
                        .ConnectionString(dbReceivables) = .ConnectionString(dbAccounting)
                        .ConnectionString(dbReceivablesDictionary) = .ConnectionString(dbAccountingDictionary)
                        Set .Databases(dbReceivables) = .Databases(dbAccounting)
                        Set .Databases(dbReceivablesDictionary) = .Databases(dbAccountingDictionary)
                    Else
                        .ConnectionString(dbReceivables) = TsConnectionString(Value(TL_AR_Data_Path), Value(Timberline_UID), Value(Timberline_PWD), StandardNames, Accounting)
                        .ConnectionString(dbReceivablesDictionary) = TsConnectionString(Value(TL_AR_Data_Path), Value(Timberline_UID), Value(Timberline_PWD), DictionaryNames, Accounting)
                        StatusLabel = "opening Sage AR database"
                        Call .Databases(dbReceivables).Open(.ConnectionString(dbReceivables))
                        StatusLabel = "opening Sage AR dictionary"
                        Call .Databases(dbReceivablesDictionary).Open(.ConnectionString(dbReceivablesDictionary))
                        If .Databases(dbReceivables).State <> adStateOpen Then
                        Select Case mModuleCode
                            Case "hfest", "hfpayables":                       s = s & "Sage Accounting (AR):" & vbCrLf & Value(TL_AR_Data_Path) & vbCrLf & Err.Description & vbCrLf & vbCrLf
                            Case "hfsales": If IsIn(userAccess, 1, 3, 7) Then s = s & "Sage Accounting (AR):" & vbCrLf & Value(TL_AR_Data_Path) & vbCrLf & Err.Description & vbCrLf & vbCrLf
                        End Select
                        End If
                    End If
'                End If
                
            Case asSimply
                
                        
            Case asMasterBuilder
                
                .ConnectionString(dbReceivables) = Sage100ConnectionString()
                .ConnectionString(dbReceivablesDictionary) = Sage100ConnectionString()
                .ConnectionString(dbAccounting) = Sage100ConnectionString()
                .ConnectionString(dbAccountingDictionary) = Sage100ConnectionString()
                StatusLabel = "opening Sage 100 database"
                Call .Databases(dbReceivables).Open(.ConnectionString(dbReceivables))
                Call .Databases(dbReceivablesDictionary).Open(.ConnectionString(dbReceivablesDictionary))
                Call .Databases(dbAccounting).Open(.ConnectionString(dbAccounting))
                Call .Databases(dbAccountingDictionary).Open(.ConnectionString(dbAccountingDictionary))
                If .Databases(dbAccounting).State <> adStateOpen Then
                Select Case mProductCode
                    Case "hfest", "hfpayables":                       s = s & "Sage 100:" & vbCrLf & Value(MasterBuilderDataFolder) & vbCrLf & Err.Description & vbCrLf & vbCrLf
                    Case "hfsales": If IsIn(userAccess, 1, 3, 7) Then s = s & "Sage 100:" & vbCrLf & Value(MasterBuilderDataFolder) & vbCrLf & Err.Description & vbCrLf & vbCrLf
                End Select
                End If
            
            
                
            
        End Select
        
        
        If ValueByName("EstimatingSystem") = esTimberline And ValueByName("SageSqlEstServer") <> "" Then
            'sage sql estimating
            .ConnectionString(dbEstimating) = SQLConnectionString(ValueByName("SageSqlEstServer"), ValueByName("SageSqlEstDatabase"), ValueByName("SageSqlEstUser"), ValueByName("SageSqlEstPswd"))
            StatusLabel = "opening Estimating database"
            Call .Databases(dbEstimating).Open(.ConnectionString(dbEstimating))
            StatusLabel = ""
            If .Databases(dbEstimating).State <> adStateOpen Then
            Select Case mModuleCode
                Case "hfest":     s = s & "Sage Estimating:" & vbCrLf & ValueByName("SageSqlEstServer") & vbCrLf & Err.Description & vbCrLf & vbCrLf
            End Select
            End If
        End If
        
    End With



On Error GoTo eh
    If s <> "" And Not Closing Then
        Call FLoginWarnings.ShowMessage(s)
    End If



    '------------------------------------------------------
    'get accounting formats and masks
    '------------------------------------------------------
    Dim s1 As Long
    Dim s2 As Long
    Dim s3 As Long
    Dim s4 As Long
    Dim sep As String
    Dim mask As String
    Dim SaveRequired As Boolean
    
    SaveRequired = False
    Select Case Value(AccountingSystem)
        Case asTimberline
            If HFApp.Databases(dbAccountingDictionary).State = adStateOpen Then
                
StatusLabel = "reading Sage formats"
                'job format
                s1 = 0
                s2 = 0
                s3 = 0
                mask = ""
                sep = ""
                Set rs = HFApp.SqlExec("select section,cslen,cnpunct from ts_ctl_record_3 where section>0 and fldcode=16", dbAccountingDictionary)
                While Not rs.EOF
                    Select Case Val("" & rs("section"))
                        Case 1: s1 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                        Case 2: s2 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                        Case 3: s3 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                    End Select
                    rs.MoveNext
                Wend
                mask = String(s1, "&") & sep & String(s2, "&") & sep & String(s3, "&")
                If sep <> "" Then
                    While Right(mask, 1) = sep
                        mask = left(mask, Len(mask) - 1)
                    Wend
                End If
                If s1 = 0 Then
                    s1 = 10
                    mask = "&&&&&&&&&&"
                End If
                If Value(Job_Mask) <> mask Then
                    If vbYes = MsgBox("The job number format in your system settings does not match Sage." & vbCrLf & "Do you want your system settings to be changed to match Sage?", vbQuestion + vbYesNo, "HomeFront") Then
                        SaveRequired = True
                        Value(Job_Mask) = mask
                        Value(Job_Separator) = sep
                        Value(Job_Section1) = s1
                        Value(Job_Section2) = s2
                        Value(Job_Section3) = s3
                    End If
                End If
                
                
                'cost code format
                s1 = 0
                s2 = 0
                s3 = 0
                s4 = 0
                mask = ""
                sep = ""
                Set rs = HFApp.SqlExec("select section,cslen,cnpunct from ts_ctl_record_3 where section>0 and fldcode=18", dbAccountingDictionary)
                While Not rs.EOF
                    Select Case Val("" & rs("section"))
                        Case 1: s1 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                        Case 2: s2 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                        Case 3: s3 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                        Case 4: s4 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                    End Select
                    rs.MoveNext
                Wend
                mask = String(s1, "&") & sep & String(s2, "&") & sep & String(s3, "&") & sep & String(s4, "&")
                If sep <> "" Then
                    While Right(mask, 1) = sep
                        mask = left(mask, Len(mask) - 1)
                    Wend
                End If
                If s1 = 0 Then
                    s1 = 12
                    mask = "&&&&&&&&&&&&"
                End If
                If Value(Cost_Code_Mask) <> mask Then
                    If vbYes = MsgBox("The cost code format in your system settings does not match Sage." & vbCrLf & "Do you want your system settings to be changed to match Sage?", vbQuestion + vbYesNo, "HomeFront") Then
                        SaveRequired = True
                        Value(Cost_Code_Mask) = mask
                        Value(Cost_Code_Separator) = sep
                        Value(Cost_Code_section1) = s1
                        Value(Cost_Code_section2) = s2
                        Value(Cost_Code_section3) = s3
                        Value(Cost_Code_section4) = s4
                    End If
                End If
                        
                
                'now get prefix usage flag
                s1 = 0
                Set rs = HFApp.SqlExec("select account_prefix_a_length+account_prefix_ab_length+account_prefix_abc_length as GLPrefixLength,use_prefix as JCUsePrefix from glm_master__account_format,jcm_master__jc_cost_controls", dbAccounting)
                If Val("" & rs("GLPrefixLength")) = 0 Then
                    s1 = 0 'prefixes not in use
                Else
                    If "" & rs("JCUsePrefix") = "True" Then
                        s1 = 2 'prefix goes in tl job's cost prefix field
                    Else
                        s1 = 1 'prefix is combined with wip acct and goes in tl job's cost account field
                    End If
                End If
                If Value(TLJCACCTFORMAT) <> s1 Then
                    If vbYes = MsgBox("The prefix usage settings in your system settings does not match Sage." & vbCrLf & "Do you want your system settings to be changed to match Sage?", vbQuestion + vbYesNo, "HomeFront") Then
                        SaveRequired = True
                        Value(TLJCACCTFORMAT) = s1
                    End If
                End If
            End If
            
            
            
            If HFApp.Databases(dbReceivablesDictionary).State = adStateOpen Then
                'customer format
                s1 = 0
                s2 = 0
                s3 = 0
                mask = ""
                sep = ""
                Set rs = HFApp.SqlExec("select section,cslen,cnpunct from ts_ctl_record_3 where section>0 and fldcode=300", dbReceivablesDictionary)
                While Not rs.EOF
                    Select Case Val("" & rs("section"))
                        Case 1: s1 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                        Case 2: s2 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                        Case 3: s3 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep = "" & rs("cnpunct")
                    End Select
                    rs.MoveNext
                Wend
                mask = String(s1, "&") & sep & String(s2, "&") & sep & String(s3, "&")
                If sep <> "" Then
                    While Right(mask, 1) = sep
                        mask = left(mask, Len(mask) - 1)
                    Wend
                End If
                If s1 = 0 Then
                    s1 = 10
                    mask = "&&&&&&&&&&"
                End If
                If Value(Customer_Mask) <> mask Then
                    If vbYes = MsgBox("The customer number format in your system settings does not match Sage." & vbCrLf & "Do you want your system settings to be changed to match Sage?", vbQuestion + vbYesNo, "HomeFront") Then
                        SaveRequired = True
                        Value(Customer_Mask) = mask
                        Value(Cust_Separator) = sep
                        Value(Cust_Section1) = s1
                        Value(Cust_Section2) = s2
                        Value(Cust_Section3) = s3
                    End If
                End If
            End If
StatusLabel = ""
            
            If SaveRequired Then
StatusLabel = "saving changes"
                Closing = True
                Call SaveData
                Closing = False
StatusLabel = ""
            End If
            
        Case Else
        'Do not forcefully set the values if the user does not use Timberline.  Allow them to set their own formats anyway.
            Set rs = HFApp.SqlExec("select ID,Cost_Code_Section1, Cost_Code_Section2, Cost_Code_Section3, Cost_Code_Section4, Cost_Code_Separator, Cost_Code_Mask, Job_Section1, Job_Section2, Job_Section3, Job_Separator, Job_Mask, Customer_Section1, Customer_Section2, Customer_Section3, Customer_Separator, Customer_Mask,Cust_Section1,Cust_Section2,Cust_Section3  from System_Setup where id = " & DbQuote(Num, HFApp.DivisionID), dbHomefront)
                
'            'job format
            Value(Job_Section1) = rs.fields("Job_Section1")
            Value(Job_Section2) = rs.fields("Job_Section2")
            Value(Job_Section3) = rs.fields("Job_Section3")
            Value(Job_Separator) = rs.fields("Job_Separator")
            Value(Job_Mask) = rs.fields("Job_Mask")

'            'cost code format
            Value(Cost_Code_section1) = rs.fields("Cost_Code_section1")
            Value(Cost_Code_section2) = rs.fields("Cost_Code_section2")
            Value(Cost_Code_section3) = rs.fields("Cost_Code_section3")
            Value(Cost_Code_section4) = rs.fields("Cost_Code_section4")
            Value(Cost_Code_Separator) = rs.fields("Cost_Code_Separator")
            Value(Cost_Code_Mask) = rs.fields("Cost_Code_Mask")

'            'now get prefix usage flag
            Value(TLJCACCTFORMAT) = 0 'prefixes not in use

'            'customer format
            Value(customer_section1) = rs.fields("customer_section1")
            Value(customer_section2) = rs.fields("customer_section2")
            Value(customer_section3) = rs.fields("customer_section3")
            Value(customer_Separator) = rs.fields("customer_Separator")
            Value(Customer_Mask) = rs.fields("Customer_Mask")

            Value(Cust_Section1) = Value(Cust_Section1)
            Value(Cust_Section2) = Value(Cust_Section2)
            Value(Cust_Section3) = Value(Cust_Section3)
            Value(Cust_Separator) = Value(customer_Separator)
            
    End Select
eh:
Exit Sub
End Sub


Public Property Get Count() As Long
    Count = [_EndOfOptions]
End Property
Public Property Get Name(Index As OptionsEnum) As String
    Name = mOptions(Index).Name
End Property

Public Property Get Value(Index As OptionsEnum)
Attribute Value.VB_UserMemId = 0
    Select Case Index
        Case Else
            Value = mOptions(Index).Value
    End Select
End Property

Public Property Let Value(Index As OptionsEnum, RHS)
    mOptions(Index).Dirty = mOptions(Index).Value <> RHS Or mOptions(Index).Dirty
    mOptions(Index).Value = RHS
End Property

Public Property Get ValueByName(OptionName As String)
On Error Resume Next
    Dim Index As Long
    Dim s As String
    Index = ItemIndex(OptionName)
    If Index <> -1 Then
        ValueByName = mOptions(ItemIndex(OptionName)).Value
    Else
        If OptionName = "IsProductionDatabase" Then
            s = IsProductionDatabase
        ElseIf OptionName = "IsProductionVerified" Then
            s = IsProductionVerified
        Else
            s = ""
            If s = "" Then s = "" & HFApp.SqlExec("select " & OptionName & " from system_setup where id=" & Val(HFApp.DivisionID), dbHomefront)(0)
            If s = "" Then s = "" & HFApp.SqlExec("select optionvalue from appoptions where isnull(divisionid,0) in(0," & DbQuote(Num, HFApp.DivisionID) & ") and optionname=" & DbQuote(Str, OptionName) & " order by divisionid desc", dbHomefront)(0)
        End If
        ValueByName = s
    End If
End Property

Public Property Let ValueByName(OptionName As String, RHS)
On Error Resume Next
    Dim Index As Long
    Index = ItemIndex(OptionName)
    If Index <> -1 Then
        mOptions(Index).Dirty = mOptions(Index).Value <> RHS Or mOptions(Index).Dirty
        mOptions(Index).Value = RHS
    End If
    
    'must save these immediately.
    Call HFApp.SqlExec("UPDATE system_setup SET " & OptionName & "=" & DbQuote(Str, RHS) & " where id=" & HFApp.DivisionID, dbHomefront)
    
    If IsIn(OptionName, "FaxServer", "FaxProvider", "FaxBillingCode", "Format_QtyRateEQZero_ForeColor", "Format_QtyRateEQZero_BackColor", "Format_QtyRateEQZero_FontStyle", "Format_QtyRateLTZero_ForeColor" _
                      , "Format_QtyRateLTZero_BackColor", "Format_QtyRateLTZero_FontStyle", "Format_InvalidData_ForeColor", "Format_InvalidData_BackColor", "Format_InvalidData_FontStyle", "Format_NonSysRate_ForeColor" _
                      , "Format_NonSysRate_BackColor", "Format_NonSysRate_FontStyle", "CRMClientID", "CRMApiKey", "CRMEnvironment") Then
        'global appoptions
        HFApp.SqlExec "INSERT INTO AppOptions(Divisionid,UID,OptionName,OptionValue,UStmp,TStmp) VALUES(0,''" & "," & DbQuote(Str, OptionName) & "," & DbQuote(Str, RHS) & "," & DbQuote(Str, HFApp.LoginID) & ",GETDATE())"
        HFApp.SqlExec "UPDATE AppOptions SET OptionValue=" & DbQuote(Str, RHS) & ", UStmp = " & DbQuote(Str, HFApp.LoginID) & ", TStmp = GETDATE()" & " WHERE UID='' AND OptionName=" & DbQuote(Str, OptionName)
    Else
        'division specific appoptions
        HFApp.SqlExec "INSERT INTO AppOptions(UID,OptionName,OptionValue,DivisionID,UStmp,TStmp) VALUES(''" & "," & DbQuote(Str, OptionName) & "," & DbQuote(Str, RHS) & "," & DbQuote(Str, HFApp.DivisionID) & "," & DbQuote(Str, HFApp.LoginID) & ",GETDATE())"
        HFApp.SqlExec "UPDATE AppOptions SET OptionValue=" & DbQuote(Str, RHS) & ", UStmp = " & DbQuote(Str, HFApp.LoginID) & ", TStmp = GETDATE()" & " WHERE UID='' AND DivisionID=" & DbQuote(Str, HFApp.DivisionID) & " and OptionName=" & DbQuote(Str, OptionName)
    End If

End Property




Public Sub ReadData()
On Error GoTo eh
    Dim rs As adodb.Recordset
    Dim Index As Long
    Dim i As Long
    Dim cfv As New CFileVersionInfo
    Dim s As String
    Dim errtext As String
    
    
    
StatusLabel = "reading system setup..."
    'read appoptions
    errtext = "Open AppOptions"
    
    s = ""
    s = s & "SELECT OptionName,OptionValue" & vbCrLf
    s = s & "  FROM AppOptions" & vbCrLf
    s = s & " WHERE (ISNULL(UID,'')='' OR UID=" & DbQuote(Str, HFApp.LoginID, False) & ")" & vbCrLf
    s = s & "   AND isnull(divisionid,0) in(0," & DbQuote(Num, HFApp.DivisionID) & ")" & vbCrLf
    s = s & "order by divisionid " & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    errtext = "Loop through AppOptions"
    While Not rs.EOF
        Index = ItemIndex("" & rs("OptionName"))
        If Index > -1 Then mOptions(Index).Value = "" & rs("OptionValue")
        rs.MoveNext
    Wend
    rs.Close
    
    
    Index = ItemIndex("SQLDate_Format")
    If Index > -1 Then mOptions(Index).Value = SQLDATEFORMAT
    
    'read system settings
    errtext = "Open System Setup"
    Set rs = HFApp.SqlExec("SELECT * FROM System_Setup where id=" & HFApp.DivisionID)
    If Not rs.EOF Then
        errtext = "Loop through system setup"
        For i = 0 To rs.fields.Count - 1
            Index = ItemIndex(rs.fields(i).Name)
            If Index > -1 Then mOptions(Index).Value = "" & rs(i)
        Next
        errtext = "Check COEntryPriortoPurchase"
        Index = ItemIndex("COEntryPriorToPurchase")
        If Index > -1 Then mOptions(Index).Value = "" & rs("COEntryPriorToPurch")
    
    End If
StatusLabel = ""
    
    For Index = [_AppOptions] To [_EndOfOptions]
        mOptions(Index).Dirty = False
    Next
    rs.Close
    
    
    
    errtext = "check Timberline data folder path"
    tlFolder = mOptions(Timberline_Data_Path).Value
    errtext = "Open 3rdPartyDbs"
    
    If Not mTrustedSkipAccounting Then
        Call Open3rdPartyDBs
    End If














Exit Sub

eh:
If Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description & vbCrLf & " " & errtext
End If
End Sub

Private Function ItemIndex(Name As String) As Long
    Dim i As Long
    ItemIndex = -1
    For i = LBound(mOptions) To UBound(mOptions)
        If UCase(mOptions(i).Name) = UCase(Name) Then
            ItemIndex = i
            Exit Function
        End If
    Next
End Function

Public Sub SaveData()
On Error Resume Next

    Dim i As Long
    Dim s As String
    For i = 1 To UBound(mOptions)
    
        If mOptions(i).Dirty Then
        
        
            Select Case i
          
                'opps - misspelled
                Case COEntryPriorToPurchase
                    HFApp.SqlExec "UPDATE system_setup SET COEntryPriorToPurch=" & DbQuote(Bit, mOptions(i).Value) & " where id=" & HFApp.DivisionID
                    
            
                'any non-string fields on system_setup have to be handled carefully
                Case SendFaxUsingPurchasersInfo, UsePricingFromEstimating, RemoveDetailsonSpecConvert, PromptForSpecOptions, COEntryPriorToPurchase, ShowAssignCustOnly:
                    HFApp.SqlExec "UPDATE system_setup SET " & mOptions(i).Name & "=" & DbQuote(Bit, mOptions(i).Value) & " where id=" & HFApp.DivisionID
                
                Case [_SystemSettings] To [_UserOptions]
                    HFApp.SqlExec "UPDATE system_setup SET " & mOptions(i).Name & "=" & DbQuote(Str, mOptions(i).Value) & " where id=" & HFApp.DivisionID
                
                'global appoptions
                Case FaxServer, FaxProvider, FaxBillingCode, Format_QtyRateEQZero_ForeColor, Format_QtyRateEQZero_BackColor, Format_QtyRateEQZero_FontStyle, Format_QtyRateLTZero_ForeColor, Format_QtyRateLTZero_BackColor, Format_QtyRateLTZero_FontStyle, Format_InvalidData_ForeColor, Format_InvalidData_BackColor, Format_InvalidData_FontStyle, Format_NonSysRate_ForeColor, Format_NonSysRate_BackColor, Format_NonSysRate_FontStyle
                    HFApp.SqlExec "INSERT INTO AppOptions(UID,OptionName,OptionValue,UStmp,TStmp) VALUES(''" & "," & DbQuote(Str, mOptions(i).Name) & "," & DbQuote(Str, mOptions(i).Value) & "," & DbQuote(Str, HFApp.LoginID) & ",GETDATE())"
                    HFApp.SqlExec "UPDATE AppOptions SET OptionValue=" & DbQuote(Str, mOptions(i).Value) & ", UStmp = " & DbQuote(Str, HFApp.LoginID) & ", TStmp = GETDATE()" & " WHERE UID='' AND OptionName=" & DbQuote(Str, mOptions(i).Name)
                
                'division specific appoption
                Case [_AppOptions] To [_SystemSettings]
                    HFApp.SqlExec "INSERT INTO AppOptions(UID,OptionName,OptionValue,DivisionID,UStmp,TStmp) VALUES(''" & "," & DbQuote(Str, mOptions(i).Name) & "," & DbQuote(Str, mOptions(i).Value) & "," & DbQuote(Str, HFApp.DivisionID) & "," & DbQuote(Str, HFApp.LoginID) & ",GETDATE())"
                    HFApp.SqlExec "UPDATE AppOptions SET OptionValue=" & DbQuote(Str, mOptions(i).Value) & ", UStmp = " & DbQuote(Str, HFApp.LoginID) & ", TStmp = GETDATE()" & " WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and UID='' AND OptionName=" & DbQuote(Str, mOptions(i).Name)
            
                'user specific appoptions
                Case [_UserOptions] To [_EndOfOptions]
                    HFApp.SqlExec "INSERT INTO AppOptions(UID,OptionName,OptionValue,UStmp,TStmp) VALUES(" & DbQuote(Str, HFApp.LoginID, False) & "," & DbQuote(Str, mOptions(i).Name) & "," & DbQuote(Str, mOptions(i).Value) & "," & DbQuote(Str, HFApp.LoginID) & ",GETDATE())"
                    HFApp.SqlExec "UPDATE AppOptions SET OptionValue=" & DbQuote(Str, mOptions(i).Value) & ", UStmp = " & DbQuote(Str, HFApp.LoginID) & ", TStmp = GETDATE()" & " WHERE UID=" & DbQuote(Str, HFApp.LoginID) & " AND OptionName=" & DbQuote(Str, mOptions(i).Name)
                    
            End Select
        End If
    Next
    
    If Not Closing Then Call Open3rdPartyDBs

End Sub

Private Sub Class_Terminate()
Closing = True
    SaveData
End Sub


Private Sub Class_Initialize()
    Dim i As Integer
    ReDim mOptions([_EndOfOptions]) As OptionRec
            
    '------------------------------------------------------------
    'App Options
    '------------------------------------------------------------
    i = i + 1: mOptions(i).Name = "SendPO_CCPrjMgr":                      mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "SendPO_ViaMapi":                       mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "SendPO_SMTPServer":                    mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "SendPO_FromCurrentUsersEmail":         mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "SQLDate_Format":                       mOptions(i).Value = "'YYYY-MM-DD'"
    
    i = i + 1: mOptions(i).Name = "WebUploadByCommunity"
    i = i + 1: mOptions(i).Name = "PurchasingRequiresSalesApproval"
    i = i + 1: mOptions(i).Name = "Msg_RFP_Subject"
    i = i + 1: mOptions(i).Name = "Msg_RFP_Body"
    i = i + 1: mOptions(i).Name = "LetPurchaserChangeSaleQty"
    i = i + 1: mOptions(i).Name = "FaxServer"
    i = i + 1: mOptions(i).Name = "FaxCoverPage"
    i = i + 1: mOptions(i).Name = "AccountingVersion"
    i = i + 1: mOptions(i).Name = "AppendUOMtoOptionDesc"
    
    i = i + 1: mOptions(i).Name = "UseCommunityStandards":                mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "CommunityStandardsArePhaseSpecific":   mOptions(i).Value = False
    
    i = i + 1: mOptions(i).Name = "RecordSalesSheetCosts":                mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "VendorPriceChangeCategory":            mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "UserControlOverBudgets":               mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "IncludeLandCostsInSalesSheetMarkup":   mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "UseJobAsMasterBuilderSubAccount":      mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "POPrefix":
    i = i + 1: mOptions(i).Name = "ZeroQtyTakeoffs":
        
    i = i + 1: mOptions(i).Name = "SalesSystem":                          mOptions(i).Value = SalesSystems.asHomeFront
    i = i + 1: mOptions(i).Name = "AccountingSystem":

    i = i + 1: mOptions(i).Name = "SimplyDataFile":
    i = i + 1: mOptions(i).Name = "QuickBooksDataFile":
    
    i = i + 1: mOptions(i).Name = "MasterBuilderDataFolder":
    i = i + 1: mOptions(i).Name = "MasterBuilderCompany":
    i = i + 1: mOptions(i).Name = "MasterBuilderUID":
    i = i + 1: mOptions(i).Name = "MasterBuilderPWD":
    
    i = i + 1: mOptions(i).Name = "FaxProvider":                          mOptions(i).Value = "Windows Fax Service"
    i = i + 1: mOptions(i).Name = "FaxBillingCode":                       mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "Format_QtyRateEQZero_ForeColor":       mOptions(i).Value = vbWindowText
    i = i + 1: mOptions(i).Name = "Format_QtyRateEQZero_BackColor":       mOptions(i).Value = 16773869 'light blue
    i = i + 1: mOptions(i).Name = "Format_QtyRateEQZero_FontStyle":       mOptions(i).Value = "Regular"
    i = i + 1: mOptions(i).Name = "Format_QtyRateLTZero_ForeColor":       mOptions(i).Value = vbWindowText
    i = i + 1: mOptions(i).Name = "Format_QtyRateLTZero_BackColor":       mOptions(i).Value = 12640511 'light pink
    i = i + 1: mOptions(i).Name = "Format_QtyRateLTZero_FontStyle":       mOptions(i).Value = "Regular"
    i = i + 1: mOptions(i).Name = "Format_InvalidData_ForeColor":         mOptions(i).Value = vbRed
    i = i + 1: mOptions(i).Name = "Format_InvalidData_BackColor":         mOptions(i).Value = vbWindowBackground
    i = i + 1: mOptions(i).Name = "Format_InvalidData_FontStyle":         mOptions(i).Value = "Bold"
    i = i + 1: mOptions(i).Name = "Format_NonSysRate_ForeColor":          mOptions(i).Value = 8404992
    i = i + 1: mOptions(i).Name = "Format_NonSysRate_BackColor":          mOptions(i).Value = vbWindowBackground
    i = i + 1: mOptions(i).Name = "Format_NonSysRate_FontStyle":          mOptions(i).Value = "Bold Italic"
    i = i + 1: mOptions(i).Name = "UsePricingFromEstimating":             mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "ZeroRateOnRefreshCosts":               mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ZeroRateOnChangeVendor":               mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "GroupPurchaseOrdersByVendor":          mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "TimberlineEstimatingPath":             mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "TimberlineEstimatingDictionaryPath":   mOptions(i).Value = "C:\Program Files\Timberline Office\Estimating"
    i = i + 1: mOptions(i).Name = "EstimatingUser":
    i = i + 1: mOptions(i).Name = "EstimatingPswd":
    i = i + 1: mOptions(i).Name = "EstimatingVersion":
    i = i + 1: mOptions(i).Name = "EstimatingMajorVersion":
    i = i + 1: mOptions(i).Name = "EstimatingMinorVersion":
    i = i + 1: mOptions(i).Name = "ImportDateARCustomers":                mOptions(i).Value = DateValue("Jan 1, 1950")
    i = i + 1: mOptions(i).Name = "ImportDateStandardCostCodes":          mOptions(i).Value = DateValue("Jan 1, 1950")
    i = i + 1: mOptions(i).Name = "ImportDateEstItems":                   mOptions(i).Value = DateValue("Jan 1, 1950")
    i = i + 1: mOptions(i).Name = "ImportDateTaxGroups":                  mOptions(i).Value = DateValue("Jan 1, 1950")
    i = i + 1: mOptions(i).Name = "ImportDateVendors":                    mOptions(i).Value = DateValue("Jan 1, 1950")
    i = i + 1: mOptions(i).Name = "ImportDateAccounts":                   mOptions(i).Value = DateValue("Jan 1, 1950")
    
    i = i + 1: mOptions(i).Name = "TaxGroupRequired":                     mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "TaxGroupLabour"
    i = i + 1: mOptions(i).Name = "TaxGroupMaterial"
    i = i + 1: mOptions(i).Name = "TaxGroupSubContract"
    i = i + 1: mOptions(i).Name = "TaxGroupEquipment"
    i = i + 1: mOptions(i).Name = "TaxGroupOverhead"
    i = i + 1: mOptions(i).Name = "TaxGroupOther"
    i = i + 1: mOptions(i).Name = "IncentivePhase":                       mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "IncentiveItem":                        mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "IncentiveItemDescription":             mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "SelectAtTakeoffItem":                  mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "SelectAtTakeoffPhase":                 mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "SelectAtTakeoffDescription":           mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "IncentiveCostPercent":                 mOptions(i).Value = 55
    i = i + 1: mOptions(i).Name = "LandPhase":                            mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "LandItem":                             mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "LandItemDescription":                  mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "GlobalPONumberSequence":               mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "POFormat"
    i = i + 1: mOptions(i).Name = "POFormatSegment1":                     mOptions(i).Value = 8
    i = i + 1: mOptions(i).Name = "POFormatSegment2":                     mOptions(i).Value = "/"
    i = i + 1: mOptions(i).Name = "POFormatSegment3":                     mOptions(i).Value = 3
    i = i + 1: mOptions(i).Name = "AdjustSellPriceWithCosts":             mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "LockPostedSalesWorksheets":            mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "AllowForecastsWhenRefreshingPOs":      mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "WarnForecastsWhenRefreshingPOs":       mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "PostSummarizedBudgets":                mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "PostSummarizedPOs":                    mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "IssueBudgetsByPhase":                  mOptions(i).Value = False
    
    i = i + 1: mOptions(i).Name = "UseEstimatingWorksheetsOnly":          mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "MarketingWorkSheetAutoRefreshCosts":   mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "BarCodeDCRoomRequired":                mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "CanAddContractItemsFromPurchasing":    mOptions(i).Value = True


    i = i + 1: mOptions(i).Name = "Msg_CancelPO_Subject":                 mOptions(i).Value = "Purchase order <%PONumber%> has been CANCELLED"
    i = i + 1: mOptions(i).Name = "Msg_CancelPO_Body":                    mOptions(i).Value = "Please be advised that purchase order <%PONumber%> has been CANCELLED."
    i = i + 1: mOptions(i).Name = "Msg_SendPO_Subject":                   mOptions(i).Value = "<%PONumber%>"
    i = i + 1: mOptions(i).Name = "Msg_SendPO_Body":                      mOptions(i).Value = "Please review the attached purchase order(s). For more information contact the project manager."
    i = i + 1: mOptions(i).Name = "Msg_SendPO_CC_PrjMgr":                 mOptions(i).Value = True
        
    
    '------------------------------------------------------------
    'System Settings
    '------------------------------------------------------------
    i = i + 1
    i = i + 1: mOptions(i).Name = "COEntryPriorToPurchase":               mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ShowAssignCustOnly":                   mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "RemoveDetailsonSpecConvert":           mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "PromptForSpecOptions":                 mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "SendFaxUsingPurchasersInfo":           mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "accounting_approve"
    i = i + 1: mOptions(i).Name = "Acct_Receivable"
    i = i + 1: mOptions(i).Name = "ACPrefix"
    i = i + 1: mOptions(i).Name = "Address1"
    i = i + 1: mOptions(i).Name = "Address2"
    i = i + 1: mOptions(i).Name = "Adjustment_Credit"
    i = i + 1: mOptions(i).Name = "allow_customprice"
    i = i + 1: mOptions(i).Name = "AlternatingGridColor"
    i = i + 1: mOptions(i).Name = "ApplyLotCostJob"
    i = i + 1: mOptions(i).Name = "ApproveCOReqQuote"
    i = i + 1: mOptions(i).Name = "AreaSpecSalesPerson"
    i = i + 1: mOptions(i).Name = "Backup_Folder"
    i = i + 1: mOptions(i).Name = "BalanceSheetMaxAcct"
    i = i + 1: mOptions(i).Name = "Cardel"
    i = i + 1: mOptions(i).Name = "City"
    i = i + 1: mOptions(i).Name = "BuilderCode"
    i = i + 1: mOptions(i).Name = "BuilderName"
    i = i + 1: mOptions(i).Name = "WebAdminUser"
    i = i + 1: mOptions(i).Name = "WebUploadPath"
    i = i + 1: mOptions(i).Name = "WebUploadUser"
    i = i + 1: mOptions(i).Name = "WebUploadPswd"
    i = i + 1: mOptions(i).Name = "Country"
    i = i + 1: mOptions(i).Name = "ContactPerson"
    i = i + 1: mOptions(i).Name = "Comm_Category"
    i = i + 1: mOptions(i).Name = "Comm_Cost_code"
    i = i + 1: mOptions(i).Name = "Comm_Type"
    i = i + 1: mOptions(i).Name = "Community_Cost_Code"
    i = i + 1: mOptions(i).Name = "CompanyLogo"
    i = i + 1: mOptions(i).Name = "CompanyName"
    i = i + 1: mOptions(i).Name = "Contract_Document"
    i = i + 1: mOptions(i).Name = "ContractDocInWord"
    i = i + 1: mOptions(i).Name = "COPrefix"
    i = i + 1: mOptions(i).Name = "CORevenue"
    i = i + 1: mOptions(i).Name = "Cost_Code_Mask"
    i = i + 1: mOptions(i).Name = "Cost_Code_Section1"
    i = i + 1: mOptions(i).Name = "Cost_Code_Section2"
    i = i + 1: mOptions(i).Name = "Cost_Code_Section3"
    i = i + 1: mOptions(i).Name = "Cost_Code_Section4"
    i = i + 1: mOptions(i).Name = "Cost_Code_Separator"
    i = i + 1: mOptions(i).Name = "CostCodeSaleClosing"
    i = i + 1: mOptions(i).Name = "CreateJobLot"
    i = i + 1: mOptions(i).Name = "Cust_Section1"
    i = i + 1: mOptions(i).Name = "Cust_Section2"
    i = i + 1: mOptions(i).Name = "Cust_Section3"
    i = i + 1: mOptions(i).Name = "Cust_Separator"
    i = i + 1: mOptions(i).Name = "Customer_Mask"
    i = i + 1: mOptions(i).Name = "Customer_Section1"
    i = i + 1: mOptions(i).Name = "Customer_Section2"
    i = i + 1: mOptions(i).Name = "Customer_Section3"
    i = i + 1: mOptions(i).Name = "Customer_Separator"
    i = i + 1: mOptions(i).Name = "CustomOptionsTaxRate"
    i = i + 1: mOptions(i).Name = "Cut_Off_Field_Name"
    i = i + 1: mOptions(i).Name = "Cut_Off_Table_Name"
    i = i + 1: mOptions(i).Name = "Database_Version"
    i = i + 1: mOptions(i).Name = "DateFormat"
    i = i + 1: mOptions(i).Name = "Days_Cutoff_Notice"
    i = i + 1: mOptions(i).Name = "Days_Deposit_Notice"
    i = i + 1: mOptions(i).Name = "DCOPTTaxRate"
    i = i + 1: mOptions(i).Name = "DCPrefix"
    i = i + 1: mOptions(i).Name = "Def_Law_Firm"
    i = i + 1: mOptions(i).Name = "Def_Lending_Company"
    i = i + 1: mOptions(i).Name = "Def_Rep_Path"
    i = i + 1: mOptions(i).Name = "Def_Spec_Doc"
    i = i + 1: mOptions(i).Name = "Def_Template_Path"
    i = i + 1: mOptions(i).Name = "Default_Category"
    i = i + 1: mOptions(i).Name = "DefaultCity"
    i = i + 1: mOptions(i).Name = "DefaultFee"
    i = i + 1: mOptions(i).Name = "DefaultProvince"
    i = i + 1: mOptions(i).Name = "DefLotStatus"
    i = i + 1: mOptions(i).Name = "dep_ar_cr_inv"
    i = i + 1: mOptions(i).Name = "Deposit_credit"
    i = i + 1: mOptions(i).Name = "Deposit_Debit"
    i = i + 1: mOptions(i).Name = "Deposit_DSN"
    i = i + 1: mOptions(i).Name = "DisableMiscApproved"
    i = i + 1: mOptions(i).Name = "Display_CatDesc"
    i = i + 1: mOptions(i).Name = "Document_Folder"
    i = i + 1: mOptions(i).Name = "Draw"
    i = i + 1: mOptions(i).Name = "Email"
    i = i + 1: mOptions(i).Name = "EnableBankingLink"
    i = i + 1: mOptions(i).Name = "EnableSelectionWeb"
    i = i + 1: mOptions(i).Name = "Estimator_Email"
    i = i + 1: mOptions(i).Name = "Export_Path"
    i = i + 1: mOptions(i).Name = "ExportExcel_Path"
    i = i + 1: mOptions(i).Name = "FAX"
    i = i + 1: mOptions(i).Name = "FileAttachmentBase"
    i = i + 1: mOptions(i).Name = "fiscal_year"
    i = i + 1: mOptions(i).Name = "french_contract"
    i = i + 1: mOptions(i).Name = "GL_AR_Account"
    i = i + 1: mOptions(i).Name = "GL_Cost_Dep_Account"
    i = i + 1: mOptions(i).Name = "GL_Lot_Inv_Account"
    i = i + 1: mOptions(i).Name = "GL_Lot_Pay_Account"
    i = i + 1: mOptions(i).Name = "GL_Prefix"
    i = i + 1: mOptions(i).Name = "GlobalJobIncr"
    i = i + 1: mOptions(i).Name = "GLOPTTaxRate"
    i = i + 1: mOptions(i).Name = "GST_Credit"
    i = i + 1: mOptions(i).Name = "GST_No"
    i = i + 1: mOptions(i).Name = "GST_Rate"
    i = i + 1: mOptions(i).Name = "GST_Rebate_Credit"
    i = i + 1: mOptions(i).Name = "GST_Rebate_Rate"
    i = i + 1: mOptions(i).Name = "HFWebUDL"
    i = i + 1: mOptions(i).Name = "HMSPWD"
    i = i + 1: mOptions(i).Name = "HMSUID"
    i = i + 1: mOptions(i).Name = "Holdback_Credit"
    i = i + 1: mOptions(i).Name = "Holdback_Debit"
    i = i + 1: mOptions(i).Name = "ID"
    i = i + 1: mOptions(i).Name = "Import_Path"
    i = i + 1: mOptions(i).Name = "IncludeTax"
    i = i + 1: mOptions(i).Name = "Interest_Cost_Code"
    i = i + 1: mOptions(i).Name = "Interior_Viewable"
    i = i + 1: mOptions(i).Name = "Job_Mask"
    i = i + 1: mOptions(i).Name = "Job_Section1"
    i = i + 1: mOptions(i).Name = "Job_Section2"
    i = i + 1: mOptions(i).Name = "Job_Section3"
    i = i + 1: mOptions(i).Name = "Job_Separator"
    i = i + 1: mOptions(i).Name = "Job1Field"
    i = i + 1: mOptions(i).Name = "Job1Incr"
    i = i + 1: mOptions(i).Name = "Job1Ornt"
    i = i + 1: mOptions(i).Name = "Job1Table"
    i = i + 1: mOptions(i).Name = "Job2Field"
    i = i + 1: mOptions(i).Name = "Job2Incr"
    i = i + 1: mOptions(i).Name = "Job2Ornt"
    i = i + 1: mOptions(i).Name = "Job2Table"
    i = i + 1: mOptions(i).Name = "Job3Field"
    i = i + 1: mOptions(i).Name = "Job3Incr"
    i = i + 1: mOptions(i).Name = "Job3Ornt"
    i = i + 1: mOptions(i).Name = "Job3Table"
    i = i + 1: mOptions(i).Name = "Land_Category"
    i = i + 1: mOptions(i).Name = "Land_Cost_Code"
    i = i + 1: mOptions(i).Name = "LogFolder"
    i = i + 1: mOptions(i).Name = "Lot_Inv_credit"
    i = i + 1: mOptions(i).Name = "Lot_Inv_debit"
    i = i + 1: mOptions(i).Name = "Lot_Inventory_Macro"
    i = i + 1: mOptions(i).Name = "Lot_Inventory_Path"
    i = i + 1: mOptions(i).Name = "Lot_LIAB_GL"
    i = i + 1: mOptions(i).Name = "LotTaxRate"
    i = i + 1: mOptions(i).Name = "Misc_Cost_Code"
    i = i + 1: mOptions(i).Name = "ModelByArea":                         mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ModelsByArea_Phase":                  mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ModelSpecDocInWord"
    i = i + 1: mOptions(i).Name = "ModelTaxRate"
    i = i + 1: mOptions(i).Name = "Monitor_DT_TL"
    i = i + 1: mOptions(i).Name = "Monitor_DT_TL_Field"
    i = i + 1: mOptions(i).Name = "MonitorUserLog"
    i = i + 1: mOptions(i).Name = "Mortgage_Credit"
    i = i + 1: mOptions(i).Name = "Mortgage_Debit"
    i = i + 1: mOptions(i).Name = "MSOPTTaxRate"
    i = i + 1: mOptions(i).Name = "MultiFamily"
    i = i + 1: mOptions(i).Name = "NETTAXRATE"
    i = i + 1: mOptions(i).Name = "Next_Adjust_Seq"
    i = i + 1: mOptions(i).Name = "Next_Batch"
    i = i + 1: mOptions(i).Name = "Next_Comm_Seq"
    i = i + 1: mOptions(i).Name = "Next_Customer_No"
    i = i + 1: mOptions(i).Name = "Next_Dep_Seq"
    i = i + 1: mOptions(i).Name = "Next_Job_Incr"
    i = i + 1: mOptions(i).Name = "Next_Lot_No"
    i = i + 1: mOptions(i).Name = "Next_Mort_Adv_Id"
    i = i + 1: mOptions(i).Name = "Next_Mort_Seq"
    i = i + 1: mOptions(i).Name = "NoImage"
    i = i + 1: mOptions(i).Name = "OptionByArea":                        mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "OptionByAreaPhase":                   mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "GlobalOptionByArea":                  mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "GlobalOptionByAreaPhase":             mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "DCOptionByArea":                      mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "DCOptionByAreaPhase":                 mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "Phone"
    i = i + 1: mOptions(i).Name = "PhoneMask"
    i = i + 1: mOptions(i).Name = "Post_Adj_Macro"
    i = i + 1: mOptions(i).Name = "Post_Adj_Path"
    i = i + 1: mOptions(i).Name = "Post_Comm_Macro"
    i = i + 1: mOptions(i).Name = "Post_Comm_Path"
    i = i + 1: mOptions(i).Name = "Post_Deposit_Macro"
    i = i + 1: mOptions(i).Name = "Post_Deposit_Path"
    i = i + 1: mOptions(i).Name = "Post_Mortgage_Macro"
    i = i + 1: mOptions(i).Name = "Post_Mortgage_Path"
    i = i + 1: mOptions(i).Name = "POUPLOAD_DETAILS"
    i = i + 1: mOptions(i).Name = "POUPLOAD_DSN"
    i = i + 1: mOptions(i).Name = "POUPLOAD_FIELD"
    i = i + 1: mOptions(i).Name = "POUPLOAD_FROM"
    i = i + 1: mOptions(i).Name = "POUPLOAD_HEADER"
    i = i + 1: mOptions(i).Name = "Process_Sales_Macro"
    i = i + 1: mOptions(i).Name = "Process_Sales_Path"
    i = i + 1: mOptions(i).Name = "Province"
    i = i + 1: mOptions(i).Name = "PST_Province"
    i = i + 1: mOptions(i).Name = "pst_rate"
    i = i + 1: mOptions(i).Name = "PSTAcct"
    i = i + 1: mOptions(i).Name = "PSTRebateAcct"
    i = i + 1: mOptions(i).Name = "Req_Prefix_Job"
    i = i + 1: mOptions(i).Name = "ReqCategory"
    i = i + 1: mOptions(i).Name = "RevenueAdjGLAcct"
    i = i + 1: mOptions(i).Name = "Sale_credit"
    i = i + 1: mOptions(i).Name = "Sale_Debit"
    i = i + 1: mOptions(i).Name = "Sales_Override"
    i = i + 1: mOptions(i).Name = "SchRevenue"
    i = i + 1: mOptions(i).Name = "ServiceModuleUDL"
    i = i + 1: mOptions(i).Name = "Show_CDN_GST"
    i = i + 1: mOptions(i).Name = "ShowMajorGroupDesc"
    i = i + 1: mOptions(i).Name = "showMsgOnEachLoad"
    i = i + 1: mOptions(i).Name = "ShowSpecDocument"
    i = i + 1: mOptions(i).Name = "Tax_Group"
    i = i + 1: mOptions(i).Name = "tax_id"
    i = i + 1: mOptions(i).Name = "taxadjust_to_bldr"
    i = i + 1: mOptions(i).Name = "Taxes_Cost_Code"
    i = i + 1: mOptions(i).Name = "Tim_Category_Table"
    i = i + 1: mOptions(i).Name = "Tim_Cost_Table"
    i = i + 1: mOptions(i).Name = "Tim_Extra_Table"
    i = i + 1: mOptions(i).Name = "Timberline_Data_Path"
    i = i + 1: mOptions(i).Name = "Timberline_PWD"
    i = i + 1: mOptions(i).Name = "Timberline_UID"
    i = i + 1: mOptions(i).Name = "TL_AR_Data_Path"
    i = i + 1: mOptions(i).Name = "TLJCACCTFORMAT"
    i = i + 1: mOptions(i).Name = "TLMunicipal_Add"
    i = i + 1: mOptions(i).Name = "UpdateDefLotStatus"
    i = i + 1: mOptions(i).Name = "Use_Addition_Changes"
    i = i + 1: mOptions(i).Name = "Use_french"
    i = i + 1: mOptions(i).Name = "USE_POUPLOADWEB"
    i = i + 1: mOptions(i).Name = "Use_Prefix"
    i = i + 1: mOptions(i).Name = "Use_Timberline"
    i = i + 1: mOptions(i).Name = "UseCM"
    i = i + 1: mOptions(i).Name = "UseElevation"
    i = i + 1: mOptions(i).Name = "UseHFWeb"
    i = i + 1: mOptions(i).Name = "UseINCPrefix_Adj"
    i = i + 1: mOptions(i).Name = "UseMajorGroup"
    i = i + 1: mOptions(i).Name = "UsePO"
    i = i + 1: mOptions(i).Name = "UsePst"
    i = i + 1: mOptions(i).Name = "USER_DSN"
    i = i + 1: mOptions(i).Name = "UseSchedule"
    i = i + 1: mOptions(i).Name = "UseSelection"
    i = i + 1: mOptions(i).Name = "UseSeqCustomer"
    i = i + 1: mOptions(i).Name = "UseServiceModule"
    i = i + 1: mOptions(i).Name = "UseSpecWordDoc"
    i = i + 1: mOptions(i).Name = "UseUserLog"
    i = i + 1: mOptions(i).Name = "Working_Progress"
    i = i + 1: mOptions(i).Name = "Zip"
        

    '------------------------------------------------------------
    'User Options
    '------------------------------------------------------------
    i = i + 1
    i = i + 1: mOptions(i).Name = "SqlLogging":                           mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "SqlLogFile":                           mOptions(i).Value = App.Path & "\SqlLog.txt"
    i = i + 1: mOptions(i).Name = "RptViewZoom":
    i = i + 1: mOptions(i).Name = "RptViewTree":
    i = i + 1: mOptions(i).Name = "POIndexView":
    i = i + 1: mOptions(i).Name = "CloseAssemblyExportDialog":            mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "ShowAssemblyExportReport":             mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "ClosePricelistExportDialog":           mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "PricelistExportPath":                  mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "PricelistExportSingleFile":            mOptions(i).Value = False
    i = i + 1: mOptions(i).Name = "PricelistExportNextPrices":            mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "PricelistImportSkipZeros":             mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "PricelistImportDestination":           mOptions(i).Value = 0 'current
    i = i + 1: mOptions(i).Name = "SalesWorksheetPhase":                  mOptions(i).Value = ""
    i = i + 1: mOptions(i).Name = "ShowBudgetExportReport":               mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "ShowBudgetProcessingJournal":          mOptions(i).Value = True

    i = i + 1: mOptions(i).Name = "ShowPOExportReport":                   mOptions(i).Value = True
    i = i + 1: mOptions(i).Name = "ShowPOProcessingJournal":              mOptions(i).Value = True
        
    i = i + 1: mOptions(i).Name = "RoundingFactor":                       mOptions(i).Value = 0
    i = i + 1: mOptions(i).Name = "SalesPricesIncludeTax":                mOptions(i).Value = True
             
             
             

    For i = 0 To [_EndOfOptions]
        If Len(mOptions(i).Name) >= 35 Then MsgBox mOptions(i).Name & " is too long (" & Len(mOptions(i).Name) & ") max is 35", vbCritical
    Next

End Sub



