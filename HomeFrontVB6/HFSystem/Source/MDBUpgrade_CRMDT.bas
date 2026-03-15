Attribute VB_Name = "MDBUpgrade_CRMDT"
Option Explicit

Public Function DT_Get_DataForCRM() As String
    Dim s As String
    
    s = ""
    s = s & "create PROCEDURE [dbo].[DT_Get_DataForCRM]" & vbCrLf
    s = s & "AS" & vbCrLf
    s = s & "    BEGIN" & vbCrLf
    s = s & "" & vbCrLf
    s = s & " insert into appoptions(divisionid,UID,OptionName,OptionValue) " & vbCrLf
    s = s & "    Select 1,'', a.OptionName,'2020-01-01'" & vbCrLf
    s = s & "    from (select 'LastUpdate<tblModels>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblSeries>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tbllocality>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<Divisions>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<CommunityPhase>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblLotInventory>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblMajorGroups>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblCategories>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tbloptions>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblGlobaloptions>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblInventoryHomes>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblInventoryHomesAddendum>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblQuotesAndSalesContract>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblChangeOrder>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<Realtors>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblMultifamilyUnit>' as OptionName" & vbCrLf
    s = s & "       union all" & vbCrLf
    s = s & "       select 'LastUpdate<tblJobStatus>' as OptionName" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ) a" & vbCrLf
    s = s & "    left outer join appoptions b " & vbCrLf
    s = s & "     on b.optionname = a.OptionName " & vbCrLf
    s = s & "    where b.optionname is null;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "   -- Get LastUpdate Dates" & vbCrLf
    s = s & "        DECLARE @tblLastUpdate TABLE" & vbCrLf
    s = s & "            (" & vbCrLf
    s = s & "              TableName NVARCHAR(50)" & vbCrLf
    s = s & "            , LastUpdateDate DATETIME" & vbCrLf
    s = s & "            );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        INSERT  INTO @tblLastUpdate" & vbCrLf
    s = s & "                ( TableName" & vbCrLf
    s = s & "                , LastUpdateDate" & vbCrLf
    s = s & "                )" & vbCrLf
    s = s & "                SELECT DISTINCT" & vbCrLf
    s = s & "                        REPLACE(REPLACE(OptionName, 'LastUpdate<', ''), '>', '')" & vbCrLf
    s = s & "                      , CAST (OptionValue AS DATETIME)" & vbCrLf
    s = s & "                FROM    AppOptions" & vbCrLf
    s = s & "                WHERE   OptionName IN ( 'LastUpdate<tblModels>', 'LastUpdate<tblSeries>', 'LastUpdate<tbllocality>'," & vbCrLf
    s = s & "                                        'LastUpdate<Divisions>', 'LastUpdate<CommunityPhase>'," & vbCrLf
    s = s & "                                        'LastUpdate<tblLotInventory>', 'LastUpdate<tblMajorGroups>'," & vbCrLf
    s = s & "                                        'LastUpdate<tblCategories>', 'LastUpdate<tbloptions>'," & vbCrLf
    s = s & "                                        'LastUpdate<tblGlobaloptions>', 'LastUpdate<tblInventoryHomes>'," & vbCrLf
    s = s & "                                        'LastUpdate<tblInventoryHomesAddendum>', 'LastUpdate<tblQuotesAndSalesContract>'," & vbCrLf
    s = s & "                                        'LastUpdate<tblChangeOrder>', 'LastUpdate<Realtors>'," & vbCrLf
    s = s & "                                        'LastUpdate<tblMultifamilyUnit>','LastUpdate<tblJobStatus>' ) AND divisionid = 1;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                                       -- List of tables and their names" & vbCrLf
    s = s & "        DECLARE @TableNames AS TABLE ( Name NVARCHAR(200) );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        INSERT  INTO @TableNames" & vbCrLf
    s = s & "                ( Name )" & vbCrLf
    s = s & "        VALUES  ( 'tblSystemSetup' )," & vbCrLf
    s = s & "               ( 'Divisions' )," & vbCrLf
    s = s & "                ( 'DivisionCommunities' )," & vbCrLf
    s = s & "                ( 'tblLotStatus' )," & vbCrLf
    s = s & "                ( 'tblmodels' )," & vbCrLf
    s = s & "                ( 'tblseries' )," & vbCrLf
    s = s & "                ( 'master_date' )," & vbCrLf
    s = s & "                ( 'CommunityPhase' )," & vbCrLf
    s = s & "                ( 'lot_type' )," & vbCrLf
    s = s & "                ( 'lotpremium' )," & vbCrLf
    s = s & "                ( 'lotpremiummaster' )," & vbCrLf
    s = s & "                ( 'tblLotInventory' )," & vbCrLf
    s = s & "                ( 'tblCategories' )," & vbCrLf
    s = s & "                ( 'tbllocality' )," & vbCrLf
    s = s & "                ( 'tblMajorGroups' )," & vbCrLf
    s = s & "                ( 'tblOptions' )," & vbCrLf
    s = s & "                ( 'tblGlobalOptions' )," & vbCrLf
    s = s & "                ( 'tblUsers' )," & vbCrLf
    s = s & "                ( 'tblUserDivision' )," & vbCrLf
    s = s & "               ( 'tblUserCommunity' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomes' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomesAddendum' )," & vbCrLf
    s = s & "                ( 'tblHFCRMEntity' )," & vbCrLf
    s = s & "                ( 'tblHFCRMEntityDetail' )," & vbCrLf
    s = s & "                ( 'tblChangeOrderMaster' )," & vbCrLf
    s = s & "                ( 'tblChangeOrderDetails' )," & vbCrLf
    s = s & "                ( 'tblPurchaseType' )," & vbCrLf
    s = s & "                ( 'tblAttributeLists' )," & vbCrLf
    s = s & "                ( 'tblAttributeListValues' )," & vbCrLf
    s = s & "                ( 'tblRealtor' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomesCOMaster' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomesCODetails' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomeSchedulingDates' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomeConditionDates' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomeDepositDates' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomesDCMaster' )," & vbCrLf
    s = s & "                ( 'tblInventoryHomesDCDetails' )," & vbCrLf
    s = s & "                ( 'tblMultifamilyUnit' )," & vbCrLf
    s = s & "               ( 'DeletedOptionsforCRM' )," & vbCrLf
    s = s & "               ( 'tblJobStatus' )," & vbCrLf
    s = s & "               ( 'tblLastUpdate' )," & vbCrLf
    s = s & "               ( 'tblTitleCompany' )," & vbCrLf
    s = s & "               ( 'DeletedModelAndOptionsForCRM' )," & vbCrLf
    s = s & "               ( 'RoomMaster' )," & vbCrLf
    s = s & "               ( 'RoomMasterSubCategory' )," & vbCrLf
    s = s & "               ( 'RoomQtyByModel' )," & vbCrLf
    s = s & "               --('CustomerDesignSelections' )" & vbCrLf
    s = s & "               ( 'tblCustomers' )," & vbCrLf
    s = s & "               ( 'tblScheduleB' )," & vbCrLf
    s = s & "               ( 'ClientDBCOMaster' )," & vbCrLf
    s = s & "               ( 'ClientDBCODetails' )," & vbCrLf
    s = s & "               ('ClientDBTime')" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "               ;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  Name" & vbCrLf
    s = s & "        FROM    @TableNames;" & vbCrLf
    s = s & "       " & vbCrLf
    s = s & "-- System Setup Options" & vbCrLf
    s = s & "       SELECT ISNULL(OptionByArea,0) AS IsOptionByCommunity" & vbCrLf
    s = s & "       ,ISNULL( OptionByAreaPhase        ,0) as IsOptionByCommunityPhase" & vbCrLf
    s = s & "       ,ISNULL( DCOptionByArea           ,0) as IsDCOptionByCommunity" & vbCrLf
    s = s & "       ,ISNULL( DCOptionByAreaPhase      ,0) as IsDCOptionByCommunityPhase" & vbCrLf
    s = s & "       ,ISNULL( GlobalOptionByArea       ,0) as IsGlobalOptionByCommunity" & vbCrLf
    s = s & "       ,ISNULL( GlobalOptionByAreaPhase  ,0) as IsGlobalOptionByCommunityPhase" & vbCrLf
    s = s & "       ,ISNULL( ModelByArea              ,0) as IsModelByCommunity" & vbCrLf
    s = s & "       ,ISNULL(  ModelsbyArea_Phase      ,0) as IsModelByCommunityPhase" & vbCrLf
    s = s & "       ,ISNULL(  ID , 0) as ClientDivisionID" & vbCrLf
    s = s & "       FROM dbo.System_Setup       " & vbCrLf
    s = s & "       " & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- Divisions" & vbCrLf
    s = s & "        SELECT  DivisionID" & vbCrLf
    s = s & "              , DivisionCode" & vbCrLf
    s = s & "              , DivisionName" & vbCrLf
    s = s & "              , d.Address1" & vbCrLf
    s = s & "              , d.Address2" & vbCrLf
    s = s & "              , d.City" & vbCrLf
    s = s & "              , d.Province" & vbCrLf
    s = s & "              , Postal" & vbCrLf
    s = s & "              , d.Country AS Country" & vbCrLf
    s = s & "              , County" & vbCrLf
    s = s & "              , d.Fax" & vbCrLf
    s = s & "              , d.Phone" & vbCrLf
    s = s & "              , s.GST_Rate AS GSTTaxRate" & vbCrLf
    s = s & "              , s.pst_rate AS PSTTaxRate" & vbCrLf
    s = s & "              , s.Show_CDN_GST AS ShowGST" & vbCrLf
    s = s & "              , s.UsePst AS ShowPST" & vbCrLf
    s = s & "              , s.MultiFamily" & vbCrLf
    s = s & "             , g.FValue1 as GSTBaseAmount" & vbCrLf
    s = s & "             , g.FValue5 as GSTMaxAmount" & vbCrLf
    s = s & "        FROM    Divisions d" & vbCrLf
    s = s & "                INNER JOIN System_Setup s ON d.DivisionID = s.ID" & vbCrLf
    s = s & "               LEFT OUTER JOIN tblGSTPSTRebate g on Tax_Type = 'GST' and tax_Province = 'AB' and taxrate = 5" & vbCrLf
    s = s & "               ;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  DivisionID" & vbCrLf
    s = s & "              , Community" & vbCrLf
    s = s & "        FROM    DivisionCommunities;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- Lot Status" & vbCrLf
    s = s & "        SELECT  CASE WHEN [Status] = 'Spec' THEN 'Spec Home' ELSE [Status] END AS [STATUS]" & vbCrLf
    s = s & "        FROM    LotStatus;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- tblModels" & vbCrLf
    s = s & "        SELECT  [Area]" & vbCrLf
    s = s & "              , [CommunityPhase]" & vbCrLf
    s = s & "              , [Model]" & vbCrLf
    s = s & "              , [Style]" & vbCrLf
    s = s & "              , [Series]" & vbCrLf
    s = s & "              , [Elevation]" & vbCrLf
    s = s & "              , [Description]" & vbCrLf
    s = s & "              , [ModelSize]" & vbCrLf
    s = s & "              , [NoOfBedRooms]" & vbCrLf
    s = s & "              , [NoOfBathRooms]" & vbCrLf
    s = s & "              , [Base_House]" & vbCrLf
    s = s & "              , [Cost_Amount]" & vbCrLf
    s = s & "              , [Inactive]" & vbCrLf
    s = s & "              , [max_width]" & vbCrLf
    s = s & "              , [max_length]" & vbCrLf
    s = s & "              , [UseTax]" & vbCrLf
    s = s & "              , [NetTax]" & vbCrLf
    s = s & "              , [TotalAmount]" & vbCrLf
    s = s & "              , [Comments]" & vbCrLf
    s = s & "              , [SalesWorksheet]" & vbCrLf
    s = s & "              , CASE WHEN Spec_Document != ''" & vbCrLf
    s = s & "                     THEN REVERSE(SUBSTRING(REVERSE(Spec_Document), 1, CHARINDEX('\', REVERSE(Spec_Document)) - 1))" & vbCrLf
    s = s & "                     ELSE ''" & vbCrLf
    s = s & "                END Spec_Document" & vbCrLf
    s = s & "              , [assembly]" & vbCrLf
    s = s & "              , [Brochure]" & vbCrLf
    s = s & "              , [IncentiveRetail]" & vbCrLf
    s = s & "              , [IncentiveCost]" & vbCrLf
    s = s & "              , [LastSalesWorksheet]" & vbCrLf
    s = s & "              , Round([Margin],2) as Margin" & vbCrLf
    s = s & "              , Round([Markup],2) as Markup" & vbCrLf
    s = s & "              , [DivisionID]" & vbCrLf
    s = s & "              , [seq] AS ClientSeq" & vbCrLf
    s = s & "        FROM    tblModels" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblModels'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "   -- Series, MasterDate, CommunityPhase" & vbCrLf
    s = s & "        SELECT  Series" & vbCrLf
    s = s & "              , Description" & vbCrLf
    s = s & "             , DivisionID" & vbCrLf
    s = s & "        FROM    tblSeries" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblSeries'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  [Date_Field]" & vbCrLf
    s = s & "              , [Description]" & vbCrLf
    s = s & "              , [Order_Index]" & vbCrLf
    s = s & "              , [notice_reqd]" & vbCrLf
    s = s & "              , [notice_days]" & vbCrLf
    s = s & "              , [Amount]" & vbCrLf
    s = s & "              , [Date_Type]" & vbCrLf
    s = s & "              , [prospectonly]" & vbCrLf
    s = s & "              , [sales_view]" & vbCrLf
    s = s & "              , [vendor_webview]" & vbCrLf
    s = s & "              , [Web_View]" & vbCrLf
    s = s & "        FROM    Master_Date;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       " & vbCrLf
    s = s & "   " & vbCrLf
    s = s & "   --Community Phase" & vbCrLf
    s = s & "       SELECT DISTINCT li.Community" & vbCrLf
    s = s & "           , Phase AS CommunityPhase" & vbCrLf
    s = s & "           , Phase AS [Description] " & vbCrLf
    s = s & "       FROM dbo.tblLotInventory li" & vbCrLf
    s = s & "       INNER JOIN dbo.DivisionCommunities dc ON li.Community = dc.Community" & vbCrLf
    s = s & "       INNER JOIN dbo.System_Setup ss ON dc.DivisionID = ss.ID" & vbCrLf
    s = s & "       WHERE ISNULL(Phase,'') != ''" & vbCrLf
    s = s & "         AND ISNULL(ss.ModelsbyArea_Phase,0) = 0" & vbCrLf
    s = s & "   UNION " & vbCrLf
    s = s & "      SELECT DISTINCT  cp.Community" & vbCrLf
    s = s & "           , CommunityPhase" & vbCrLf
    s = s & "           , Description" & vbCrLf
    s = s & "       FROM    CommunityPhase cp" & vbCrLf
    s = s & "       INNER JOIN dbo.DivisionCommunities dc ON cp.Community = dc.Community" & vbCrLf
    s = s & "       INNER JOIN dbo.System_Setup ss ON dc.DivisionID = ss.ID" & vbCrLf
    s = s & "       WHERE ISNULL(CommunityPhase,'') != ''" & vbCrLf
    s = s & "       AND ISNULL(ss.ModelsbyArea_Phase,0) = 1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "   -- LotType, LotPremium, LotPreMaster, LotInventory" & vbCrLf
    s = s & "        SELECT  ID" & vbCrLf
    s = s & "              , LOT_TYPE" & vbCrLf
    s = s & "        FROM    LOT_TYPE;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  [Lot_No]" & vbCrLf
    s = s & "              , [PremiumID]" & vbCrLf
    s = s & "              , [Description]" & vbCrLf
    s = s & "              , [PremiumValue]" & vbCrLf
    s = s & "              , [PremiumOverride]" & vbCrLf
    s = s & "              , [UseOverride]" & vbCrLf
    s = s & "        FROM    LotPremium;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  PremiumID" & vbCrLf
    s = s & "              , Description" & vbCrLf
    s = s & "        FROM    LotPremiumMaster;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  [Lot_No]" & vbCrLf
    s = s & "              , [Job_No]" & vbCrLf
    s = s & "              , [Lot]" & vbCrLf
    s = s & "              , [Block]" & vbCrLf
    s = s & "              , [Cost]" & vbCrLf
    s = s & "              , [Selling_Price]" & vbCrLf
    s = s & "              , [Municipal_Address]" & vbCrLf
    s = s & "              , [Community]" & vbCrLf
    s = s & "              , [Phase]" & vbCrLf
    s = s & "              , CASE WHEN [Status] = 'Spec' THEN 'Spec Home' ELSE [Status] END AS [STATUS]" & vbCrLf
    s = s & "              , [LOT_TYPE]" & vbCrLf
    s = s & "              , [LegalAddress]" & vbCrLf
    s = s & "              , [LOTPLAN]" & vbCrLf
    s = s & "              , [City]" & vbCrLf
    s = s & "              , [Province]" & vbCrLf
    s = s & "              , [Zip]" & vbCrLf
    s = s & "              , [County]" & vbCrLf
    s = s & "              , [Township]" & vbCrLf
    s = s & "              , [Country]" & vbCrLf
    s = s & "              , [Comments]" & vbCrLf
    s = s & "              , [Garage_Orientation]" & vbCrLf
    s = s & "              , [max_width]" & vbCrLf
    s = s & "              , [max_length]" & vbCrLf
    s = s & "              , [FrontWidth]" & vbCrLf
    s = s & "              , [BackWidth]" & vbCrLf
    s = s & "              , [ActualLength]" & vbCrLf
    s = s & "              , [ActualWidth]" & vbCrLf
    s = s & "              , [LeftSideLength]" & vbCrLf
    s = s & "              , [RightSideLength]" & vbCrLf
    s = s & "              , [ConstructionStatus]" & vbCrLf
    s = s & "        FROM    tblLotInventory" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblLotInventory'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       -- Categories, Communities, MajorGroups" & vbCrLf
    s = s & "        SELECT  Category" & vbCrLf
    s = s & "              , Description" & vbCrLf
    s = s & "              , Group_Code" & vbCrLf
    s = s & "        FROM    tblcategories" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblCategories'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  [Area]" & vbCrLf
    s = s & "              , [Description]" & vbCrLf
    s = s & "              , CASE WHEN Contract_Document != ''" & vbCrLf
    s = s & "                     THEN REVERSE(SUBSTRING(REVERSE(Contract_Document), 1," & vbCrLf
    s = s & "                                            CHARINDEX('\', REVERSE(Contract_Document)) - 1))" & vbCrLf
    s = s & "                     ELSE ''" & vbCrLf
    s = s & "                END Contract_Document" & vbCrLf
    s = s & "              , [CompanyName]" & vbCrLf
    s = s & "              , [Address1]" & vbCrLf
    s = s & "              , [Address2]" & vbCrLf
    s = s & "              , [City]" & vbCrLf
    s = s & "              , [Province]" & vbCrLf
    s = s & "              , [Country]" & vbCrLf
    s = s & "              , [County]" & vbCrLf
    s = s & "              , [Zip]" & vbCrLf
    s = s & "              , [GSTNumber]" & vbCrLf
    s = s & "              , [Phone]" & vbCrLf
    s = s & "              , [FAX]" & vbCrLf
    s = s & "              , [Inactive]" & vbCrLf
    s = s & "              , [UsesPhases]" & vbCrLf
    s = s & "              , [GSTRate]" & vbCrLf
    s = s & "              , [PSTRate]" & vbCrLf
    s = s & "              , [SalesManagerEmail]" & vbCrLf
    s = s & "        FROM    tblLocality" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tbllocality'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  Major_Group" & vbCrLf
    s = s & "              , Description" & vbCrLf
    s = s & "        FROM    tblMajorGroups" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblMajorGroups'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       -- Option & GlobalOptions" & vbCrLf
    s = s & "        SELECT  [Area]" & vbCrLf
    s = s & "              , [CommunityPhase]" & vbCrLf
    s = s & "              , [Model]" & vbCrLf
    s = s & "              , [Elevation]" & vbCrLf
    s = s & "              , [Series]" & vbCrLf
    s = s & "              , [OPT]" & vbCrLf
    s = s & "              , [Description]" & vbCrLf
    s = s & "              , [Construction_Cut_Off]" & vbCrLf
    s = s & "              , [D_Structural_Change]" & vbCrLf
    s = s & "              , [Cost_Amount]" & vbCrLf
    s = s & "              , [Price]" & vbCrLf
    s = s & "              , [CO_PRICE]" & vbCrLf
    s = s & "              , [Category]" & vbCrLf
    s = s & "              , [UOM]" & vbCrLf
    s = s & "              , [NoCommission]" & vbCrLf
    s = s & "              , [qty]" & vbCrLf
    s = s & "              , [Comments]" & vbCrLf
    s = s & "              , [UseTax]" & vbCrLf
    s = s & "              , [NetTax]" & vbCrLf
    s = s & "              , [TotalAmount]" & vbCrLf
    s = s & "              , [NetCOTax]" & vbCrLf
    s = s & "              , [TotalCOAmount]" & vbCrLf
    s = s & "              , [SalesWorksheet]" & vbCrLf
    s = s & "              , 1 as [LastSalesWorksheet]" & vbCrLf
    s = s & "              , [assembly]" & vbCrLf
    s = s & "              , [ApplyIncentive]" & vbCrLf
    s = s & "              , Round([Margin],2) as Margin" & vbCrLf
    s = s & "              , round([Markup],2) as Markup" & vbCrLf
    s = s & "              , [EstimatorNotes]" & vbCrLf
    s = s & "              , [Location]" & vbCrLf
    s = s & "              , [Color]" & vbCrLf
    s = s & "              , [Style]" & vbCrLf
    s = s & "              , [Finish]" & vbCrLf
    s = s & "              , [Other]" & vbCrLf
    s = s & "              , [ColorListID]" & vbCrLf
    s = s & "              , [StyleListID]" & vbCrLf
    s = s & "              , [FinishListID]" & vbCrLf
    s = s & "              , [OtherListID]" & vbCrLf
    s = s & "              , [InActive]" & vbCrLf
    s = s & "              , [IncludedOption]" & vbCrLf
    s = s & "              , [DivisionID]" & vbCrLf
    s = s & "              , SEQ AS clientSeq" & vbCrLf
    s = s & "             , SelectByRoom" & vbCrLf
    s = s & "             , DisplayTotalonly" & vbCrLf
    s = s & "             , DesignCenterSalesOnly" & vbCrLf
    s = s & "        FROM    tblOptions" & vbCrLf
    s = s & "        WHERE   ModifiedDate >=        ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tbloptions'" & vbCrLf
    s = s & "                                )" & vbCrLf
    s = s & "               " & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  [Community]" & vbCrLf
    s = s & "              , [CommunityPhase]" & vbCrLf
    s = s & "              , [OPT]" & vbCrLf
    s = s & "              , [Description]" & vbCrLf
    s = s & "              , [Construction_Cut_Off]" & vbCrLf
    s = s & "              , [D_Structural_Change]" & vbCrLf
    s = s & "              , [Cost_Amount]" & vbCrLf
    s = s & "              , [Price]" & vbCrLf
    s = s & "              , [CO_Price]" & vbCrLf
    s = s & "              , [Category]" & vbCrLf
    s = s & "              , [UOM]" & vbCrLf
    s = s & "              , [NoCommission]" & vbCrLf
    s = s & "              , [Comments]" & vbCrLf
    s = s & "              , [UseTax]" & vbCrLf
    s = s & "              , [NetTax]" & vbCrLf
    s = s & "              , [TotalAmount]" & vbCrLf
    s = s & "              , [NetCOTax]" & vbCrLf
    s = s & "              , [TotalCOAmount]" & vbCrLf
    s = s & "              , [SalesWorksheet]" & vbCrLf
    s = s & "              , 2 as [LastSalesWorksheet]" & vbCrLf
    s = s & "              , [assembly]" & vbCrLf
    s = s & "              , [ApplyIncentive]" & vbCrLf
    s = s & "              , Round([Margin],2) as Margin" & vbCrLf
    s = s & "              , round([Markup],2) as Markup" & vbCrLf
    s = s & "              , [EstimatorNotes]" & vbCrLf
    s = s & "              , [Location]" & vbCrLf
    s = s & "              , [Color]" & vbCrLf
    s = s & "              , [Style]" & vbCrLf
    s = s & "              , [Finish]" & vbCrLf
    s = s & "              , [Other]" & vbCrLf
    s = s & "              , [ColorListID]" & vbCrLf
    s = s & "              , [StyleListID]" & vbCrLf
    s = s & "              , [FinishListID]" & vbCrLf
    s = s & "              , [OtherListID]" & vbCrLf
    s = s & "              , [InActive]" & vbCrLf
    s = s & "              , [IncludedOption]" & vbCrLf
    s = s & "              , DivisionID" & vbCrLf
    s = s & "              , qty" & vbCrLf
    s = s & "              , SEQ AS clientSeq" & vbCrLf
    s = s & "             , SelectByRoom" & vbCrLf
    s = s & "             , DisplayTotalonly" & vbCrLf
    s = s & "             , DesignCenterSalesOnly" & vbCrLf
    s = s & "        FROM    tblGlobalOptions" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblGlobaloptions'" & vbCrLf
    s = s & "                                )" & vbCrLf
    s = s & "       UNION                       " & vbCrLf
    s = s & "       SELECT  [Community]" & vbCrLf
    s = s & "              , [CommunityPhase]" & vbCrLf
    s = s & "              , [OPT]" & vbCrLf
    s = s & "              , [Description]" & vbCrLf
    s = s & "              , 0 as Construction_Cut_Off" & vbCrLf
    s = s & "              , 0 as [D_Structural_Change]" & vbCrLf
    s = s & "              , [Item1_Cost] as Cost_Amount" & vbCrLf
    s = s & "              , Item1 as [Price]" & vbCrLf
    s = s & "              , Item1 as [CO_Price]" & vbCrLf
    s = s & "              , [Category]" & vbCrLf
    s = s & "              , [UOM]" & vbCrLf
    s = s & "              , [NoCommission]" & vbCrLf
    s = s & "              , [Comments]" & vbCrLf
    s = s & "              , [UseTax]" & vbCrLf
    s = s & "              , [NetTax]" & vbCrLf
    s = s & "              , [TotalAmount]" & vbCrLf
    s = s & "              , NetTax1" & vbCrLf
    s = s & "              , Item1 as [TotalCOAmount]" & vbCrLf
    s = s & "              , [SalesWorksheet]" & vbCrLf
    s = s & "              , 3 as [LastSalesWorksheet]" & vbCrLf
    s = s & "              , [assembly]" & vbCrLf
    s = s & "              , [ApplyIncentive]" & vbCrLf
    s = s & "              , [Margin1]" & vbCrLf
    s = s & "              , [Markup1]" & vbCrLf
    s = s & "              , [EstimatorNotes]" & vbCrLf
    s = s & "              , [Location]" & vbCrLf
    s = s & "              , [Color]" & vbCrLf
    s = s & "              , [Style]" & vbCrLf
    s = s & "              , [Finish]" & vbCrLf
    s = s & "              , [Other]" & vbCrLf
    s = s & "              , [ColorListID]" & vbCrLf
    s = s & "              , [StyleListID]" & vbCrLf
    s = s & "              , [FinishListID]" & vbCrLf
    s = s & "              , [OtherListID]" & vbCrLf
    s = s & "              , [InActive]" & vbCrLf
    s = s & "              , 0 as [IncludedOption]" & vbCrLf
    s = s & "              , DivisionID" & vbCrLf
    s = s & "              , qty" & vbCrLf
    s = s & "              , SEQ AS clientSeq" & vbCrLf
    s = s & "             , SelectByRoom" & vbCrLf
    s = s & "             , DisplayTotalonly" & vbCrLf
    s = s & "             , DesignCenterSalesOnly" & vbCrLf
    s = s & "        FROM    tblDCOptions" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblGlobaloptions'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       -- Users and User Permissions, User Division" & vbCrLf
    s = s & "        SELECT  u.[User_Id] AS UserID" & vbCrLf
    s = s & "              , u.[User_Name] AS UserName" & vbCrLf
    s = s & "              , u.Designation AS FirstName" & vbCrLf
    s = s & "              , u.Email AS Email" & vbCrLf
    s = s & "              , u.User_Password AS HashPassword" & vbCrLf
    s = s & "              , '' AS HashPasswordSalt" & vbCrLf
    s = s & "              , 0 AS IsInactive" & vbCrLf
    s = s & "              , g.User_Access AS UserRole" & vbCrLf
    s = s & "              , ISNULL(u.AccessMenu, 0) AccessMenu" & vbCrLf
    s = s & "              , ISNULL(u.AddAddendum, 0) AddAddendum" & vbCrLf
    s = s & "              , ISNULL(u.AddCO, 0) AddCO" & vbCrLf
    s = s & "              , ISNULL(u.AddCustomerDates, 0) AddCustomerDates" & vbCrLf
    s = s & "              , ISNULL(u.AddDC, 0) AddDC" & vbCrLf
    s = s & "              , ISNULL(u.AllowAddCO_LOCKCO, 0) AllowAddCO_LOCKCO" & vbCrLf
    s = s & "              , ISNULL(u.ALLOWDELCO, 0) ALLOWDELCO" & vbCrLf
    s = s & "              , ISNULL(u.ApproveCustomer, 0) ApproveCustomer" & vbCrLf
    s = s & "              , ISNULL(u.BarcodeEntry, 0) BarcodeEntry" & vbCrLf
    s = s & "              , ISNULL(u.BldrApproveCustomer, 0) BldrApproveCustomer" & vbCrLf
    s = s & "              , ISNULL(u.ChangeDate, 0) ChangeDate" & vbCrLf
    s = s & "              , ISNULL(u.ChangeGridLayout, 0) ChangeGridLayout" & vbCrLf
    s = s & "              , ISNULL(u.ChangeMiscPurchased, 0) ChangeMiscPurchased" & vbCrLf
    s = s & "              , ISNULL(u.CopyCustomer, 0) CopyCustomer" & vbCrLf
    s = s & "              , ISNULL(u.CreateJob, 0) CreateJob" & vbCrLf
    s = s & "              , ISNULL(u.CreateNewCustomer, 0) CreateNewCustomer" & vbCrLf
    s = s & "              , ISNULL(u.CreateSchedule, 0) CreateSchedule" & vbCrLf
    s = s & "              , ISNULL(u.CustomPricing, 0) CustomPricing" & vbCrLf
    s = s & "              , ISNULL(u.DeleteCustomer, 0) DeleteCustomer" & vbCrLf
    s = s & "              , ISNULL(u.DeleteSchedule, 0) DeleteSchedule" & vbCrLf
    s = s & "              , ISNULL(u.EditActualEndDate, 0) EditActualEndDate" & vbCrLf
    s = s & "              , ISNULL(u.EditAssemblies, 0) EditAssemblies" & vbCrLf
    s = s & "              , ISNULL(u.EditAttachments, 0) EditAttachments" & vbCrLf
    s = s & "              , ISNULL(u.EditCompletedCustom, 0) EditCompletedCustom" & vbCrLf
    s = s & "              , ISNULL(u.EditItemdb, 0) EditItemdb" & vbCrLf
    s = s & "              , ISNULL(u.FileOpenData, 0) FileOpenData" & vbCrLf
    s = s & "              , ISNULL(u.GeneratePOs, 0) GeneratePOs" & vbCrLf
    s = s & "              , ISNULL(u.InqDCOptions, 0) InqDCOptions" & vbCrLf
    s = s & "              , ISNULL(u.InqGlobalOptions, 0) InqGlobalOptions" & vbCrLf
    s = s & "              , ISNULL(u.InqLockedCustomer, 0) InqLockedCustomer" & vbCrLf
    s = s & "              , ISNULL(u.InqModelOptions, 0) InqModelOptions" & vbCrLf
    s = s & "              , ISNULL(u.InqNewDeposits, 0) InqNewDeposits" & vbCrLf
    s = s & "              , ISNULL(u.InqNewMortgage, 0) InqNewMortgage" & vbCrLf
    s = s & "              , ISNULL(u.InqPostDeposits, 0) InqPostDeposits" & vbCrLf
    s = s & "              , ISNULL(u.InqPostMortgage, 0) InqPostMortgage" & vbCrLf
    s = s & "              , ISNULL(u.InquiryMenu, 0) InquiryMenu" & vbCrLf
    s = s & "              , ISNULL(u.IssueBudgets, 0) IssueBudgets" & vbCrLf
    s = s & "              , ISNULL(u.IssuePOs, 0) IssuePOs" & vbCrLf
    s = s & "              , ISNULL(u.LimitAttachmentClasses, 0) LimitAttachmentClasses" & vbCrLf
    s = s & "              , ISNULL(u.ModifyCustom, 0) ModifyCustom" & vbCrLf
    s = s & "              , ISNULL(u.OpenEstimatingWorksheets, 0) OpenEstimatingWorksheets" & vbCrLf
    s = s & "              , ISNULL(u.OpenMarketingWorksheets, 0) OpenMarketingWorksheets" & vbCrLf
    s = s & "              , ISNULL(u.ParkingStallMenu, 0) ParkingStallMenu" & vbCrLf
    s = s & "              , ISNULL(u.PendingCustomerPick, 0) PendingCustomerPick" & vbCrLf
    s = s & "              , ISNULL(u.PostBudgets, 0) PostBudgets" & vbCrLf
    s = s & "              , ISNULL(u.PostCommitments, 0) PostCommitments" & vbCrLf
    s = s & "              , ISNULL(u.PostToTimberline, 0) PostToTimberline" & vbCrLf
    s = s & "              , ISNULL(u.PrecisionBldrOptions, 0) PrecisionBldrOptions" & vbCrLf
    s = s & "              , ISNULL(u.PromptforCOApproval, 0) PromptforCOApproval" & vbCrLf
    s = s & "              , ISNULL(u.PublishWorksheets, 0) PublishWorksheets" & vbCrLf
    s = s & "              , ISNULL(u.PurchaseCustomer, 0) PurchaseCustomer" & vbCrLf
    s = s & "              , ISNULL(u.ReportsMenu, 0) ReportsMenu" & vbCrLf
    s = s & "              , ISNULL(u.ReportsRepMan, 0) ReportsRepMan" & vbCrLf
    s = s & "              , ISNULL(u.ReverseCustDateTab, 0) ReverseCustDateTab" & vbCrLf
    s = s & "              , ISNULL(u.SalesOverride, 0) SalesOverride" & vbCrLf
    s = s & "              , ISNULL(u.SendPO, 0) SendPO" & vbCrLf
    s = s & "              , ISNULL(u.SetupCategory, 0) SetupCategory" & vbCrLf
    s = s & "              , ISNULL(u.SetupCommunity, 0) SetupCommunity" & vbCrLf
    s = s & "              , ISNULL(u.SetupConstStatus, 0) SetupConstStatus" & vbCrLf
    s = s & "              , ISNULL(u.SetupCustomer, 0) SetupCustomer" & vbCrLf
    s = s & "              , ISNULL(u.SetupCustSrvPerson, 0) SetupCustSrvPerson" & vbCrLf
    s = s & "              , ISNULL(u.SetupDCArea, 0) SetupDCArea" & vbCrLf
    s = s & "              , ISNULL(u.SetupDCOption, 0) SetupDCOption" & vbCrLf
    s = s & "              , ISNULL(u.SetupDefaultVendor, 0) SetupDefaultVendor" & vbCrLf
    s = s & "              , ISNULL(u.SetupDevelopers, 0) SetupDevelopers" & vbCrLf
    s = s & "              , ISNULL(u.SetupGlobalOptions, 0) SetupGlobalOptions" & vbCrLf
    s = s & "              , ISNULL(u.SetupGroupTemplate, 0) SetupGroupTemplate" & vbCrLf
    s = s & "              , ISNULL(u.SetupJobs, 0) SetupJobs" & vbCrLf
    s = s & "              , ISNULL(u.SetupLawFirms, 0) SetupLawFirms" & vbCrLf
    s = s & "              , ISNULL(u.SetupLendingCompany, 0) SetupLendingCompany" & vbCrLf
    s = s & "              , ISNULL(u.SetupLotInventory, 0) SetupLotInventory" & vbCrLf
    s = s & "              , ISNULL(u.SetupMajorGroup, 0) SetupMajorGroup" & vbCrLf
    s = s & "              , ISNULL(u.SetupMenu, 0) SetupMenu" & vbCrLf
    s = s & "              , ISNULL(u.SetupModel, 0) SetupModel" & vbCrLf
    s = s & "              , ISNULL(u.SetupModelOptions, 0) SetupModelOptions" & vbCrLf
    s = s & "              , ISNULL(u.SetupPOIndex, 0) SetupPOIndex" & vbCrLf
    s = s & "              , ISNULL(u.SetupProjectManager, 0) SetupProjectManager" & vbCrLf
    s = s & "              , ISNULL(u.SetupSalesPerson, 0) SetupSalesPerson" & vbCrLf
    s = s & "              , ISNULL(u.SetupSchedulingData, 0) SetupSchedulingData" & vbCrLf
    s = s & "              , ISNULL(u.SetupSeries, 0) SetupSeries" & vbCrLf
    s = s & "              , ISNULL(u.SetupTLDateList, 0) SetupTLDateList" & vbCrLf
    s = s & "              , ISNULL(u.SetupVendorPricing, 0) SetupVendorPricing" & vbCrLf
    s = s & "              , ISNULL(u.SetupVendors, 0) SetupVendors" & vbCrLf
    s = s & "              , ISNULL(u.ShowAccountingInfo, 0) ShowAccountingInfo" & vbCrLf
    s = s & "              , ISNULL(u.ShowAddendum, 0) ShowAddendum" & vbCrLf
    s = s & "              , ISNULL(u.ShowBuyerInfo, 0) ShowBuyerInfo" & vbCrLf
    s = s & "              , ISNULL(u.ShowCO, 0) ShowCO" & vbCrLf
    s = s & "              , ISNULL(u.ShowContractSummary, 0) ShowContractSummary" & vbCrLf
    s = s & "              , ISNULL(u.ShowDates, 0) ShowDates" & vbCrLf
    s = s & "              , ISNULL(u.ShowDC, 0) ShowDC" & vbCrLf
    s = s & "              , ISNULL(u.ShowDeletedItems, 0) ShowDeletedItems" & vbCrLf
    s = s & "              , ISNULL(u.ShowHomeInfo, 0) ShowHomeInfo" & vbCrLf
    s = s & "              , ISNULL(u.ShowLotPremium, 0) ShowLotPremium" & vbCrLf
    s = s & "              , ISNULL(u.ShowSelections, 0) ShowSelections" & vbCrLf
    s = s & "              , ISNULL(u.TaskEntDep, 0) TaskEntDep" & vbCrLf
    s = s & "              , ISNULL(u.TaskEntMort, 0) TaskEntMort" & vbCrLf
    s = s & "              , ISNULL(u.TaskEntQuote, 0) TaskEntQuote" & vbCrLf
    s = s & "              , ISNULL(u.TaskImportTLCustom, 0) TaskImportTLCustom" & vbCrLf
    s = s & "              , ISNULL(u.TaskPostAdj, 0) TaskPostAdj" & vbCrLf
    s = s & "              , ISNULL(u.TaskPostDep, 0) TaskPostDep" & vbCrLf
    s = s & "              , ISNULL(u.TaskPostMort, 0) TaskPostMort" & vbCrLf
    s = s & "              , ISNULL(u.TasksMenu, 0) TasksMenu" & vbCrLf
    s = s & "              , ISNULL(u.TaskSyncDate, 0) TaskSyncDate" & vbCrLf
    s = s & "              , ISNULL(u.TaskSyncGroup, 0) TaskSyncGroup" & vbCrLf
    s = s & "              , ISNULL(u.TaskUpdAddrJobCost, 0) TaskUpdAddrJobCost" & vbCrLf
    s = s & "              , ISNULL(u.TaskUpdDatesTL, 0) TaskUpdDatesTL" & vbCrLf
    s = s & "              , ISNULL(u.TaskUpdTotalJobCost, 0) TaskUpdTotalJobCost" & vbCrLf
    s = s & "              , ISNULL(u.ToolsAdminMaintain, 0) ToolsAdminMaintain" & vbCrLf
    s = s & "              , ISNULL(u.ToolsCommSettings, 0) ToolsCommSettings" & vbCrLf
    s = s & "              , ISNULL(u.ToolsCustDesc, 0) ToolsCustDesc" & vbCrLf
    s = s & "              , ISNULL(u.ToolsCustMaintain, 0) ToolsCustMaintain" & vbCrLf
    s = s & "              , ISNULL(u.ToolsCustRepPackage, 0) ToolsCustRepPackage" & vbCrLf
    s = s & "              , ISNULL(u.ToolsFieldVal, 0) ToolsFieldVal" & vbCrLf
    s = s & "              , ISNULL(u.ToolsGSTPSTRebate, 0) ToolsGSTPSTRebate" & vbCrLf
    s = s & "              , ISNULL(u.ToolsLotStatus, 0) ToolsLotStatus" & vbCrLf
    s = s & "              , ISNULL(u.ToolsMenu, 0) ToolsMenu" & vbCrLf
    s = s & "              , ISNULL(u.ToolsSystemSetup, 0) ToolsSystemSetup" & vbCrLf
    s = s & "              , ISNULL(u.ToolsTLFieldDesc, 0) ToolsTLFieldDesc" & vbCrLf
    s = s & "              , ISNULL(u.ToolsUsrAdmin, 0) ToolsUsrAdmin" & vbCrLf
    s = s & "              , ISNULL(u.ToolsWorkInProgress, 0) ToolsWorkInProgress" & vbCrLf
    s = s & "              , ISNULL(u.UnApproveCO, 0) UnApproveCO" & vbCrLf
    s = s & "              , ISNULL(u.UnPurchase, 0) UnPurchase" & vbCrLf
    s = s & "              , ISNULL(u.UPDAddendum, 0) UPDAddendum" & vbCrLf
    s = s & "              , ISNULL(u.UpdateConstStatus, 0) UpdateConstStatus" & vbCrLf
    s = s & "              , ISNULL(u.UpdateCostInfo, 0) UpdateCostInfo" & vbCrLf
    s = s & "              , ISNULL(u.UPDCO, 0) UPDCO" & vbCrLf
    s = s & "              , ISNULL(u.UPDContractSummary, 0) UPDContractSummary" & vbCrLf
    s = s & "              , ISNULL(u.UPDCustomer, 0) UPDCustomer" & vbCrLf
    s = s & "              , ISNULL(u.UPDCustomerDates, 0) UPDCustomerDates" & vbCrLf
    s = s & "              , ISNULL(u.UPDCustomerGroups, 0) UPDCustomerGroups" & vbCrLf
    s = s & "              , ISNULL(u.UPDDC, 0) UPDDC" & vbCrLf
    s = s & "              , ISNULL(u.UPDDeletedItems, 0) UPDDeletedItems" & vbCrLf
    s = s & "              , ISNULL(u.UPDFromTL, 0) UPDFromTL" & vbCrLf
    s = s & "              , ISNULL(u.UPDSpecShow, 0) UPDSpecShow" & vbCrLf
    s = s & "              , ISNULL(u.UPDTLJob, 0) UPDTLJob" & vbCrLf
    s = s & "              , ISNULL(u.UPDToTL, 0) UPDToTL" & vbCrLf
    s = s & "              , ISNULL(u.ViewAllSchedules, 0) ViewAllSchedules" & vbCrLf
    s = s & "              , ISNULL(u.WTCustomerApprove, 0) WTCustomerApprove" & vbCrLf
    s = s & "              , ISNULL(u.WTEditGridLayouts, 0) WTEditGridLayouts" & vbCrLf
    s = s & "              , ISNULL(u.WTFieldApprove, 0) WTFieldApprove" & vbCrLf
    s = s & "              , ISNULL(u.WTOfficeApprove, 0) WTOfficeApprove" & vbCrLf
    s = s & "        FROM    dbo.User_Manager u" & vbCrLf
    s = s & "       join SecurityGroups g on u.SecGroupID = g.SecGroupID;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  du.USERID AS UserID" & vbCrLf
    s = s & "              , d.DivisionID AS DivisionID" & vbCrLf
    s = s & "        FROM    Divisions d" & vbCrLf
    s = s & "                INNER JOIN DivisionUsers du ON du.DivisionID = d.DivisionID; " & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       SELECT   A.Sales_Person_ID as [UserID]" & vbCrLf
    s = s & "               ,A.Area  " & vbCrLf
    s = s & "        FROM    tblSalesPersonArea A" & vbCrLf
    s = s & "                INNER JOIN User_Manager U on u.[User_Id] = A.Sales_Person_ID " & vbCrLf
    s = s & "               INNER JOIN tblLocality L on L.Area = A.Area AND L.Inactive = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "               -- Inventory Homes" & vbCrLf
    s = s & "        SELECT  i.Description" & vbCrLf
    s = s & "              , i.Home_Selection as ClientPurchaseType " & vbCrLf
    s = s & "             --'Spec Home' AS ClientPurchaseType" & vbCrLf
    s = s & "              , i.Job_No Job" & vbCrLf
    s = s & "             , i.Home_Selection as ClientLotStatus " & vbCrLf
    s = s & "             -- , 'Spec Home' AS ClientLotStatus" & vbCrLf
    s = s & "              , i.ReservedExpiry" & vbCrLf
    s = s & "              , i.Customer_No AS CustomerNo" & vbCrLf
    s = s & "              , Community AS CommunityClientCode" & vbCrLf
    s = s & "              , ( CASE WHEN ISNULL(i.[Floor], '') = '' THEN '0'" & vbCrLf
    s = s & "                       ELSE i.[Floor]" & vbCrLf
    s = s & "                  END ) AS floor" & vbCrLf
    s = s & "              , i.Unit_No AS UnitNo" & vbCrLf
    s = s & "              , i.Legal_Unit_No AS LegalUnit" & vbCrLf
    s = s & "              , i.DirectionID AS ClientDirectionFacingID" & vbCrLf
    s = s & "              , i.UnitFactor" & vbCrLf
    s = s & "              , i.CondoFees" & vbCrLf
    s = s & "              , i.Model AS ModelClientCode" & vbCrLf
    s = s & "              , 0.00 BalconySize" & vbCrLf
    s = s & "              , i.Override_Price ModelPrice" & vbCrLf
    s = s & "              , 0.00 ModelCost" & vbCrLf
    s = s & "              , 0.00 ModelTax" & vbCrLf
    s = s & "              , i.ScheduleB_total AddendumTotal" & vbCrLf
    s = s & "              , 0.00 OptionsTax" & vbCrLf
    s = s & "              , i.GST_Rate" & vbCrLf
    s = s & "              , i.GST" & vbCrLf
    s = s & "              , i.GST_Rebate" & vbCrLf
    s = s & "              , ISNULL(i.RevenueAdjustment,0.00) AS RevenueAdj" & vbCrLf
    s = s & "              , 0.00 RevenueAdj2" & vbCrLf
    s = s & "              , i.Lot" & vbCrLf
    s = s & "              , i.Lot_No AS LotNo" & vbCrLf
    s = s & "              , ISNULL(i.Lot_Price,0.00) AS LotPrice" & vbCrLf
    s = s & "              , i.Lot_Cost AS LotCost" & vbCrLf
    s = s & "              , 0.00 LotTax" & vbCrLf
    s = s & "              , ISNULL(i.LotPremiumValue,0.00) AS PremiumAmount" & vbCrLf
    s = s & "              , ISNULL(i.LotPremiumPreTax,0.00) AS PremiumPreTax" & vbCrLf
    s = s & "              , i.Parking1" & vbCrLf
    s = s & "              , i.Parking1_Cost" & vbCrLf
    s = s & "              , i.Parking1_Price" & vbCrLf
    s = s & "              , i.Parking2" & vbCrLf
    s = s & "              , i.Parking2_Cost" & vbCrLf
    s = s & "              , i.Parking2_Price" & vbCrLf
    s = s & "              , i.Parking3" & vbCrLf
    s = s & "              , i.Parking3_Cost" & vbCrLf
    s = s & "              , i.Parking3_Price" & vbCrLf
    s = s & "              , i.Parking4" & vbCrLf
    s = s & "              , i.Parking4_Cost" & vbCrLf
    s = s & "              , i.Parking4_Price" & vbCrLf
    s = s & "              , i.Parking5" & vbCrLf
    s = s & "              , i.Parking5_Cost" & vbCrLf
    s = s & "              , i.Parking5_Price" & vbCrLf
    s = s & "              , i.Parking6" & vbCrLf
    s = s & "              , i.Parking6_Cost" & vbCrLf
    s = s & "              , i.Parking6_Price" & vbCrLf
    s = s & "              , i.Parking7" & vbCrLf
    s = s & "              , i.Parking7_Cost" & vbCrLf
    s = s & "              , i.Parking7_Price" & vbCrLf
    s = s & "              , i.PST_Rate" & vbCrLf
    s = s & "              , i.PST" & vbCrLf
    s = s & "              , i.PSTRebate" & vbCrLf
    s = s & "              , i.Comments" & vbCrLf
    s = s & "              , i.ModelAssembly" & vbCrLf
    s = s & "              , i.ModelSalesWorksheet" & vbCrLf
    s = s & "              , i.ModelSpecDoc" & vbCrLf
    s = s & "              , i.ModelSpecDate" & vbCrLf
    s = s & "              , ISNULL(i.CO_PD_Total, 0.00) + ISNULL(i.CO_AP_Total, 0.00) ChangeOrderTotal" & vbCrLf
    s = s & "              , 0 IsOffersPending" & vbCrLf
    s = s & "              , CASE WHEN ISNULL(i.NotAvailableforSale, 0) = 0 THEN 1" & vbCrLf
    s = s & "                     ELSE 0" & vbCrLf
    s = s & "                END AS IsAvailableForSale" & vbCrLf
    s = s & "              , NULL InventoryHomeSelected" & vbCrLf
    s = s & "              , 0 IsInventoryHomeSelected" & vbCrLf
    s = s & "              , i.Sales_Person_ID AS UID" & vbCrLf
    s = s & "              , i.PM AS SiteSupervisor" & vbCrLf
    s = s & "              , 1 AS IsSpecHome" & vbCrLf
    s = s & "              , i.Phase AS PhaseClientCode" & vbCrLf
    s = s & "              , i.Series AS SeriesClientCode" & vbCrLf
    s = s & "              , i.Square_Footage AS SquareF" & vbCrLf
    s = s & "              , i.Expected_Occupancy AS CompletionDate" & vbCrLf
    s = s & "              , i.Construction_Status" & vbCrLf
    s = s & "              , i.Total_Sales_Price AS TotalPrice" & vbCrLf
    s = s & "              , i.LastEstimateIndex" & vbCrLf
    s = s & "              , i.EstimateIndex" & vbCrLf
    s = s & "              , i.DivisionID AS ClientDivisionID" & vbCrLf
    s = s & "             , CASE WHEN s.ModelByArea = 1 THEN i.Community" & vbCrLf
    s = s & "                     ELSE ''" & vbCrLf
    s = s & "                END ModelCommunity" & vbCrLf
    s = s & "              , CASE WHEN s.ModelsbyArea_Phase = 1 THEN i.Phase" & vbCrLf
    s = s & "                     ELSE ''" & vbCrLf
    s = s & "                END ModelCommunityPhase" & vbCrLf
    s = s & "             , i.Cancelled " & vbCrLf
    s = s & "           " & vbCrLf
    s = s & "             , i.ModelIncentive" & vbCrLf
    s = s & "             , i.Sales_Initiative As SalesIncentive" & vbCrLf
    s = s & "             , i.Sales_Deduction AS SalesDeduction" & vbCrLf
    s = s & "             , i.TotalIncentives" & vbCrLf
    s = s & "             , i.Elevation" & vbCrLf
    s = s & "             , CRMID EntityID" & vbCrLf
    s = s & "             , i.ModifiedBy" & vbCrLf
    s = s & "             , i.Municipal_Address" & vbCrLf
    s = s & "        FROM    tblCustomers i" & vbCrLf
    s = s & "                JOIN System_Setup s ON i.DivisionID = s.ID" & vbCrLf
    s = s & "        WHERE    ( isnull(i.PreSale_Selection,'PreSale') = 'PreSale'" & vbCrLf
    s = s & "                     AND isnull(i.PreSale_Selection,'PreSale') != 'Lot Only'" & vbCrLf
    s = s & "                     AND ( i.Home_Selection = 'Spec Home' OR i.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "                 )" & vbCrLf
    s = s & "                 AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "                 AND i.Sold =  0" & vbCrLf
    s = s & "                 AND Inactive = 0" & vbCrLf
    s = s & "                 --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "                 --AND i.CRMID IS NULL" & vbCrLf
    s = s & "                 AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "                 AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "                 AND ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                       FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                       WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                    );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- Inventory Homes Addendum" & vbCrLf
    s = s & "        SELECT  sb.Customer_No AS CustomerNo" & vbCrLf
    s = s & "              , OPT AS ClientOption" & vbCrLf
    s = s & "              , Option_Type AS OptionType" & vbCrLf
    s = s & "              , Custom AS IsCustom" & vbCrLf
    s = s & "              , QTY" & vbCrLf
    s = s & "              , UOM" & vbCrLf
    s = s & "              , Base_Price AS [BasePrice]" & vbCrLf
    s = s & "              , Rate" & vbCrLf
    s = s & "              , TAX" & vbCrLf
    s = s & "              , sb.Cost_Amount AS [CostRate]" & vbCrLf
    s = s & "              , sb.Override_Price AS [PreTaxTotal]" & vbCrLf
    s = s & "              , TotalAmount AS TaxInTotal" & vbCrLf
    s = s & "              , sb.GrandTotal" & vbCrLf
    s = s & "              , sb.[Description]" & vbCrLf
    s = s & "              , sb.comments" & vbCrLf
    s = s & "              , Major_Group AS ClientCategory" & vbCrLf
    s = s & "              , Category AS ClientSubCategory" & vbCrLf
    s = s & "              , Assembly" & vbCrLf
    s = s & "              , Color AS ClientColor" & vbCrLf
    s = s & "              , Style" & vbCrLf
    s = s & "              , Finish" & vbCrLf
    s = s & "              , SalesWorksheet" & vbCrLf
    s = s & "              , ApplyIncentive AS IsApplyIncentive" & vbCrLf
    s = s & "              , IncentiveValue" & vbCrLf
    s = s & "              , Location AS ClientLocation" & vbCrLf
    s = s & "              , sb.EstimateIndex" & vbCrLf
    s = s & "              , Other" & vbCrLf
    s = s & "              , sb.seq ClientSequence" & vbCrLf
    s = s & "              , sb.Override_Price OverridePrice" & vbCrLf
    s = s & "             , sb.CRMID EntityID" & vbCrLf
    s = s & "             , c.DivisionID AS ClientDivisionID" & vbCrLf
    s = s & "             , c.CRMID InventoryHomeCRMID" & vbCrLf
    s = s & "             , sb.ModifiedBy" & vbCrLf
    s = s & "             , CONVERT(DATETIME, CONVERT(CHAR(19), sb.ModifiedDate, 126)) AS ModifiedDate" & vbCrLf
    s = s & "             , sb.Require_Quote AS IsRequireQuote" & vbCrLf
    s = s & "        FROM    [dbo].[tblScheduleB] sb" & vbCrLf
    s = s & "                INNER JOIN tblCustomers c ON sb.Customer_No = c.Customer_No" & vbCrLf
    s = s & "        WHERE    ( c.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "               AND c.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "               AND c.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "               AND ( c.Home_Selection = 'Spec Home' OR c.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "           )" & vbCrLf
    s = s & "           AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "           AND c.Sold =  0" & vbCrLf
    s = s & "           AND Inactive = 0" & vbCrLf
    s = s & "           --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "           --AND c.CRMID IS NULL" & vbCrLf
    s = s & "           AND ISNULL(c.Model, '') != ''" & vbCrLf
    s = s & "           AND ISNULL(c.Job_No, '') != ''" & vbCrLf
    s = s & "                AND sb.ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                      FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                      WHERE     lu.TableName = 'tblInventoryHomesAddendum'" & vbCrLf
    s = s & "                                    );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & " -- Quotes and Sales Contract" & vbCrLf
    s = s & "        SELECT  CRMID [EntityID]" & vbCrLf
    s = s & "              , [Community]" & vbCrLf
    s = s & "              , DivisionID [ClientDivisionID]" & vbCrLf
    s = s & "              , Sales_Person_ID [SalesPerson]" & vbCrLf
    s = s & "              , [Description]" & vbCrLf
    s = s & "              , Job_No [JobNo]" & vbCrLf
    s = s & "              , Expected_Occupancy [ExpectedOccupancy]" & vbCrLf
    s = s & "              , Lot_No [LotNo]" & vbCrLf
    s = s & "              , [Lot]" & vbCrLf
    s = s & "              , [Block]" & vbCrLf
    s = s & "              , [Phase]" & vbCrLf
    s = s & "              , LOTPLAN [Plan]" & vbCrLf
    s = s & "              , Approved_By [ApprovedBy]" & vbCrLf
    s = s & "              , Approved_Date [ApprovedDate]" & vbCrLf
    s = s & "              , Contract_Assigned [IsContractSigned]" & vbCrLf
    s = s & "              , Contract_Signed_Date [ContractSignedDate]" & vbCrLf
    s = s & "              , Pos_Date [PossessionDate]" & vbCrLf
    s = s & "              , Construction_Status [ConstructionStatus]" & vbCrLf
    s = s & "              , [Comments]" & vbCrLf
    s = s & "              , Municipal_Address [MunicipalAddress]" & vbCrLf
    s = s & "              , [County]" & vbCrLf
    s = s & "              , [Township]" & vbCrLf
    s = s & "              , [LegalAddress]" & vbCrLf
    s = s & "              , Square_Footage [SquareFootage]" & vbCrLf
    s = s & "              , ChkGST [GSTOverride]" & vbCrLf
    s = s & "              , ChkGstRebate [GSTRebateOverride]" & vbCrLf
    s = s & "              , Last_CO [LastCO]" & vbCrLf
    s = s & "              , Last_Add_Chng [LastAddChng]" & vbCrLf
    s = s & "              , Cancelled [IsCancelled]" & vbCrLf
    s = s & "              , Home_Selection [HomeSelection]" & vbCrLf
    s = s & "              , PreSale_Selection [PurchaseType]" & vbCrLf
    s = s & "              , Sold [IsSold]" & vbCrLf
    s = s & "              , Last_DC [LastDC]" & vbCrLf
    s = s & "              , [IntUID]" & vbCrLf
    s = s & "              , [IntPwd]" & vbCrLf
    s = s & "              , [Referral]" & vbCrLf
    s = s & "              , Approved [Ratified]" & vbCrLf
    s = s & "              , [Purchased]" & vbCrLf
    s = s & "              , Purchased_Date [PurchasedDate]" & vbCrLf
    s = s & "              , Unit_No [UnitNo]" & vbCrLf
    s = s & "              , Legal_Unit_No [LegalUnit]" & vbCrLf
    s = s & "              , [Floor]" & vbCrLf
    s = s & "              , [LotNetTAX]" & vbCrLf
    s = s & "              , [ModelAssembly]" & vbCrLf
    s = s & "              , [ModelSalesWorksheet]" & vbCrLf
    s = s & "              , [ModelSpecDoc]" & vbCrLf
    s = s & "              , [ModelSpecDate]" & vbCrLf
    s = s & "              , [RealtorID]" & vbCrLf
    s = s & "              , [SellingAgent]" & vbCrLf
    s = s & "              , [CancelledDate]" & vbCrLf
    s = s & "              , CONVERT(DATETIME, CONVERT(CHAR(19), [CreatedDate], 126)) [CreatedDate]" & vbCrLf
    s = s & "              , CONVERT(DATETIME, CONVERT(CHAR(19), [ModifiedDate], 126)) [ModifiedDate]" & vbCrLf
    s = s & "              , [PST]" & vbCrLf
    s = s & "              , [PSTRebate]" & vbCrLf
    s = s & "              , [GST]" & vbCrLf
    s = s & "              , GST_Rebate [GSTRebate]" & vbCrLf
    s = s & "              , GST_Rate [GSTRate]" & vbCrLf
    s = s & "              , PST_Rate [PSTRate]" & vbCrLf
    s = s & "              , Base_House [BaseHome]" & vbCrLf
    s = s & "              , Override_Price [BaseHomeSalesPrice]" & vbCrLf
    s = s & "              , ISNULL(Lot_Price,0.00) [LotPrice]" & vbCrLf
    s = s & "              , ISNULL([LotPremiumValue],0.00) AS LotPremiumValue" & vbCrLf
    s = s & "              , Options [ContractAddendumTotal]" & vbCrLf
    s = s & "              , CRMType" & vbCrLf
    s = s & "              , ISNULL(Closed, 0) Closed" & vbCrLf
    s = s & "             " & vbCrLf
    s = s & "             , ModelIncentive" & vbCrLf
    s = s & "             , Sales_Initiative AS SalesIncentive" & vbCrLf
    s = s & "             , Sales_Deduction AS SalesDeduction" & vbCrLf
    s = s & "             , TotalIncentives" & vbCrLf
    s = s & "             , ModifiedBy" & vbCrLf
    s = s & "        FROM    tblCustomers" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblQuotesAndSalesContract'" & vbCrLf
    s = s & "                                )" & vbCrLf
    s = s & "                AND CRMID IS NOT NULL" & vbCrLf
    s = s & "                AND CRMType IS NOT NULL;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "--ScheduleB" & vbCrLf
    s = s & "        SELECT  s.CRMID [EntityDetailID]" & vbCrLf
    s = s & "              , s.OPT [Option]" & vbCrLf
    s = s & "              , s.Option_Type [OptionType]" & vbCrLf
    s = s & "              , s.Custom [IsCustom]" & vbCrLf
    s = s & "              , s.Require_Quote [RequireQuote]" & vbCrLf
    s = s & "              , s.[Description]" & vbCrLf
    s = s & "              , s.[QTY]" & vbCrLf
    s = s & "              , s.Cost_Amount [CostAmount]" & vbCrLf
    s = s & "              , s.Rate [Rate]" & vbCrLf
    s = s & "              , s.GrandTotal [TotalPrice]" & vbCrLf
    s = s & "              , s.[comments]" & vbCrLf
    s = s & "              , s.[UOM]" & vbCrLf
    s = s & "              , s.Major_Group [CategoryName]" & vbCrLf
    s = s & "              , s.Category [SubCategoryName]" & vbCrLf
    s = s & "              , s.[Assembly]" & vbCrLf
    s = s & "              , s.[SalesWorksheet]" & vbCrLf
    s = s & "              , s.TAX [TaxAmount]" & vbCrLf
    s = s & "              , s.AddFromSpec [AddFromSpec]" & vbCrLf
    s = s & "              , s.SpecDetailReference [SpecDetailReference]" & vbCrLf
    s = s & "              , s.EstimatorNotes [EstimatorNotes]" & vbCrLf
    s = s & "              , s.CONumber [CONumber]" & vbCrLf
    s = s & "              , s.Bldr_Declined [IsBldrDeclined]" & vbCrLf
    s = s & "              , s.Bldr_Declined_Reason [BldrDeclinedReason]" & vbCrLf
    s = s & "              , s.[Color]" & vbCrLf
    s = s & "              , s.[Location]" & vbCrLf
    s = s & "              , s.[Style]" & vbCrLf
    s = s & "              , s.[Finish]" & vbCrLf
    s = s & "              , s.[Other]" & vbCrLf
    s = s & "              , s.CRMType" & vbCrLf
    s = s & "              , s.Override_Price OverridePrice" & vbCrLf
    s = s & "              , s.seq AS ClientSeq" & vbCrLf
    s = s & "             , s.ModifiedBy" & vbCrLf
    s = s & "             , CONVERT(DATETIME, CONVERT(CHAR(19), s.ModifiedDate, 126)) AS ModifiedDate" & vbCrLf
    s = s & "             , s.EstimateIndex AS EstimateIndex" & vbCrLf
    s = s & "        FROM    tblScheduleB s" & vbCrLf
    s = s & "                INNER JOIN tblCustomers b ON s.Customer_No = b.Customer_No" & vbCrLf
    s = s & "        WHERE   s.ModifiedDate >= ( SELECT  lu.LastUpdateDate" & vbCrLf
    s = s & "                                    FROM    @tblLastUpdate lu" & vbCrLf
    s = s & "                                    WHERE   lu.TableName = 'tblQuotesAndSalesContract'" & vbCrLf
    s = s & "                                  )" & vbCrLf
    s = s & "                AND s.CRMID IS NOT NULL" & vbCrLf
    s = s & "                AND s.CRMType IS NOT NULL;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & " -- Change Orders" & vbCrLf
    s = s & "        SELECT  CO.Customer_No" & vbCrLf
    s = s & "              , CO.[Change_Order_No]" & vbCrLf
    s = s & "              , CO.[Change_Date]" & vbCrLf
    s = s & "              , CO.[Bldr_Approved_By]" & vbCrLf
    s = s & "              , CO.[Bldr_Approved_Date]" & vbCrLf
    s = s & "              , CO.[Cust_Approved_Date]" & vbCrLf
    s = s & "              , CO.[Paid_Date]" & vbCrLf
    s = s & "              , CO.[Next_Item_No]" & vbCrLf
    s = s & "              , CO.[Cust_Approved]" & vbCrLf
    s = s & "              , CO.[Bldr_Approved]" & vbCrLf
    s = s & "              , CO.[Total_Price]" & vbCrLf
    s = s & "              , CO.[Amount_Paid]" & vbCrLf
    s = s & "              , CO.[Addition_Change]" & vbCrLf
    s = s & "              , CO.[Declined]" & vbCrLf
    s = s & "              , CO.[Declined_By]" & vbCrLf
    s = s & "              , CO.[Notice_Estimating]" & vbCrLf
    s = s & "              , CO.[Office_Change]" & vbCrLf
    s = s & "              , CO.[Declined_Notice]" & vbCrLf
    s = s & "              , CO.[UserID]" & vbCrLf
    s = s & "              , CO.[NoCommTotal]" & vbCrLf
    s = s & "              , CO.[ChargedFee]" & vbCrLf
    s = s & "              , CO.[COLocked]" & vbCrLf
    s = s & "              , CO.[CreationDate]" & vbCrLf
    s = s & "              , CO.[RebateExempt]" & vbCrLf
    s = s & "              , CO.[WebUpdated]" & vbCrLf
    s = s & "              , CO.[NoDataChanged]" & vbCrLf
    s = s & "              , CO.[Sales_Person_ID]" & vbCrLf
    s = s & "              , CO.[contractexempt]" & vbCrLf
    s = s & "              , CO.[EnerGuideRatingSum]" & vbCrLf
    s = s & "              , CO.[GreenHouseRatingSum]" & vbCrLf
    s = s & "              , CO.[FuelCostSum]" & vbCrLf
    s = s & "              , CO.[REMOVEITEM]" & vbCrLf
    s = s & "              , CO.[SumOfTotalAmount]" & vbCrLf
    s = s & "              , CO.[SumOfNetTAX]" & vbCrLf
    s = s & "              , CO.[NotTaxedTotal]" & vbCrLf
    s = s & "              , CO.[ExportedToBMT]" & vbCrLf
    s = s & "              , CO.[DataChangedBMTExport]" & vbCrLf
    s = s & "              , CO.[EstimateIndex]" & vbCrLf
    s = s & "              , CO.[ApplyIncentive]" & vbCrLf
    s = s & "              , CO.[BillingReady]" & vbCrLf
    s = s & "              , CO.[AmountBilled]" & vbCrLf
    s = s & "              , CO.[Invoice]" & vbCrLf
    s = s & "              , CO.[UseAltJCCodes]" & vbCrLf
    s = s & "              , CO.[PurchasingCO]" & vbCrLf
    s = s & "              , CO.[comments]" & vbCrLf
    s = s & "              , CO.[Accounted_For]" & vbCrLf
    s = s & "              , CO.CRMID" & vbCrLf
    s = s & "              , c.CRMID CRMQuoteID" & vbCrLf
    s = s & "        FROM    ChangeOrderMaster CO" & vbCrLf
    s = s & "                INNER JOIN tblCustomers c ON CO.Customer_No = c.Customer_No" & vbCrLf
    s = s & "        WHERE   CO.CRMID IS NOT NULL" & vbCrLf
    s = s & "               AND CO.CRMType != 4" & vbCrLf
    s = s & "                AND CO.TSTMP >= ( SELECT lu.LastUpdateDate" & vbCrLf
    s = s & "                                     FROM   @tblLastUpdate lu" & vbCrLf
    s = s & "                                     WHERE  lu.TableName = 'tblChangeOrder'" & vbCrLf
    s = s & "                                   );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  CO.Customer_No" & vbCrLf
    s = s & "              , CO.[Change_Order_No]" & vbCrLf
    s = s & "              , COD.[Bldr_Approved_By]" & vbCrLf
    s = s & "              , COD.[Bldr_Approved_Date]" & vbCrLf
    s = s & "              , COD.[Paid_Date]" & vbCrLf
    s = s & "              , COD.[Item_No]" & vbCrLf
    s = s & "              , isnull(COD.[Custom],0) Custom" & vbCrLf
    s = s & "              , COD.[Structural_Change]" & vbCrLf
    s = s & "              , COD.[Require_Quote]" & vbCrLf
    s = s & "              , COD.[Description]" & vbCrLf
    s = s & "              , COD.[QTY]" & vbCrLf
    s = s & "              , COD.[Option_Type]" & vbCrLf
    s = s & "              , COD.[Email_Sent]" & vbCrLf
    s = s & "              , COD.[Cost_Amount]" & vbCrLf
    s = s & "              , COD.[Base_price]" & vbCrLf
    s = s & "              , COD.[Amount_Paid]" & vbCrLf
    s = s & "              , COD.[Rate]" & vbCrLf
    s = s & "              , COD.[Override_Price]" & vbCrLf
    s = s & "              , COD.[TL_extra]" & vbCrLf
    s = s & "              , COD.[Extra_Amount]" & vbCrLf
    s = s & "              , isnull(COD.[gstrebate_exempt],0) gstrebate_exempt" & vbCrLf
    s = s & "              , COD.[UOM]" & vbCrLf
    s = s & "              , COD.[OPT]" & vbCrLf
    s = s & "              , COD.[Category]" & vbCrLf
    s = s & "              , COD.[UserID]" & vbCrLf
    s = s & "              , COD.[Notice_Estimating]" & vbCrLf
    s = s & "              , COD.[NoCommission]" & vbCrLf
    s = s & "              , COD.[Major_Group]" & vbCrLf
    s = s & "              , COD.[CreationDate]" & vbCrLf
    s = s & "              , COD.[UnSentEmail]" & vbCrLf
    s = s & "              , COD.seq AS [item_seq]  -- This is an alias for a work arround because Seq is the key" & vbCrLf
    s = s & "              , COD.[ref1]" & vbCrLf
    s = s & "              , COD.[WebUpdated]" & vbCrLf
    s = s & "              , COD.[NoDataChanged]" & vbCrLf
    s = s & "              , COD.[PROCESSPO]" & vbCrLf
    s = s & "              , COD.[POCREATED]" & vbCrLf
    s = s & "              , COD.[OPTDELETED]" & vbCrLf
    s = s & "              , COD.[EnerGuideRating]" & vbCrLf
    s = s & "              , COD.[GreenHouseRating]" & vbCrLf
    s = s & "              , COD.[FuelCost]" & vbCrLf
    s = s & "              , COD.[REMOVEITEM]" & vbCrLf
    s = s & "              , COD.[UseTax]" & vbCrLf
    s = s & "              , COD.[NetTax]" & vbCrLf
    s = s & "              , COD.[TotalAmount]" & vbCrLf
    s = s & "              , COD.[GrandTotal]" & vbCrLf
    s = s & "              , COD.[ExportedToBMT]" & vbCrLf
    s = s & "              , COD.[DataChangedBMTExport]" & vbCrLf
    s = s & "              , COD.[StoredFile]" & vbCrLf
    s = s & "              , COD.[FileExtension]" & vbCrLf
    s = s & "              , COD.[EstimateIndex]" & vbCrLf
    s = s & "              , COD.[Assembly]" & vbCrLf
    s = s & "              , COD.[SalesWorksheet]" & vbCrLf
    s = s & "              , isnull(COD.[TAX],0) as Tax" & vbCrLf
    s = s & "              , COD.[ApplyIncentive]" & vbCrLf
    s = s & "              , COD.[IncentiveValue]" & vbCrLf
    s = s & "              , COD.[Declined]" & vbCrLf
    s = s & "              , COD.[EstimatorNotes]" & vbCrLf
    s = s & "              , COD.[AddendumSeq]" & vbCrLf
    s = s & "              , COD.[Bldr_Declined]" & vbCrLf
    s = s & "              , COD.[Bldr_Declined_Reason]" & vbCrLf
    s = s & "              , COD.[comments]" & vbCrLf
    s = s & "              , COD.[Color]" & vbCrLf
    s = s & "              , COD.[Location]" & vbCrLf
    s = s & "              , COD.[SpecSeq]" & vbCrLf
    s = s & "              , COD.[Style]" & vbCrLf
    s = s & "              , COD.[Finish]" & vbCrLf
    s = s & "              , COD.[Other]" & vbCrLf
    s = s & "              , COD.[ColorListID]" & vbCrLf
    s = s & "              , COD.[StyleListID]" & vbCrLf
    s = s & "              , COD.[FinishListID]" & vbCrLf
    s = s & "              , COD.[OtherListID]" & vbCrLf
    s = s & "              , COD.[quote_entered_date]" & vbCrLf
    s = s & "              , COD.[CRMID]" & vbCrLf
    s = s & "             , CONVERT(DATETIME, CONVERT(CHAR(19), COD.TSTMP, 126)) AS ModifiedDate" & vbCrLf
    s = s & "        FROM    ChangeOrderDetails COD" & vbCrLf
    s = s & "               INNER JOIN ChangeOrderMaster CO ON COD.Customer_No = CO.Customer_No" & vbCrLf
    s = s & "                   AND COD.Change_Order_No = CO.Change_Order_No" & vbCrLf
    s = s & "        WHERE   CO.CRMID IS NOT NULL AND CO.CRMType != 4" & vbCrLf
    s = s & "                AND CO.TSTMP >= ( SELECT  lu.LastUpdateDate" & vbCrLf
    s = s & "                                        FROM    @tblLastUpdate lu" & vbCrLf
    s = s & "                                        WHERE   lu.TableName = 'tblChangeOrder'" & vbCrLf
    s = s & "                                      );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & " -- Purchase Type" & vbCrLf
    s = s & "        SELECT  Item AS PurchaseTypeName" & vbCrLf
    s = s & "              , Custom_Description AS [Description]" & vbCrLf
    s = s & "        FROM    CustomDescriptions" & vbCrLf
    s = s & "        WHERE   Item IN ( 'PreSale', 'Spec Home', 'Show Home' );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "  -- Attribute Lists" & vbCrLf
    s = s & "        SELECT  al.ListID" & vbCrLf
    s = s & "              , al.Name AttributeListName" & vbCrLf
    s = s & "              , al.Required IsRequired" & vbCrLf
    s = s & "              , al.StrictList IsStrictList" & vbCrLf
    s = s & "        FROM    dbo.AttributeLists al;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  alv.ListID" & vbCrLf
    s = s & "              , alv.Value" & vbCrLf
    s = s & "              , ISNULL(alv.SortOrder, 0) SortOrder" & vbCrLf
    s = s & "             ,alv.OtherListID " & vbCrLf
    s = s & "        FROM    dbo.AttributeListValues alv;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "   -- Realtors" & vbCrLf
    s = s & "        SELECT  [RealtorID] ClientRealtorID" & vbCrLf
    s = s & "              , [Description] CompanyName" & vbCrLf
    s = s & "              , [RealtorName]" & vbCrLf
    s = s & "              , [Address1]" & vbCrLf
    s = s & "              , [Address2]" & vbCrLf
    s = s & "              , [City]" & vbCrLf
    s = s & "              , [Prov] State" & vbCrLf
    s = s & "              , [Zip]" & vbCrLf
    s = s & "              , [Phone]" & vbCrLf
    s = s & "              , [Fax]" & vbCrLf
    s = s & "              , [Email]" & vbCrLf
    s = s & "        FROM    [dbo].[Realtors]" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'Realtors'" & vbCrLf
    s = s & "                                ) OR ModifiedDate IS NULL;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "   -- Inventory Homes CO" & vbCrLf
    s = s & "        SELECT  i.Customer_No CustomerNo" & vbCrLf
    s = s & "              , i.Job_No JobNo" & vbCrLf
    s = s & "              , cm.Change_Order_No ChangeOrderNo" & vbCrLf
    s = s & "              , cm.Change_Date ChangeDate" & vbCrLf
    s = s & "              , cm.Bldr_Approved IsBldrApproved" & vbCrLf
    s = s & "              , cm.Bldr_Approved_By BldrApprovedBy" & vbCrLf
    s = s & "              , cm.Bldr_Approved_Date BldrApprovedDate" & vbCrLf
    s = s & "              , cm.Paid_Date PaidDate" & vbCrLf
    s = s & "              , cm.Next_Item_No NextItemNo" & vbCrLf
    s = s & "              , cm.Total_Price TotalPrice" & vbCrLf
    s = s & "              , isnull(cm.Amount_Paid,0) AmountPaid" & vbCrLf
    s = s & "              , cm.Addition_Change IsAdditionChange" & vbCrLf
    s = s & "              , cm.Declined IsDeclined" & vbCrLf
    s = s & "              , cm.Declined_By DeclinedBy" & vbCrLf
    s = s & "              , cm.Declined_Notice IsDeclinedNotice" & vbCrLf
    s = s & "              , cm.UserID ClientUserID" & vbCrLf
    s = s & "              , cm.NoCommTotal NoCommTotal" & vbCrLf
    s = s & "              , cm.ChargedFee ChangeFee" & vbCrLf
    s = s & "              , cm.COLocked IsCOLocked" & vbCrLf
    s = s & "              , cm.CreationDate" & vbCrLf
    s = s & "              , cm.RebateExempt IsRebateExempt" & vbCrLf
    s = s & "              , cm.Sales_Person_ID ClientSalesPersonID" & vbCrLf
    s = s & "              , cm.contractexempt IsContractExempt" & vbCrLf
    s = s & "              , cm.SumOfTotalAmount" & vbCrLf
    s = s & "              , cm.SumOfNetTAX" & vbCrLf
    s = s & "              , cm.NotTaxedTotal" & vbCrLf
    s = s & "              , cm.EstimateIndex" & vbCrLf
    s = s & "              , cm.ApplyIncentive IsApplyIncentive" & vbCrLf
    s = s & "              , cm.comments Comments" & vbCrLf
    s = s & "             , cm.CRMID EntityID" & vbCrLf
    s = s & "        FROM    tblCustomers i" & vbCrLf
    s = s & "                INNER JOIN ChangeOrderMaster cm ON i.Customer_No = cm.Customer_No" & vbCrLf
    s = s & "        WHERE    ( i.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "               AND ( i.Home_Selection = 'Spec Home' OR i.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "           )" & vbCrLf
    s = s & "           AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "           AND i.Sold =  0" & vbCrLf
    s = s & "           AND Inactive = 0" & vbCrLf
    s = s & "           --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "           --AND i.CRMID IS NULL" & vbCrLf
    s = s & "           AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "           AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "                AND cm.TSTMP >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                      FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                      WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                    );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  i.Customer_No CustomerNo" & vbCrLf
    s = s & "              , i.Job_No JobNo" & vbCrLf
    s = s & "              , cd.Change_Order_No ChangeOrderNo" & vbCrLf
    s = s & "              , cd.Bldr_Approved_By BldrApprovedBy" & vbCrLf
    s = s & "              , cd.Bldr_Approved_Date BldrApprovedDate" & vbCrLf
    s = s & "              , cd.Item_No ItemNo" & vbCrLf
    s = s & "              , cd.Custom IsCustom" & vbCrLf
    s = s & "              , cd.Require_Quote IsRequireQuote" & vbCrLf
    s = s & "              , cd.Description" & vbCrLf
    s = s & "              , cd.QTY" & vbCrLf
    s = s & "              , cd.Option_Type OptionType" & vbCrLf
    s = s & "              , cd.Email_Sent IsEmailSent" & vbCrLf
    s = s & "              , cd.Cost_Amount Cost" & vbCrLf
    s = s & "              , cd.Base_price PreTaxSellingPrice" & vbCrLf
    s = s & "              , cd.Rate" & vbCrLf
    s = s & "              , cd.Override_Price OverridePrice" & vbCrLf
    s = s & "              , cd.UOM" & vbCrLf
    s = s & "              , cd.OPT ClientOption" & vbCrLf
    s = s & "              , cd.Category" & vbCrLf
    s = s & "              , cd.UserID ClientUserID" & vbCrLf
    s = s & "              , cd.NoCommission IsNoCommission" & vbCrLf
    s = s & "              , cd.OPTDELETED IsOptDeleted" & vbCrLf
    s = s & "              , cd.EnerGuideRating" & vbCrLf
    s = s & "              , cd.GreenHouseRating" & vbCrLf
    s = s & "              , cd.FuelCost" & vbCrLf
    s = s & "              , cd.CreationDate" & vbCrLf
    s = s & "              , cd.UseTax IsUseTax" & vbCrLf
    s = s & "              , cd.NetTax" & vbCrLf
    s = s & "              , cd.GrandTotal" & vbCrLf
    s = s & "              , cd.EstimateIndex" & vbCrLf
    s = s & "              , cd.Assembly" & vbCrLf
    s = s & "              , cd.SalesWorksheet" & vbCrLf
    s = s & "              , cd.TAX" & vbCrLf
    s = s & "              , cd.ApplyIncentive IsApplyIncentive" & vbCrLf
    s = s & "              , cd.IncentiveValue" & vbCrLf
    s = s & "              , cd.Declined IsDeclined" & vbCrLf
    s = s & "              , cd.EstimatorNotes" & vbCrLf
    s = s & "              , cd.Bldr_Declined IsBldrDeclined" & vbCrLf
    s = s & "              , cd.Bldr_Declined_Reason BldrDeclinedReason" & vbCrLf
    s = s & "              , cd.comments Comments" & vbCrLf
    s = s & "              , cd.Color" & vbCrLf
    s = s & "              , cd.Location" & vbCrLf
    s = s & "              , cd.Style" & vbCrLf
    s = s & "              , cd.Finish" & vbCrLf
    s = s & "              , cd.Other" & vbCrLf
    s = s & "              , cd.ColorListID" & vbCrLf
    s = s & "              , cd.StyleListID" & vbCrLf
    s = s & "              , cd.FinishListID" & vbCrLf
    s = s & "              , cd.OtherListID" & vbCrLf
    s = s & "             , cd.CRMID EntityID" & vbCrLf
    s = s & "             , cd.seq ClientSeq" & vbCrLf
    s = s & "             , CONVERT(DATETIME, CONVERT(CHAR(19), CD.TSTMP, 126)) AS ModifiedDate" & vbCrLf
    s = s & "        FROM    tblCustomers i" & vbCrLf
    s = s & "                INNER JOIN ChangeOrderMaster cm ON i.Customer_No = cm.Customer_No" & vbCrLf
    s = s & "                INNER JOIN ChangeOrderDetails cd ON i.Customer_No = cd.Customer_No" & vbCrLf
    s = s & "                                                    AND cm.Change_Order_No = cd.Change_Order_No" & vbCrLf
    s = s & "        WHERE    ( i.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "               AND ( i.Home_Selection = 'Spec Home' OR i.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "           )" & vbCrLf
    s = s & "           AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "           AND i.Sold =  0" & vbCrLf
    s = s & "           AND Inactive = 0" & vbCrLf
    s = s & "           --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "           --AND i.CRMID IS NULL" & vbCrLf
    s = s & "           AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "           AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "                AND cd.TSTMP >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                      FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                      WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                    );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & " -- Inventory Homes Dates" & vbCrLf
    s = s & " -- Scheduling" & vbCrLf
    s = s & "        SELECT  i.Customer_No CustomerNo" & vbCrLf
    s = s & "              , i.Job_No JobNo" & vbCrLf
    s = s & "              , cd.Date_Field DateField" & vbCrLf
    s = s & "              , cd.Description" & vbCrLf
    s = s & "              , cd.Order_index OrderIndex" & vbCrLf
    s = s & "              , cd.CustDate StartDate" & vbCrLf
    s = s & "              , cd.Date_Completed FinishDate" & vbCrLf
    s = s & "              , cd.Check_Box IsCompleted" & vbCrLf
    s = s & "              , ISNULL(cd.Amount, 0)" & vbCrLf
    s = s & "              , ISNULL(cd.AmountPaid,0)" & vbCrLf
    s = s & "              , cd.PaidDate AmountPaidDate" & vbCrLf
    s = s & "              , cd.Supplier" & vbCrLf
    s = s & "              , cd.PercentComplete PercentCompleted" & vbCrLf
    s = s & "              , cd.Comments" & vbCrLf
    s = s & "              , cd.sales_view IsSalesView" & vbCrLf
    s = s & "              , cd.vendor_webview IsVendorWebView" & vbCrLf
    s = s & "              , cd.Web_View IsCustomerWebView" & vbCrLf
    s = s & "        FROM    Customer_Date cd" & vbCrLf
    s = s & "                INNER JOIN dbo.tblCustomers i ON i.Customer_No = cd.Customer_No" & vbCrLf
    s = s & "        WHERE    ( i.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "               AND ( i.Home_Selection = 'Spec Home' OR i.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "           )" & vbCrLf
    s = s & "           AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "           AND i.Sold =  0" & vbCrLf
    s = s & "           AND Inactive = 0" & vbCrLf
    s = s & "           --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "           --AND i.CRMID IS NULL" & vbCrLf
    s = s & "           AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "           AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "            AND cd.DateType = 2" & vbCrLf
    s = s & "            AND ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                    FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                    WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "               -- Condition" & vbCrLf
    s = s & "        SELECT  i.Customer_No CustomerNo" & vbCrLf
    s = s & "              , i.Job_No JobNo" & vbCrLf
    s = s & "              , cd.Date_Field DateField" & vbCrLf
    s = s & "              , cd.Description" & vbCrLf
    s = s & "              , cd.Order_index OrderIndex" & vbCrLf
    s = s & "              , cd.CustDate ExpectedRemovalDate" & vbCrLf
    s = s & "              , cd.Date_Completed ActualRemovalDate" & vbCrLf
    s = s & "              , cd.Check_Box IsCompleted" & vbCrLf
    s = s & "              , ISNULL(cd.Amount, 0)" & vbCrLf
    s = s & "              , cd.Comments" & vbCrLf
    s = s & "              , cd.sales_view IsSalesView" & vbCrLf
    s = s & "              , cd.vendor_webview IsVendorWebView" & vbCrLf
    s = s & "              , cd.Web_View IsCustomerWebView" & vbCrLf
    s = s & "        FROM    Customer_Date cd" & vbCrLf
    s = s & "                INNER JOIN dbo.tblCustomers i ON i.Customer_No = cd.Customer_No" & vbCrLf
    s = s & "        WHERE    ( i.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "               AND ( i.Home_Selection = 'Spec Home' OR i.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "           )" & vbCrLf
    s = s & "           AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "           AND i.Sold =  0" & vbCrLf
    s = s & "           AND Inactive = 0" & vbCrLf
    s = s & "           --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "           --AND i.CRMID IS NULL" & vbCrLf
    s = s & "           AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "           AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "            AND cd.DateType = 4" & vbCrLf
    s = s & "            AND ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                    FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                    WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- Deposit" & vbCrLf
    s = s & "        SELECT  i.Customer_No CustomerNo" & vbCrLf
    s = s & "              , i.Job_No JobNo" & vbCrLf
    s = s & "              , cd.Date_Field DateField" & vbCrLf
    s = s & "              , cd.Description" & vbCrLf
    s = s & "              , cd.Order_index OrderIndex" & vbCrLf
    s = s & "              , cd.CustDate DueDate" & vbCrLf
    s = s & "              , cd.Date_Completed ReceivedDate" & vbCrLf
    s = s & "              , cd.Check_Box IsCompleted" & vbCrLf
    s = s & "              , cd.notice_reqd IsNoticeRequired" & vbCrLf
    s = s & "              , ISNULL(cd.Amount, 0 )" & vbCrLf
    s = s & "              , ISNULL(cd.AmountPaid, 0)" & vbCrLf
    s = s & "              , cd.PaidDate AmountPaidDate" & vbCrLf
    s = s & "              , cd.Comments" & vbCrLf
    s = s & "              , cd.sales_view IsSalesView" & vbCrLf
    s = s & "              , cd.vendor_webview IsVendorWebView" & vbCrLf
    s = s & "              , cd.Web_View IsCustomerWebView" & vbCrLf
    s = s & "        FROM    Customer_Date cd" & vbCrLf
    s = s & "                INNER JOIN dbo.tblCustomers i ON i.Customer_No = cd.Customer_No" & vbCrLf
    s = s & "        WHERE    ( i.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "               AND ( i.Home_Selection = 'Spec Home' OR i.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "           )" & vbCrLf
    s = s & "           AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "           AND i.Sold =  0" & vbCrLf
    s = s & "           AND Inactive = 0" & vbCrLf
    s = s & "           --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "           --AND i.CRMID IS NULL" & vbCrLf
    s = s & "           AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "           AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "            AND cd.DateType = 3" & vbCrLf
    s = s & "            AND ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                    FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                    WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- Inventory Home Design Center" & vbCrLf
    s = s & "        SELECT  i.Customer_No CustomerNo" & vbCrLf
    s = s & "              , i.Job_No JobNo" & vbCrLf
    s = s & "              , cm.Change_Order_No ChangeOrderNo" & vbCrLf
    s = s & "              , cm.Change_Date ChangeDate" & vbCrLf
    s = s & "              , cm.Bldr_Approved IsBldrApproved" & vbCrLf
    s = s & "              , cm.Bldr_Approved_By BldrApprovedBy" & vbCrLf
    s = s & "              , cm.Bldr_Approved_Date BldrApprovedDate" & vbCrLf
    s = s & "              , cm.Paid_Date PaidDate" & vbCrLf
    s = s & "              , cm.Next_Item_No NextItemNo" & vbCrLf
    s = s & "              , cm.Total_Price TotalPrice" & vbCrLf
    s = s & "              , cm.Amount_Paid AmountPaid" & vbCrLf
    s = s & "              , 0 IsAdditionChange" & vbCrLf
    s = s & "              , cm.Declined IsDeclined" & vbCrLf
    s = s & "              , cm.Declined_By DeclinedBy" & vbCrLf
    s = s & "              , cm.Declined_Notice IsDeclinedNotice" & vbCrLf
    s = s & "              , cm.UserID ClientUserID" & vbCrLf
    s = s & "              , cm.NoCommTotal NoCommTotal" & vbCrLf
    s = s & "              , 0 ChangeFee" & vbCrLf
    s = s & "              , cm.COLocked IsCOLocked" & vbCrLf
    s = s & "              , cm.CreationDate" & vbCrLf
    s = s & "              , 0 IsRebateExempt" & vbCrLf
    s = s & "              , cm.Sales_Person_ID ClientSalesPersonID" & vbCrLf
    s = s & "              , cm.ContractExempt IsContractExempt" & vbCrLf
    s = s & "              , cm.SumOfTotalAmount" & vbCrLf
    s = s & "              , cm.SumOfNetTAX" & vbCrLf
    s = s & "              , cm.NotTaxedTotal" & vbCrLf
    s = s & "              , cm.EstimateIndex" & vbCrLf
    s = s & "              , cm.ApplyIncentive IsApplyIncentive" & vbCrLf
    s = s & "              , cm.Comments Comments" & vbCrLf
    s = s & "            -- , i.Sold" & vbCrLf
    s = s & "        FROM    tblCustomers i" & vbCrLf
    s = s & "                INNER JOIN dbo.DesignCenterMaster cm ON i.Customer_No = cm.Customer_No" & vbCrLf
    s = s & "        WHERE    ( i.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "               AND ( i.Home_Selection = 'Spec Home' OR i.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "           )" & vbCrLf
    s = s & "           AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "           AND i.Sold =  0" & vbCrLf
    s = s & "           AND Inactive = 0" & vbCrLf
    s = s & "           --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "           --AND i.CRMID IS NULL" & vbCrLf
    s = s & "           AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "           AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "            AND ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                    FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                    WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SELECT  i.Customer_No CustomerNo" & vbCrLf
    s = s & "              , i.Job_No JobNo" & vbCrLf
    s = s & "              , cd.Change_Order_No ChangeOrderNo" & vbCrLf
    s = s & "              , NULL BldrApprovedBy" & vbCrLf
    s = s & "              , NULL BldrApprovedDate" & vbCrLf
    s = s & "              , cd.Item_No ItemNo" & vbCrLf
    s = s & "              , cd.Custom IsCustom" & vbCrLf
    s = s & "              , cd.Require_Quote IsRequireQuote" & vbCrLf
    s = s & "              , cd.Description" & vbCrLf
    s = s & "              , cd.QTY" & vbCrLf
    s = s & "              , NULL OptionType" & vbCrLf
    s = s & "              , 0 IsEmailSent" & vbCrLf
    s = s & "              , cd.Cost_Amount Cost" & vbCrLf
    s = s & "              , cd.Amount PreTaxSellingPrice" & vbCrLf
    s = s & "              , cd.Rate" & vbCrLf
    s = s & "              , cd.GrandTotal OverridePrice" & vbCrLf
    s = s & "              , cd.UOM" & vbCrLf
    s = s & "              , cd.OPT ClientOption" & vbCrLf
    s = s & "              , cd.Major_Group AS ClientCategory" & vbCrLf
    s = s & "              , cd.Category AS ClientSubCategory" & vbCrLf
    s = s & "              , cd.UserID ClientUserID" & vbCrLf
    s = s & "              , cd.NoCommission IsNoCommission" & vbCrLf
    s = s & "              , cd.OPTDELETED IsOptDeleted" & vbCrLf
    s = s & "              , cd.EnerGuideRating" & vbCrLf
    s = s & "              , cd.GreenHouseRating" & vbCrLf
    s = s & "              , cd.FuelCost" & vbCrLf
    s = s & "              , cd.CreationDate" & vbCrLf
    s = s & "              , cd.UseTax IsUseTax" & vbCrLf
    s = s & "              , cd.NetTax" & vbCrLf
    s = s & "              , cd.GrandTotal" & vbCrLf
    s = s & "              , cd.EstimateIndex" & vbCrLf
    s = s & "              , cd.Assembly" & vbCrLf
    s = s & "              , cd.SalesWorksheet" & vbCrLf
    s = s & "              , cd.TAX" & vbCrLf
    s = s & "              , cd.ApplyIncentive IsApplyIncentive" & vbCrLf
    s = s & "              , cd.IncentiveValue" & vbCrLf
    s = s & "              , cd.Declined IsDeclined" & vbCrLf
    s = s & "              , cd.EstimatorNotes" & vbCrLf
    s = s & "              , cd.Bldr_Declined IsBldrDeclined" & vbCrLf
    s = s & "              , cd.Bldr_Declined_Reason BldrDeclinedReason" & vbCrLf
    s = s & "              , cd.comments Comments" & vbCrLf
    s = s & "              , cd.Color" & vbCrLf
    s = s & "              , cd.Location" & vbCrLf
    s = s & "              , cd.Style" & vbCrLf
    s = s & "              , cd.Finish" & vbCrLf
    s = s & "              , cd.Other" & vbCrLf
    s = s & "              , cd.ColorListID" & vbCrLf
    s = s & "              , cd.StyleListID" & vbCrLf
    s = s & "              , cd.FinishListID" & vbCrLf
    s = s & "              , cd.OtherListID" & vbCrLf
    s = s & "              , cd.StoredFile" & vbCrLf
    s = s & "              , cd.seq AS ClientSeq" & vbCrLf
    s = s & "             " & vbCrLf
    s = s & "        FROM    tblCustomers i" & vbCrLf
    s = s & "                INNER JOIN dbo.DesignCenterMaster cm ON i.Customer_No = cm.Customer_No" & vbCrLf
    s = s & "                INNER JOIN dbo.DesignCenterDetails cd ON i.Customer_No = cd.Customer_No" & vbCrLf
    s = s & "                                                         AND cm.Change_Order_No = cd.Change_Order_No" & vbCrLf
    s = s & "        WHERE    ( i.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "               AND i.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "               AND ( i.Home_Selection = 'Spec Home' OR i.Home_Selection = 'Show Home')" & vbCrLf
    s = s & "           )" & vbCrLf
    s = s & "           AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "           AND i.Sold =  0" & vbCrLf
    s = s & "           AND Inactive = 0" & vbCrLf
    s = s & "           --AND NotAvailableforSale = 0" & vbCrLf
    s = s & "           --AND i.CRMID IS NULL" & vbCrLf
    s = s & "           AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "           AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "            AND ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                    FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                    WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- Multi Family Parking" & vbCrLf
    s = s & "        SELECT  Community AS ClientCommunity" & vbCrLf
    s = s & "              , ParkingNo" & vbCrLf
    s = s & "              , Description" & vbCrLf
    s = s & "              , CostAmount" & vbCrLf
    s = s & "              , SellingPrice" & vbCrLf
    s = s & "              , UnitNo" & vbCrLf
    s = s & "              , LegalUnitNo" & vbCrLf
    s = s & "              , Comments" & vbCrLf
    s = s & "              , Sold" & vbCrLf
    s = s & "              , SoldToCustomer" & vbCrLf
    s = s & "              , IsParking" & vbCrLf
    s = s & "              , CommunityPhase AS ClientCommunityPhase" & vbCrLf
    s = s & "              , InventoryHomeID AS ClientIHID" & vbCrLf
    s = s & "              , ItemType" & vbCrLf
    s = s & "              , CRMID" & vbCrLf
    s = s & "        FROM    dbo.tblParking" & vbCrLf
    s = s & "        WHERE   ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblMultifamilyUnit'" & vbCrLf
    s = s & "                                );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       --Deleted Options" & vbCrLf
    s = s & "       SELECT  [CustomerNO]," & vbCrLf
    s = s & "                  [DeletedFromTable]," & vbCrLf
    s = s & "                  [CRMID]," & vbCrLf
    s = s & "                  [CRMTYPE]," & vbCrLf
    s = s & "                  [ProcessedInCRM]" & vbCrLf
    s = s & "       FROM DeletedOptionsforCRM" & vbCrLf
    s = s & "       WHERE ISNULL([ProcessedInCRM],0) = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       SELECT [Job_Status] [JobStatusID]" & vbCrLf
    s = s & "               ,[Description]" & vbCrLf
    s = s & "               ,[CreatedDate]" & vbCrLf
    s = s & "               ,[ModifiedDate]" & vbCrLf
    s = s & "       from tblJobStatus" & vbCrLf
    s = s & "       WHERE ModifiedDate >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                      FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                      WHERE     lu.TableName = 'tblJobStatus'" & vbCrLf
    s = s & "                                    );" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       SELECT TableName, LastUpdateDate FROM @tblLastUpdate;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       -- Law_Firms & Lending_Companies are stored in the same table in CRM, " & vbCrLf
    s = s & "       SELECT  Law_Firm ClientID" & vbCrLf
    s = s & "             , Firm_Name Name" & vbCrLf
    s = s & "             , 0 IsLenderCompany" & vbCrLf
    s = s & "       FROM    dbo.Law_Firms" & vbCrLf
    s = s & "       WHERE   ISNULL(Law_Firm, '') != ''" & vbCrLf
    s = s & "       UNION ALL" & vbCrLf
    s = s & "       SELECT  Lending_Company ClientID" & vbCrLf
    s = s & "             , Company_Name Name" & vbCrLf
    s = s & "             , 1 IsLenderCompany" & vbCrLf
    s = s & "       FROM    dbo.Lending_Companies" & vbCrLf
    s = s & "       WHERE   ISNULL(Lending_Company, '') != '';" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       --Deleted Global Options, Options and Models " & vbCrLf
    s = s & "       SELECT  DivisionID" & vbCrLf
    s = s & "             , DeletedFromTable" & vbCrLf
    s = s & "             , RowSequence" & vbCrLf
    s = s & "             , DateDeleted" & vbCrLf
    s = s & "             , DeletedBy" & vbCrLf
    s = s & "             , ProcessedInCRM" & vbCrLf
    s = s & "             , Community" & vbCrLf
    s = s & "             , Model" & vbCrLf
    s = s & "             , Series" & vbCrLf
    s = s & "             , OptionID" & vbCrLf
    s = s & "             , Description" & vbCrLf
    s = s & "       FROM    dbo.DeletedMasterListDataforCRM" & vbCrLf
    s = s & "       WHERE   ISNULL([ProcessedInCRM], 0) = 0 and DeletedFromTable<>'tblLotInventory';" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       Select DivisionID,Room from RoomMaster;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       Select m.DivisionID,m.Room,c.Subcategory,c.UOM " & vbCrLf
    s = s & "       from RoomMasterSubCategory c" & vbCrLf
    s = s & "       Join RoomMaster m on m.RoomID = c.RoomID" & vbCrLf
    s = s & "       ;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       select m.DivisionID,m.Room,q.Model,q.Subcategory,q.UOM,q.Qty " & vbCrLf
    s = s & "       from RoomQtybyModel q" & vbCrLf
    s = s & "       Join RoomMaster m on m.RoomID = q.RoomID" & vbCrLf
    s = s & "       ;" & vbCrLf
    s = s & "-- Customers to insert into CRM" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "select [Customer_No] ," & vbCrLf
    s = s & "   [Community] ," & vbCrLf
    s = s & "   [DivisionID] ," & vbCrLf
    s = s & "   Sales_Person_ID ," & vbCrLf
    s = s & "   [Description] ," & vbCrLf
    s = s & "   [Job_No] ," & vbCrLf
    s = s & "   [Expected_Occupancy] ," & vbCrLf
    s = s & "   [Lot_No] ," & vbCrLf
    s = s & "   [Lot] ," & vbCrLf
    s = s & "   [Block] ," & vbCrLf
    s = s & "   [Phase] ," & vbCrLf
    s = s & "   [LotPlan] ," & vbCrLf
    s = s & "   [Approved_By] ," & vbCrLf
    s = s & "   [Approved_Date] ," & vbCrLf
    s = s & "   Contract_Assigned ," & vbCrLf
    s = s & "   Contract_Signed_Date ," & vbCrLf
    s = s & "   Pos_Date ," & vbCrLf
    s = s & "   Construction_Status  ," & vbCrLf
    s = s & "   [Comments]  ," & vbCrLf
    s = s & "   Municipal_Address ," & vbCrLf
    s = s & "   [County] ," & vbCrLf
    s = s & "   [Township] ," & vbCrLf
    s = s & "   [LegalAddress] ," & vbCrLf
    s = s & "   Square_Footage ," & vbCrLf
    s = s & "   ChkGST ," & vbCrLf
    s = s & "   ChkGstRebate ," & vbCrLf
    s = s & "   Last_CO ," & vbCrLf
    s = s & "   Last_Add_Chng ," & vbCrLf
    s = s & "   Cancelled," & vbCrLf
    s = s & "   Home_Selection ," & vbCrLf
    s = s & "   PreSale_Selection ," & vbCrLf
    s = s & "   Sold ," & vbCrLf
    s = s & "   Last_DC ," & vbCrLf
    s = s & "   [IntUID] ," & vbCrLf
    s = s & "   [IntPwd]  ," & vbCrLf
    s = s & "   [Referral] ," & vbCrLf
    s = s & "   Approved ," & vbCrLf
    s = s & "   [Purchased] ," & vbCrLf
    s = s & "   [Purchased_Date] ," & vbCrLf
    s = s & "   [Unit_No] ," & vbCrLf
    s = s & "   [Legal_Unit_No] ," & vbCrLf
    s = s & "   [Floor] ," & vbCrLf
    s = s & "   [LotNetTAX] [float] ," & vbCrLf
    s = s & "   [ModelAssembly] ," & vbCrLf
    s = s & "   [ModelSalesWorksheet] ," & vbCrLf
    s = s & "   [ModelSpecDoc] ," & vbCrLf
    s = s & "   [ModelSpecDate] ," & vbCrLf
    s = s & "   [RealtorID] ," & vbCrLf
    s = s & "   [SellingAgent] ," & vbCrLf
    s = s & "   [CancelledDate] ," & vbCrLf
    s = s & "    CONVERT(DATETIME, CONVERT(CHAR(19), [CreatedDate], 126)) [CreatedDate] ," & vbCrLf
    s = s & "    CONVERT(DATETIME, CONVERT(CHAR(19), [ModifiedDate], 126)) [ModifiedDate] ," & vbCrLf
    s = s & "   [PST] ," & vbCrLf
    s = s & "   [PSTRebate] ," & vbCrLf
    s = s & "   [GST] ," & vbCrLf
    s = s & "   [GST_Rebate] ," & vbCrLf
    s = s & "   [GST_Rate] ," & vbCrLf
    s = s & "   [PST_Rate] ," & vbCrLf
    s = s & "   Base_House ," & vbCrLf
    s = s & "   Override_Price ," & vbCrLf
    s = s & "   Lot_Price ," & vbCrLf
    s = s & "   [LotPremiumValue] ," & vbCrLf
    s = s & "   ScheduleB_total ," & vbCrLf
    s = s & "   [CRMType] ," & vbCrLf
    s = s & "   [Closed] ," & vbCrLf
    s = s & "   [ModelIncentive] ," & vbCrLf
    s = s & "   Sales_Initiative," & vbCrLf
    s = s & "   sales_deduction ," & vbCrLf
    s = s & "   [TotalIncentives] ," & vbCrLf
    s = s & "   [ModifiedBy] ," & vbCrLf
    s = s & "   [Model] ," & vbCrLf
    s = s & "   [Series] ," & vbCrLf
    s = s & "   [UnitFactor] ," & vbCrLf
    s = s & "   [CondoFees] ," & vbCrLf
    s = s & "   [Customer_Name] ," & vbCrLf
    s = s & "   Customer_LName ," & vbCrLf
    s = s & "   '' [MiddleName] ," & vbCrLf
    s = s & "   [Email] ," & vbCrLf
    s = s & "   [Address1] ," & vbCrLf
    s = s & "   [Address2] ," & vbCrLf
    s = s & "   [City] ," & vbCrLf
    s = s & "   Province ," & vbCrLf
    s = s & "   [Zip] ," & vbCrLf
    s = s & "    Phone ," & vbCrLf
    s = s & "   W_Phone ," & vbCrLf
    s = s & "   CellPhone ," & vbCrLf
    s = s & "   CoBuyer_Name," & vbCrLf
    s = s & "   CoBuyer_LName ," & vbCrLf
    s = s & "   '' [CoBuyerMiddleName] ," & vbCrLf
    s = s & "   C_Email ," & vbCrLf
    s = s & "   [C_Address1] ," & vbCrLf
    s = s & "   [C_Address2] ," & vbCrLf
    s = s & "   [C_City] ," & vbCrLf
    s = s & "   [C_Province] ," & vbCrLf
    s = s & "   [C_Zip] ," & vbCrLf
    s = s & "   [C_Phone] ," & vbCrLf
    s = s & "   C_W_Phone ," & vbCrLf
    s = s & "   [C_CellPhone] " & vbCrLf
    s = s & "        FROM  dbo.tblCustomers" & vbCrLf
    s = s & "        WHERE CRMID IS  NULL and cancelled = 0 " & vbCrLf
    s = s & "       and purchased = 1 " & vbCrLf
    s = s & "       and approved = 1 " & vbCrLf
    s = s & "       and pos_date is null " & vbCrLf
    s = s & "       and recordtype like '%Customer%' " & vbCrLf
    s = s & "       and isnull(IsCRMQuote,0)=0" & vbCrLf
    s = s & "       and sendtoCRM = 1;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "--ScheduleB" & vbCrLf
    s = s & "        SELECT  s.Customer_No CustomerNo" & vbCrLf
    s = s & "              , s.OPT [Option]" & vbCrLf
    s = s & "              , s.Option_Type [OptionType]" & vbCrLf
    s = s & "              , s.Custom [IsCustom]" & vbCrLf
    s = s & "              , s.Require_Quote [RequireQuote]" & vbCrLf
    s = s & "              , s.[Description]" & vbCrLf
    s = s & "              , s.[QTY]" & vbCrLf
    s = s & "              , s.Cost_Amount [CostAmount]" & vbCrLf
    s = s & "              , s.Rate [Rate]" & vbCrLf
    s = s & "              , s.GrandTotal [TotalPrice]" & vbCrLf
    s = s & "              , s.[comments]" & vbCrLf
    s = s & "              , s.[UOM]" & vbCrLf
    s = s & "              , s.Major_Group [CategoryName]" & vbCrLf
    s = s & "              , s.Category [SubCategoryName]" & vbCrLf
    s = s & "              , s.[Assembly]" & vbCrLf
    s = s & "              , s.[SalesWorksheet]" & vbCrLf
    s = s & "              , s.TAX [TaxAmount]" & vbCrLf
    s = s & "              , s.AddFromSpec [AddFromSpec]" & vbCrLf
    s = s & "              , s.SpecDetailReference [SpecDetailReference]" & vbCrLf
    s = s & "              , s.EstimatorNotes [EstimatorNotes]" & vbCrLf
    s = s & "              , s.CONumber [CONumber]" & vbCrLf
    s = s & "              , s.Bldr_Declined [IsBldrDeclined]" & vbCrLf
    s = s & "              , s.Bldr_Declined_Reason [BldrDeclinedReason]" & vbCrLf
    s = s & "              , s.[Color]" & vbCrLf
    s = s & "              , s.[Location]" & vbCrLf
    s = s & "              , s.[Style]" & vbCrLf
    s = s & "              , s.[Finish]" & vbCrLf
    s = s & "              , s.[Other]" & vbCrLf
    s = s & "              , s.CRMType" & vbCrLf
    s = s & "              , s.Override_Price OverridePrice" & vbCrLf
    s = s & "              , s.seq AS ClientSeq" & vbCrLf
    s = s & "             , s.ModifiedBy" & vbCrLf
    s = s & "             , CONVERT(DATETIME, CONVERT(CHAR(19), s.ModifiedDate, 126)) AS ModifiedDate" & vbCrLf
    s = s & "             , s.EstimateIndex AS EstimateIndex" & vbCrLf
    s = s & "        FROM    tblScheduleB s" & vbCrLf
    s = s & "                INNER JOIN tblCustomers b ON s.Customer_No = b.Customer_No" & vbCrLf
    s = s & "               WHERE b.CRMID IS  NULL and s.CRMID is null and cancelled = 0 " & vbCrLf
    s = s & "               and b.purchased = 1 " & vbCrLf
    s = s & "               and b.approved = 1 " & vbCrLf
    s = s & "               and b.pos_date is null " & vbCrLf
    s = s & "               and b.recordtype like '%Customer%' " & vbCrLf
    s = s & "               and isnull(IsCRMQuote,0)=0" & vbCrLf
    s = s & "               and b.sendtoCRM = 1;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        " & vbCrLf
    s = s & "" & vbCrLf
    s = s & " -- Change Orders" & vbCrLf
    s = s & "        SELECT  CO.Customer_No" & vbCrLf
    s = s & "              , CO.[Change_Order_No]" & vbCrLf
    s = s & "              , CO.[Change_Date]" & vbCrLf
    s = s & "              , CO.[Bldr_Approved_By]" & vbCrLf
    s = s & "              , CO.[Bldr_Approved_Date]" & vbCrLf
    s = s & "              , CO.[Cust_Approved_Date]" & vbCrLf
    s = s & "              , CO.[Paid_Date]" & vbCrLf
    s = s & "              , CO.[Next_Item_No]" & vbCrLf
    s = s & "              , CO.[Cust_Approved]" & vbCrLf
    s = s & "              , CO.[Bldr_Approved]" & vbCrLf
    s = s & "              , CO.[Total_Price]" & vbCrLf
    s = s & "              , CO.[Amount_Paid]" & vbCrLf
    s = s & "              , CO.[Addition_Change]" & vbCrLf
    s = s & "              , CO.[Declined]" & vbCrLf
    s = s & "              , CO.[Declined_By]" & vbCrLf
    s = s & "              , CO.[Notice_Estimating]" & vbCrLf
    s = s & "              , CO.[Office_Change]" & vbCrLf
    s = s & "              , CO.[Declined_Notice]" & vbCrLf
    s = s & "              , CO.[UserID]" & vbCrLf
    s = s & "              , CO.[NoCommTotal]" & vbCrLf
    s = s & "              , CO.[ChargedFee]" & vbCrLf
    s = s & "              , CO.[COLocked]" & vbCrLf
    s = s & "              , CO.[CreationDate]" & vbCrLf
    s = s & "              , CO.[RebateExempt]" & vbCrLf
    s = s & "              , CO.[WebUpdated]" & vbCrLf
    s = s & "              , CO.[NoDataChanged]" & vbCrLf
    s = s & "              , CO.[Sales_Person_ID]" & vbCrLf
    s = s & "              , CO.[contractexempt]" & vbCrLf
    s = s & "              , CO.[EnerGuideRatingSum]" & vbCrLf
    s = s & "              , CO.[GreenHouseRatingSum]" & vbCrLf
    s = s & "              , CO.[FuelCostSum]" & vbCrLf
    s = s & "              , CO.[REMOVEITEM]" & vbCrLf
    s = s & "              , CO.[SumOfTotalAmount]" & vbCrLf
    s = s & "              , CO.[SumOfNetTAX]" & vbCrLf
    s = s & "              , CO.[NotTaxedTotal]" & vbCrLf
    s = s & "              , CO.[ExportedToBMT]" & vbCrLf
    s = s & "              , CO.[DataChangedBMTExport]" & vbCrLf
    s = s & "              , CO.[EstimateIndex]" & vbCrLf
    s = s & "              , CO.[ApplyIncentive]" & vbCrLf
    s = s & "              , CO.[BillingReady]" & vbCrLf
    s = s & "              , CO.[AmountBilled]" & vbCrLf
    s = s & "              , CO.[Invoice]" & vbCrLf
    s = s & "              , CO.[UseAltJCCodes]" & vbCrLf
    s = s & "              , CO.[PurchasingCO]" & vbCrLf
    s = s & "              , CO.[comments]" & vbCrLf
    s = s & "              , CO.[Accounted_For]" & vbCrLf
    s = s & "              , CO.CRMID" & vbCrLf
    s = s & "              , c.CRMID CRMQuoteID" & vbCrLf
    s = s & "        FROM    ChangeOrderMaster CO" & vbCrLf
    s = s & "                INNER JOIN tblCustomers c ON CO.Customer_No = c.Customer_No" & vbCrLf
    s = s & "               WHERE CO.CRMID is null AND c.CRMID IS  NULL AND cancelled = 0 " & vbCrLf
    s = s & "               AND c.purchased = 1 " & vbCrLf
    s = s & "               AND c.approved = 1 " & vbCrLf
    s = s & "               AND c.pos_date is null " & vbCrLf
    s = s & "               AND c.recordtype like '%Customer%' " & vbCrLf
    s = s & "               AND isnull(c.IsCRMQuote,0)=0" & vbCrLf
    s = s & "               and c.sendtoCRM = 1;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "-- Change Order Details" & vbCrLf
    s = s & "        SELECT  CO.Customer_No" & vbCrLf
    s = s & "              , CO.[Change_Order_No]" & vbCrLf
    s = s & "              , COD.[Bldr_Approved_By]" & vbCrLf
    s = s & "              , COD.[Bldr_Approved_Date]" & vbCrLf
    s = s & "              , COD.[Paid_Date]" & vbCrLf
    s = s & "              , COD.[Item_No]" & vbCrLf
    s = s & "              , COD.[Custom]" & vbCrLf
    s = s & "              , COD.[Structural_Change]" & vbCrLf
    s = s & "              , COD.[Require_Quote]" & vbCrLf
    s = s & "              , COD.[Description]" & vbCrLf
    s = s & "              , COD.[QTY]" & vbCrLf
    s = s & "              , COD.[Option_Type]" & vbCrLf
    s = s & "              , COD.[Email_Sent]" & vbCrLf
    s = s & "              , COD.[Cost_Amount]" & vbCrLf
    s = s & "              , COD.[Base_price]" & vbCrLf
    s = s & "              , COD.[Amount_Paid]" & vbCrLf
    s = s & "              , COD.[Rate]" & vbCrLf
    s = s & "              , COD.[Override_Price]" & vbCrLf
    s = s & "              , COD.[TL_extra]" & vbCrLf
    s = s & "              , COD.[Extra_Amount]" & vbCrLf
    s = s & "              , isnull(COD.[gstrebate_exempt],0) as Gstrebate_exempt" & vbCrLf
    s = s & "              , COD.[UOM]" & vbCrLf
    s = s & "              , COD.[OPT]" & vbCrLf
    s = s & "              , COD.[Category]" & vbCrLf
    s = s & "              , COD.[UserID]" & vbCrLf
    s = s & "              , COD.[Notice_Estimating]" & vbCrLf
    s = s & "              , COD.[NoCommission]" & vbCrLf
    s = s & "              , COD.[Major_Group]" & vbCrLf
    s = s & "              , COD.[CreationDate]" & vbCrLf
    s = s & "              , COD.[UnSentEmail]" & vbCrLf
    s = s & "              , COD.seq AS [item_seq]  -- This is an alias for a work arround because Seq is the key" & vbCrLf
    s = s & "              , COD.[ref1]" & vbCrLf
    s = s & "              , COD.[WebUpdated]" & vbCrLf
    s = s & "              , COD.[NoDataChanged]" & vbCrLf
    s = s & "              , COD.[PROCESSPO]" & vbCrLf
    s = s & "              , COD.[POCREATED]" & vbCrLf
    s = s & "              , COD.[OPTDELETED]" & vbCrLf
    s = s & "              , COD.[EnerGuideRating]" & vbCrLf
    s = s & "              , COD.[GreenHouseRating]" & vbCrLf
    s = s & "              , COD.[FuelCost]" & vbCrLf
    s = s & "              , COD.[REMOVEITEM]" & vbCrLf
    s = s & "              , COD.[UseTax]" & vbCrLf
    s = s & "              , COD.[NetTax]" & vbCrLf
    s = s & "              , COD.[TotalAmount]" & vbCrLf
    s = s & "              , COD.[GrandTotal]" & vbCrLf
    s = s & "              , COD.[ExportedToBMT]" & vbCrLf
    s = s & "              , COD.[DataChangedBMTExport]" & vbCrLf
    s = s & "              , COD.[StoredFile]" & vbCrLf
    s = s & "              , COD.[FileExtension]" & vbCrLf
    s = s & "              , COD.[EstimateIndex]" & vbCrLf
    s = s & "              , COD.[Assembly]" & vbCrLf
    s = s & "              , COD.[SalesWorksheet]" & vbCrLf
    s = s & "              , isnull(COD.[TAX],0) Tax" & vbCrLf
    s = s & "              , COD.[ApplyIncentive]" & vbCrLf
    s = s & "              , COD.[IncentiveValue]" & vbCrLf
    s = s & "              , COD.[Declined]" & vbCrLf
    s = s & "              , COD.[EstimatorNotes]" & vbCrLf
    s = s & "              , COD.[AddendumSeq]" & vbCrLf
    s = s & "              , COD.[Bldr_Declined]" & vbCrLf
    s = s & "              , COD.[Bldr_Declined_Reason]" & vbCrLf
    s = s & "              , COD.[comments]" & vbCrLf
    s = s & "              , COD.[Color]" & vbCrLf
    s = s & "              , COD.[Location]" & vbCrLf
    s = s & "              , COD.[SpecSeq]" & vbCrLf
    s = s & "              , COD.[Style]" & vbCrLf
    s = s & "              , COD.[Finish]" & vbCrLf
    s = s & "              , COD.[Other]" & vbCrLf
    s = s & "              , COD.[ColorListID]" & vbCrLf
    s = s & "              , COD.[StyleListID]" & vbCrLf
    s = s & "              , COD.[FinishListID]" & vbCrLf
    s = s & "              , COD.[OtherListID]" & vbCrLf
    s = s & "              , COD.[quote_entered_date]" & vbCrLf
    s = s & "              , COD.[CRMID]" & vbCrLf
    s = s & "             , CONVERT(DATETIME, CONVERT(CHAR(19), COD.TSTMP, 126)) AS ModifiedDate" & vbCrLf
    s = s & "        FROM    ChangeOrderDetails COD" & vbCrLf
    s = s & "               INNER JOIN ChangeOrderMaster CO ON COD.Customer_No = CO.Customer_No AND COD.Change_Order_No = CO.Change_Order_No" & vbCrLf
    s = s & "               INNER JOIN tblCustomers c ON CO.Customer_No = c.Customer_No" & vbCrLf
    s = s & "        WHERE   CO.CRMID IS NULL AND c.CRMID IS  NULL AND cancelled = 0 " & vbCrLf
    s = s & "               AND c.purchased = 1 " & vbCrLf
    s = s & "               AND c.approved = 1 " & vbCrLf
    s = s & "               AND c.pos_date is null " & vbCrLf
    s = s & "               AND c.recordtype like '%Customer%' " & vbCrLf
    s = s & "               AND isnull(c.IsCRMQuote,0)=0" & vbCrLf
    s = s & "               and c.sendtoCRM = 1" & vbCrLf
    s = s & "               ;                     " & vbCrLf
    s = s & "                                     " & vbCrLf
    s = s & "                                     ;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    END;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "   select getdate()" & vbCrLf
    
    DT_Get_DataForCRM = s

End Function

Private Function DT_Set_DataFromCRM_a() As String
    Dim s As String


s = ""
s = s & "CREATE PROCEDURE [dbo].[DT_Set_DataFromCRM]  " & vbCrLf
s = s & "    (  " & vbCrLf
s = s & "     @tblSalesContract dbo.[tblSalesContract_Type] READONLY  " & vbCrLf
s = s & "   , @tblSalesContractAddendum dbo.[tblSalesContractAddendum_Type] READONLY  " & vbCrLf
s = s & "   , @tblQuotes dbo.[tblQuotes_Type] READONLY  " & vbCrLf
s = s & "   , @tblQuoteDetails dbo.[tblQuoteDetails_Type] READONLY  " & vbCrLf
s = s & "   , @tblChangeOrderMaster dbo.[ChangeOrderMaster_Type] READONLY  " & vbCrLf
s = s & "   , @tblChangeOrderDetails dbo.[ChangeOrderDetails_Type] READONLY  " & vbCrLf
s = s & "   , @tblDepositDates dbo.[tblDepositDates_Type] READONLY  " & vbCrLf
s = s & "   , @tblSchedulingDates dbo.[tblSchedulingDates_Type] READONLY  " & vbCrLf
s = s & "   , @tblConditionDates dbo.[tblConditionDates_Type] READONLY  " & vbCrLf
s = s & "   , @tblInventoryHome dbo.[tblInventoryHome_Type] READONLY  " & vbCrLf
s = s & "   , @tblInventoryHomesAddendum dbo.[tblInventoryHomesAddendum_Type] READONLY  " & vbCrLf
s = s & "   , @tblInventoryHomesCOMaster dbo.[tblInventoryHomesCOMaster_Type] READONLY  " & vbCrLf
s = s & "   , @tblInventoryHomesCODetails dbo.[tblInventoryHomesCODetails_Type] READONLY  " & vbCrLf
s = s & "   , @tblMiscDates dbo.[tblMiscDates_Type] READONLY  " & vbCrLf
s = s & "   , @tblMultifamilyUnit dbo.[tblMultifamilyUnit_Type] READONLY  " & vbCrLf
s = s & "   , @tblLastPullDate dbo.[tblLastPullDate_Type] READONLY  " & vbCrLf
s = s & "   , @tblTitleCompany dbo.[tblTitleCompany_Type] READONLY  " & vbCrLf
s = s & "   , @tblTitleCompanyContact dbo.[tblTitleCompanyContact_Type] READONLY  " & vbCrLf
s = s & "   , @tblRealtors dbo.[tblRealtors_Type] READONLY  " & vbCrLf
s = s & "   , @tblModels dbo.[tblModels_Type] READONLY  " & vbCrLf
s = s & "   , @tblDeletedOptionsForHomefront dbo.[DeletedOptionsForHomefront_Type] READONLY  " & vbCrLf
s = s & "   , @tblLotStatusUpdates dbo.[LotStatusUpdates_Type] READONLY  " & vbCrLf
s = s & "   , @tblCustomerDesignSelections dbo.[CustomerDesignSelections_Type] READONLY  " & vbCrLf
s = s & "    )  " & vbCrLf
s = s & "AS  " & vbCrLf
s = s & "    BEGIN  " & vbCrLf
s = s & "  " & vbCrLf
s = s & " --DECLARE @LastPullDate DATETIME  " & vbCrLf
s = s & " --SELECT TOP 1 @LastPullDate = LastPullDate FROM @tblLastPullDate  " & vbCrLf
s = s & "  " & vbCrLf
s = s & " DECLARE @TimeZoneUTCOffSet INTEGER   " & vbCrLf
s = s & " --SELECT TOP 1 @TimeZoneUTCOffSet = TimeZoneOffSet FROM @tblLastPullDate  " & vbCrLf
s = s & " set @TimeZoneUTCOffSet  = isnull((select datediff(HOUR,TimeZoneOffSet,getdate()) from @tblLastPullDate),0)" & vbCrLf
s = s & "" & vbCrLf
s = s & " DECLARE @sqlStatement varchar(50)" & vbCrLf
s = s & " " & vbCrLf
s = s & " DECLARE @IgnoreApprovedDate varchar(10)" & vbCrLf
s = s & "" & vbCrLf
s = s & " set @IgnoreApprovedDate = isnull((Select isnull(OptionValue,'') from AppOptions where OptionName = 'IgnoreApprovedDate' and divisionid = 1),'')" & vbCrLf
s = s & " " & vbCrLf
s = s & " " & vbCrLf
s = s & " BEGIN TRY " & vbCrLf
s = s & " BEGIN TRANSACTION " & vbCrLf
s = s & "-- Sales Contract  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & " set @sqlStatement = 'Insert tblCustomers from @tblSalesContract'" & vbCrLf
s = s & "" & vbCrLf
s = s & "if (@IgnoreApprovedDate='True') BEGIN  " & vbCrLf
s = s & "INSERT  INTO tblCustomers  " & vbCrLf
s = s & "                (CRMID  " & vbCrLf
s = s & "   , Customer_No  " & vbCrLf
s = s & "   , Community  " & vbCrLf
s = s & "   , DivisionID  " & vbCrLf
s = s & "" & vbCrLf
s = s & "   , Customer_Name  " & vbCrLf
s = s & "   , Customer_LName  " & vbCrLf
s = s & "   , Phone  " & vbCrLf
s = s & "   , Fax  " & vbCrLf
s = s & "   , Email  " & vbCrLf
s = s & "   , Address1  " & vbCrLf
s = s & "   , Address2  " & vbCrLf
s = s & "   , City  " & vbCrLf
s = s & "   , Province  " & vbCrLf
s = s & "   , Zip  " & vbCrLf
s = s & "" & vbCrLf
s = s & "   , CoBuyer_Name  " & vbCrLf
s = s & "   , CoBuyer_LName  " & vbCrLf
s = s & "   , C_Phone  " & vbCrLf
s = s & "   , C_Fax  " & vbCrLf
s = s & "   , C_Email   " & vbCrLf
s = s & "   , C_Address1  " & vbCrLf
s = s & "   , C_Address2  " & vbCrLf
s = s & "   , C_City  " & vbCrLf
s = s & "   , C_Province  " & vbCrLf
s = s & "   , C_Zip  " & vbCrLf
s = s & "" & vbCrLf
s = s & "   , Model  " & vbCrLf
s = s & "   , Series  " & vbCrLf
s = s & "   , Sales_Person_ID  " & vbCrLf
s = s & "   , [Description]  " & vbCrLf
s = s & "   , Job_No  " & vbCrLf
s = s & "   , Expected_Occupancy  " & vbCrLf
s = s & "   , Lot  " & vbCrLf
s = s & "   , Lot_No  " & vbCrLf
s = s & "   , Block  " & vbCrLf
s = s & "   , Phase  " & vbCrLf
s = s & "   , LOTPLAN  " & vbCrLf
s = s & "   , Approved_By  " & vbCrLf
s = s & "   --        , Approved_Date  " & vbCrLf
s = s & "   , Contract_Assigned  " & vbCrLf
s = s & "   , Contract_Signed_Date  " & vbCrLf
s = s & "   , Pos_Date  " & vbCrLf
s = s & "   , Construction_Status  " & vbCrLf
s = s & "   , Comments  " & vbCrLf
s = s & "   , Municipal_Address  " & vbCrLf
s = s & "   , County  " & vbCrLf
s = s & "   , Township  " & vbCrLf
s = s & "   , LegalAddress  " & vbCrLf
s = s & "   , Square_Footage  " & vbCrLf
s = s & "   , ChkGST  " & vbCrLf
s = s & "   , ChkGstRebate  " & vbCrLf
s = s & "   , Last_CO  " & vbCrLf
s = s & "   , Last_Add_Chng  " & vbCrLf
s = s & "   , Cancelled  " & vbCrLf
s = s & "   , Home_Selection  " & vbCrLf
s = s & "   , PreSale_Selection  " & vbCrLf
s = s & "   , Sold  " & vbCrLf
s = s & "   , Last_DC  " & vbCrLf
s = s & "   , IntUID  " & vbCrLf
s = s & "   , IntPwd  " & vbCrLf
s = s & "   , Referral  " & vbCrLf
s = s & "   , Approved  " & vbCrLf
s = s & "   , Purchased  " & vbCrLf
s = s & "   , Purchased_Date  " & vbCrLf
s = s & "   , Unit_No  " & vbCrLf
s = s & "   , Legal_Unit_No  " & vbCrLf
s = s & "   , [Floor]  " & vbCrLf
s = s & "   --, DirectionID  " & vbCrLf
s = s & "   , LotNetTAX  " & vbCrLf
s = s & "   , ModelAssembly  " & vbCrLf
s = s & "   , ModelSalesWorksheet  " & vbCrLf
s = s & "   , ModelSpecDoc  " & vbCrLf
s = s & "   , ModelSpecDate  " & vbCrLf
s = s & "   , RealtorID  " & vbCrLf
s = s & "   , SellingAgent  " & vbCrLf
s = s & "   , CancelledDate  " & vbCrLf
s = s & "   , CreatedDate  " & vbCrLf
s = s & "   , ModifiedDate  " & vbCrLf
s = s & "   , PST  " & vbCrLf
s = s & "   , PSTRebate  " & vbCrLf
s = s & "   , GST  " & vbCrLf
s = s & "   , GST_Rebate  " & vbCrLf
s = s & "   , GST_Rate  " & vbCrLf
s = s & "   , PST_Rate  " & vbCrLf
s = s & "   , Base_House  " & vbCrLf
s = s & "   , Override_Price  " & vbCrLf
s = s & "   , Lot_Price  " & vbCrLf
s = s & "   , LotPremiumValue  " & vbCrLf
s = s & "   , LotPremiumPreTax  " & vbCrLf
s = s & "   , ScheduleB_total  " & vbCrLf
s = s & "   , CRMType  " & vbCrLf
s = s & "   , LeadStatus  " & vbCrLf
s = s & "   , Bal_Sheet_Prefix  " & vbCrLf
s = s & "   , Income_Prefix  " & vbCrLf
s = s & "   , Sold_To_Customer  " & vbCrLf
s = s & "   , Total_Sales_Price  " & vbCrLf
s = s & "   , Closed  " & vbCrLf
s = s & "   , Lot_City  " & vbCrLf
s = s & "   , Lot_Province  " & vbCrLf
s = s & "   , Lot_Zip  " & vbCrLf
s = s & "   , Law_Firm  " & vbCrLf
s = s & "   , Lending_Institution  " & vbCrLf
s = s & "" & vbCrLf
s = s & "   , ModelIncentive  " & vbCrLf
s = s & "   , Sales_Initiative  " & vbCrLf
s = s & "   , Sales_Deduction  " & vbCrLf
s = s & "   , TotalIncentives  " & vbCrLf
s = s & "   , Bank_Contact  " & vbCrLf
s = s & "   , Cost_Amount  " & vbCrLf
s = s & "   , co_PD_Total  " & vbCrLf
s = s & "   , CO_AP_Total  " & vbCrLf
s = s & "   , COApExempt  " & vbCrLf
s = s & "   , COPdExempt  " & vbCrLf
s = s & "   , TotalExempt  " & vbCrLf
s = s & "   , AddedToServiceModule  " & vbCrLf
s = s & "   , EstimateIndex " & vbCrLf
s = s & "   , CondoFees" & vbCrLf
s = s & "   , UnitFactor" & vbCrLf
s = s & "   , LOCKCO " & vbCrLf
s = s & "   , Parking1" & vbCrLf
s = s & "   , Parking1_Cost" & vbCrLf
s = s & "    , Parking1_Price" & vbCrLf
s = s & "    , Parking2" & vbCrLf
s = s & "    , Parking2_Cost" & vbCrLf
s = s & "    , Parking2_Price" & vbCrLf
s = s & "    , Parking3" & vbCrLf
s = s & "    , Parking3_Cost" & vbCrLf
s = s & "    , Parking3_Price" & vbCrLf
s = s & "       )  " & vbCrLf
s = s & "                SELECT  s.[QuoteID]  " & vbCrLf
s = s & "                      , s.Customer_No  " & vbCrLf
s = s & "                      , s.Community  " & vbCrLf
s = s & "                      , s.ClientDivisionID  " & vbCrLf
s = s & "                        " & vbCrLf
s = s & "       , s.BuyerFirstName  " & vbCrLf
s = s & "                      , s.BuyerLastName  " & vbCrLf
s = s & "       , s.Phone  " & vbCrLf
s = s & "       , s.Fax  " & vbCrLf
s = s & "       , s.Email  " & vbCrLf
s = s & "       , s.Address1  " & vbCrLf
s = s & "                      , s.Address2  " & vbCrLf
s = s & "                      , s.City  " & vbCrLf
s = s & "                      , s.[State]  " & vbCrLf
s = s & "                      , s.Zip  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                      , s.CoBuyerFirstName  " & vbCrLf
s = s & "                      , s.CoBuyerLastName  " & vbCrLf
s = s & "       , s.CoBuyerPhone  " & vbCrLf
s = s & "       , s.CoBuyerFax  " & vbCrLf
s = s & "       , s.CoBuyerEmail  " & vbCrLf
s = s & "       , s.CoBuyerAddress1  " & vbCrLf
s = s & "       , s.CoBuyerAddress2  " & vbCrLf
s = s & "       , s.CoBuyerCity  " & vbCrLf
s = s & "       , s.CoBuyerState  " & vbCrLf
s = s & "       , s.CoBuyerZip  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                      , s.Model  " & vbCrLf
s = s & "                      , s.Series  " & vbCrLf
s = s & "                      , s.SalesPerson  " & vbCrLf
s = s & "                      , s.[Description]  " & vbCrLf
s = s & "                      , s.[JobNo]  " & vbCrLf
s = s & "                      , s.[ExpectedOccupancy]  " & vbCrLf
s = s & "                      , s.[Lot]  " & vbCrLf
s = s & "                      , s.LotNo  " & vbCrLf
s = s & "                      , s.[Block]  " & vbCrLf
s = s & "                      , s.[Phase]  " & vbCrLf
s = s & "                      , s.[Plan]  " & vbCrLf
s = s & "                      , s.ApprovedBy  " & vbCrLf
s = s & "                 --     , s.[ApprovedDate]  " & vbCrLf
s = s & "                      , s.[IsContractSigned]  " & vbCrLf
s = s & "                      , s.[ContractSignedDate]  " & vbCrLf
s = s & "                      , s.ActualClosingDate--, s.[PossessionDate]  " & vbCrLf
s = s & "                      , s.[ConstructionStatus]  " & vbCrLf
s = s & "                      , s.[Comments]  " & vbCrLf
s = s & "                      , s.[MunicipalAddress]  " & vbCrLf
s = s & "                      , s.[County]  " & vbCrLf
s = s & "                      , s.[Township]  " & vbCrLf
s = s & "                      , s.[LegalAddress]  " & vbCrLf
s = s & "                      , s.[SquareFootage]  " & vbCrLf
s = s & "                      , s.[GSTOverride]  " & vbCrLf
s = s & "                      , s.[GSTRebateOverride]  " & vbCrLf
s = s & "                      , s.[LastCO]  " & vbCrLf
s = s & "                      , s.[LastAddChng]  " & vbCrLf
s = s & "                      , s.[IsCancelled]  " & vbCrLf
s = s & "                      , s.HomeSelection  " & vbCrLf
s = s & "                      , s.[PurchaseType]  " & vbCrLf
s = s & "                      , s.[IsSold]  " & vbCrLf
s = s & "                      , s.[LastDC]  " & vbCrLf
s = s & "                      , s.[IntUID]  " & vbCrLf
s = s & "                      , s.[IntPwd]  " & vbCrLf
s = s & "                      , s.[Referral]  " & vbCrLf
s = s & "                      , 0 --s.[Ratified]  Jeff wants approved set to 0 instead of using bldr approved" & vbCrLf
s = s & "                      , s.Purchased  " & vbCrLf
s = s & "                      , s.[PurchasedDate]  " & vbCrLf
s = s & "                      , s.[UnitNo]  " & vbCrLf
s = s & "                      , s.[LegalUnit]  " & vbCrLf
s = s & "                      , s.[Floor]  " & vbCrLf
s = s & "   --, s.DirectionID ??  " & vbCrLf
s = s & "                      , s.[LotNetTAX]  " & vbCrLf
s = s & "                      , s.[ModelAssembly]  " & vbCrLf
s = s & "                      , s.[ModelSalesWorksheet]  " & vbCrLf
s = s & "                      , s.[ModelSpecDoc]  " & vbCrLf
s = s & "                      , s.[ModelSpecDate]  " & vbCrLf
s = s & "                      , s.[RealtorID]  " & vbCrLf
s = s & "                      , s.[SellingAgent]  " & vbCrLf
s = s & "                      , s.[CancelledDate]  " & vbCrLf
s = s & "                      , s.[CreatedDate]  " & vbCrLf
s = s & "                      , s.[ModifiedDate]  " & vbCrLf
s = s & "                      , s.[PST]  " & vbCrLf
s = s & "                      , s.[PSTRebate]  " & vbCrLf
s = s & "                      , s.[GST]  " & vbCrLf
s = s & "                      , s.[GSTRebate]  " & vbCrLf
s = s & "                      , s.[GSTRate]  " & vbCrLf
s = s & "                      , s.[PSTRate]  " & vbCrLf
s = s & "                      , s.BaseHome  " & vbCrLf
s = s & "                      , s.BaseHomeSalesPrice  " & vbCrLf
s = s & "                      , s.LotPrice  " & vbCrLf
s = s & "                      , ISNULL(s.LotPremiumValue,0.00)  " & vbCrLf
s = s & "       , ISNULL(s.LotPremiumValue,0.00) as LotPremiumPreTax  " & vbCrLf
s = s & "                      , ISNULL(s.ContractAddendumTotal,0.00)  " & vbCrLf
s = s & "                      , s.CRMType  " & vbCrLf
s = s & "                      , 3  " & vbCrLf
s = s & "                      , CASE WHEN ISNULL(l.Bal_Sheet_Prefix, '') !='' THEN l.Bal_Sheet_Prefix ELSE c.Prefix1 END AS Prefix1  " & vbCrLf
s = s & "                      , c.Prefix2  " & vbCrLf
s = s & "                      , s.ClientInventoryHomeID  " & vbCrLf
s = s & "                      , s.TotalSalesPrice  " & vbCrLf
s = s & "       , s.Closed  " & vbCrLf
s = s & "       , s.LotCity  " & vbCrLf
s = s & "       , s.LotProvince  " & vbCrLf
s = s & "       , s.LotZip  " & vbCrLf
s = s & "       , s.LawFirm  " & vbCrLf
s = s & "       , s.LendingInstitution  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "       , s.ModelIncentive  " & vbCrLf
s = s & "       , s.SalesIncentive  " & vbCrLf
s = s & "       , s.SalesDeduction  " & vbCrLf
s = s & "       , s.TotalIncentives  " & vbCrLf
s = s & "       , s.LenderContactName  " & vbCrLf
s = s & "       , s.ModelCostAmount  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "       , s.[PendingCOTotal]  AS co_PD_Total  " & vbCrLf
s = s & "       , s.[ApprovedCOTotal]  AS CO_AP_Total  " & vbCrLf
s = s & "       --, s.[COTotal]  " & vbCrLf
s = s & "       , s.[ApprovedCOTotalExempt] COApExcempt  " & vbCrLf
s = s & "       , s.[PendingCOTotalExempt]  CoPdExempt  " & vbCrLf
s = s & "          , s.[COTotalExempt] TotalExempt  " & vbCrLf
s = s & "       , 1 AS AddedToServiceModule  " & vbCrLf
s = s & "       , isnull(s.EstimateIndex,0)" & vbCrLf
s = s & "      , s.CondoFees" & vbCrLf
s = s & "      , s.UnitFactor" & vbCrLf
s = s & "      , s.[COLocked] " & vbCrLf
s = s & "      , s.Parking1ID" & vbCrLf
s = s & "      , s.Parking1Cost" & vbCrLf
s = s & "      , s.Parking1SellingPrice" & vbCrLf
s = s & "      , s.Parking2ID" & vbCrLf
s = s & "      , s.Parking2Cost" & vbCrLf
s = s & "      , s.Parking2SellingPrice" & vbCrLf
s = s & "      , s.Parking3ID" & vbCrLf
s = s & "      , s.Parking3Cost" & vbCrLf
s = s & "      , s.Parking3SellingPrice" & vbCrLf
s = s & "                FROM    @tblSalesContract s  " & vbCrLf
s = s & "                        LEFT OUTER JOIN tblCustomers cust ON (cust.CRMID = s.QuoteID or cust.Customer_no  =s.Customer_No)  " & vbCrLf
s = s & "                        LEFT OUTER JOIN tblLocality c ON c.Area = s.Community  " & vbCrLf
s = s & "      LEFT OUTER JOIN tbllotinventory l on l.community = s.community AND l.Lot_No = s.LotNo  " & vbCrLf
s = s & "                WHERE   cust.CRMID IS NULL  " & vbCrLf
s = s & "                        AND s.CRMType = 2  " & vbCrLf
s = s & "                        AND s.IsCancelled = 0 " & vbCrLf
s = s & "  END " & vbCrLf
s = s & "  ELSE BEGIN" & vbCrLf
s = s & "  INSERT  INTO tblCustomers  " & vbCrLf
s = s & "                (CRMID  " & vbCrLf
s = s & "               , Customer_No  " & vbCrLf
s = s & "               , Community  " & vbCrLf
s = s & "               , DivisionID  " & vbCrLf
s = s & "                 " & vbCrLf
s = s & "      , Customer_Name  " & vbCrLf
s = s & "               , Customer_LName  " & vbCrLf
s = s & "      , Phone  " & vbCrLf
s = s & "      , Fax  " & vbCrLf
s = s & "      , Email  " & vbCrLf
s = s & "               , Address1  " & vbCrLf
s = s & "               , Address2  " & vbCrLf
s = s & "               , City  " & vbCrLf
s = s & "               , Province  " & vbCrLf
s = s & "               , Zip  " & vbCrLf
s = s & "                 " & vbCrLf
s = s & "      , CoBuyer_Name  " & vbCrLf
s = s & "               , CoBuyer_LName  " & vbCrLf
s = s & "      , C_Phone  " & vbCrLf
s = s & "      , C_Fax  " & vbCrLf
s = s & "      , C_Email   " & vbCrLf
s = s & "      , C_Address1  " & vbCrLf
s = s & "      , C_Address2  " & vbCrLf
s = s & "      , C_City  " & vbCrLf
s = s & "      , C_Province  " & vbCrLf
s = s & "               , C_Zip  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "               , Model  " & vbCrLf
s = s & "               , Series  " & vbCrLf
s = s & "               , Sales_Person_ID  " & vbCrLf
s = s & "               , [Description]  " & vbCrLf
s = s & "               , Job_No  " & vbCrLf
s = s & "               , Expected_Occupancy  " & vbCrLf
s = s & "               , Lot  " & vbCrLf
s = s & "               , Lot_No  " & vbCrLf
s = s & "               , Block  " & vbCrLf
s = s & "               , Phase  " & vbCrLf
s = s & "               , LOTPLAN  " & vbCrLf
s = s & "               , Approved_By  " & vbCrLf
s = s & "               , Approved_Date  " & vbCrLf
s = s & "               , Contract_Assigned  " & vbCrLf
s = s & "               , Contract_Signed_Date  " & vbCrLf
s = s & "               , Pos_Date  " & vbCrLf
s = s & "               , Construction_Status  " & vbCrLf
s = s & "               , Comments  " & vbCrLf
s = s & "               , Municipal_Address  " & vbCrLf
s = s & "               , County  " & vbCrLf
s = s & "               , Township  " & vbCrLf
s = s & "               , LegalAddress  " & vbCrLf
s = s & "               , Square_Footage  " & vbCrLf
s = s & "               , ChkGST  " & vbCrLf
s = s & "               , ChkGstRebate  " & vbCrLf
s = s & "               , Last_CO  " & vbCrLf
s = s & "               , Last_Add_Chng  " & vbCrLf
s = s & "               , Cancelled  " & vbCrLf
s = s & "               , Home_Selection  " & vbCrLf
s = s & "               , PreSale_Selection  " & vbCrLf
s = s & "               , Sold  " & vbCrLf
s = s & "               , Last_DC  " & vbCrLf
s = s & "               , IntUID  " & vbCrLf
s = s & "               , IntPwd  " & vbCrLf
s = s & "               , Referral  " & vbCrLf
s = s & "               , Approved  " & vbCrLf
s = s & "               , Purchased  " & vbCrLf
s = s & "               , Purchased_Date  " & vbCrLf
s = s & "               , Unit_No  " & vbCrLf
s = s & "               , Legal_Unit_No  " & vbCrLf
s = s & "               , [Floor]  " & vbCrLf
s = s & " --, DirectionID  " & vbCrLf
s = s & "               , LotNetTAX  " & vbCrLf
s = s & "               , ModelAssembly  " & vbCrLf
s = s & "               , ModelSalesWorksheet  " & vbCrLf
s = s & "               , ModelSpecDoc  " & vbCrLf
s = s & "               , ModelSpecDate  " & vbCrLf
s = s & "               , RealtorID  " & vbCrLf
s = s & "               , SellingAgent  " & vbCrLf
s = s & "               , CancelledDate  " & vbCrLf
s = s & "               , CreatedDate  " & vbCrLf
s = s & "               , ModifiedDate  " & vbCrLf
s = s & "               , PST  " & vbCrLf
s = s & "               , PSTRebate  " & vbCrLf
s = s & "               , GST  " & vbCrLf
s = s & "               , GST_Rebate  " & vbCrLf
s = s & "               , GST_Rate  " & vbCrLf
s = s & "               , PST_Rate  " & vbCrLf
s = s & "               , Base_House  " & vbCrLf
s = s & "               , Override_Price  " & vbCrLf
s = s & "               , Lot_Price  " & vbCrLf
s = s & "               , LotPremiumValue  " & vbCrLf
s = s & "      , LotPremiumPreTax  " & vbCrLf
s = s & "               , ScheduleB_total  " & vbCrLf
s = s & "               , CRMType  " & vbCrLf
s = s & "               , LeadStatus  " & vbCrLf
s = s & "               , Bal_Sheet_Prefix  " & vbCrLf
s = s & "               , Income_Prefix  " & vbCrLf
s = s & "               , Sold_To_Customer  " & vbCrLf
s = s & "               , Total_Sales_Price  " & vbCrLf
s = s & "      , Closed  " & vbCrLf
s = s & "      , Lot_City  " & vbCrLf
s = s & "      , Lot_Province  " & vbCrLf
s = s & "      , Lot_Zip  " & vbCrLf
s = s & "      , Law_Firm  " & vbCrLf
s = s & "      , Lending_Institution  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "      , ModelIncentive  " & vbCrLf
s = s & "      , Sales_Initiative  " & vbCrLf
s = s & "      , Sales_Deduction  " & vbCrLf
s = s & "      , TotalIncentives  " & vbCrLf
s = s & "      , Bank_Contact  " & vbCrLf
s = s & "      , Cost_Amount  " & vbCrLf
s = s & "        , co_PD_Total  " & vbCrLf
s = s & "      , CO_AP_Total  " & vbCrLf
s = s & "      , COApExempt  " & vbCrLf
s = s & "      , COPdExempt  " & vbCrLf
s = s & "         , TotalExempt  " & vbCrLf
s = s & "      , AddedToServiceModule  " & vbCrLf
s = s & "      , EstimateIndex " & vbCrLf
s = s & "     , CondoFees" & vbCrLf
s = s & "     , UnitFactor" & vbCrLf
s = s & "     , LOCKCO " & vbCrLf
s = s & "     , Parking1" & vbCrLf
s = s & "     , Parking1_Cost" & vbCrLf
s = s & "     , Parking1_Price" & vbCrLf
s = s & "     , Parking2" & vbCrLf
s = s & "     , Parking2_Cost" & vbCrLf
s = s & "     , Parking2_Price" & vbCrLf
s = s & "     , Parking3" & vbCrLf
s = s & "     , Parking3_Cost" & vbCrLf
s = s & "     , Parking3_Price" & vbCrLf
s = s & "       )  " & vbCrLf
s = s & "                SELECT  s.[QuoteID]  " & vbCrLf
s = s & "                      , s.Customer_No  " & vbCrLf
s = s & "                      , s.Community  " & vbCrLf
s = s & "                      , s.ClientDivisionID  " & vbCrLf
s = s & "                        " & vbCrLf
s = s & "       , s.BuyerFirstName  " & vbCrLf
s = s & "                      , s.BuyerLastName  " & vbCrLf
s = s & "       , s.Phone  " & vbCrLf
s = s & "       , s.Fax  " & vbCrLf
s = s & "       , s.Email  " & vbCrLf
s = s & "       , s.Address1  " & vbCrLf
s = s & "                      , s.Address2  " & vbCrLf
s = s & "                      , s.City  " & vbCrLf
s = s & "                      , s.[State]  " & vbCrLf
s = s & "                      , s.Zip  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                      , s.CoBuyerFirstName  " & vbCrLf
s = s & "                      , s.CoBuyerLastName  " & vbCrLf
s = s & "       , s.CoBuyerPhone  " & vbCrLf
s = s & "       , s.CoBuyerFax  " & vbCrLf
s = s & "       , s.CoBuyerEmail  " & vbCrLf
s = s & "       , s.CoBuyerAddress1  " & vbCrLf
s = s & "       , s.CoBuyerAddress2  " & vbCrLf
s = s & "       , s.CoBuyerCity  " & vbCrLf
s = s & "       , s.CoBuyerState  " & vbCrLf
s = s & "       , s.CoBuyerZip  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                      , s.Model  " & vbCrLf
s = s & "                      , s.Series  " & vbCrLf
s = s & "                      , s.SalesPerson  " & vbCrLf
s = s & "                      , s.[Description]  " & vbCrLf
s = s & "                      , s.[JobNo]  " & vbCrLf
s = s & "                      , s.[ExpectedOccupancy]  " & vbCrLf
s = s & "                      , s.[Lot]  " & vbCrLf
s = s & "                      , s.LotNo  " & vbCrLf
s = s & "                      , s.[Block]  " & vbCrLf
s = s & "                      , s.[Phase]  " & vbCrLf
s = s & "                      , s.[Plan]  " & vbCrLf
s = s & "                      , s.ApprovedBy  " & vbCrLf
s = s & "                      , s.[ApprovedDate]  " & vbCrLf
s = s & "                      , s.[IsContractSigned]  " & vbCrLf
s = s & "                      , s.[ContractSignedDate]  " & vbCrLf
s = s & "                      , s.ActualClosingDate--, s.[PossessionDate]  " & vbCrLf
s = s & "                      , s.[ConstructionStatus]  " & vbCrLf
s = s & "                      , s.[Comments]  " & vbCrLf
s = s & "                      , s.[MunicipalAddress]  " & vbCrLf
s = s & "                      , s.[County]  " & vbCrLf
s = s & "                      , s.[Township]  " & vbCrLf
s = s & "                      , s.[LegalAddress]  " & vbCrLf
s = s & "                      , s.[SquareFootage]  " & vbCrLf
s = s & "                      , s.[GSTOverride]  " & vbCrLf
s = s & "                      , s.[GSTRebateOverride]  " & vbCrLf
s = s & "                      , s.[LastCO]  " & vbCrLf
s = s & "                      , s.[LastAddChng]  " & vbCrLf
s = s & "                      , s.[IsCancelled]  " & vbCrLf
s = s & "                      , s.HomeSelection  " & vbCrLf
s = s & "                      , s.[PurchaseType]  " & vbCrLf
s = s & "                      , s.[IsSold]  " & vbCrLf
s = s & "                      , s.[LastDC]  " & vbCrLf
s = s & "                      , s.[IntUID]  " & vbCrLf
s = s & "                      , s.[IntPwd]  " & vbCrLf
s = s & "                      , s.[Referral]  " & vbCrLf
s = s & "                      , s.[Ratified]  " & vbCrLf
s = s & "                      , s.Purchased  " & vbCrLf
s = s & "                      , s.[PurchasedDate]  " & vbCrLf
s = s & "                      , s.[UnitNo]  " & vbCrLf
s = s & "                      , s.[LegalUnit]  " & vbCrLf
s = s & "                      , s.[Floor]  " & vbCrLf
s = s & "   --, s.DirectionID ??  " & vbCrLf
s = s & "                      , s.[LotNetTAX]  " & vbCrLf
s = s & "                      , s.[ModelAssembly]  " & vbCrLf
s = s & "                      , s.[ModelSalesWorksheet]  " & vbCrLf
s = s & "                      , s.[ModelSpecDoc]  " & vbCrLf
s = s & "                      , s.[ModelSpecDate]  " & vbCrLf
s = s & "                      , s.[RealtorID]  " & vbCrLf
s = s & "                      , s.[SellingAgent]  " & vbCrLf
s = s & "                      , s.[CancelledDate]  " & vbCrLf
s = s & "                      , s.[CreatedDate]  " & vbCrLf
s = s & "                      , s.[ModifiedDate]  " & vbCrLf
s = s & "                      , s.[PST]  " & vbCrLf
s = s & "                      , s.[PSTRebate]  " & vbCrLf
s = s & "                      , s.[GST]  " & vbCrLf
s = s & "                      , s.[GSTRebate]  " & vbCrLf
s = s & "                      , s.[GSTRate]  " & vbCrLf
s = s & "                      , s.[PSTRate]  " & vbCrLf
s = s & "                      , s.BaseHome  " & vbCrLf
s = s & "                      , s.BaseHomeSalesPrice  " & vbCrLf
s = s & "                      , s.LotPrice  " & vbCrLf
s = s & "                      , ISNULL(s.LotPremiumValue,0.00)  " & vbCrLf
s = s & "       , ISNULL(s.LotPremiumValue,0.00) as LotPremiumPreTax  " & vbCrLf
s = s & "                      , ISNULL(s.ContractAddendumTotal,0.00)  " & vbCrLf
s = s & "                      , s.CRMType  " & vbCrLf
s = s & "                      , 3  " & vbCrLf
s = s & "                      , CASE WHEN ISNULL(l.Bal_Sheet_Prefix, '') !='' THEN l.Bal_Sheet_Prefix ELSE c.Prefix1 END AS Prefix1  " & vbCrLf
s = s & "                      , c.Prefix2  " & vbCrLf
s = s & "                      , s.ClientInventoryHomeID  " & vbCrLf
s = s & "                      , s.TotalSalesPrice  " & vbCrLf
s = s & "       , s.Closed  " & vbCrLf
s = s & "       , s.LotCity  " & vbCrLf
s = s & "       , s.LotProvince  " & vbCrLf
s = s & "       , s.LotZip  " & vbCrLf
s = s & "       , s.LawFirm  " & vbCrLf
s = s & "       , s.LendingInstitution  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "       , s.ModelIncentive  " & vbCrLf
s = s & "       , s.SalesIncentive  " & vbCrLf
s = s & "       , s.SalesDeduction  " & vbCrLf
s = s & "       , s.TotalIncentives  " & vbCrLf
s = s & "       , s.LenderContactName  " & vbCrLf
s = s & "       , s.ModelCostAmount  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "       , s.[PendingCOTotal]  AS co_PD_Total  " & vbCrLf
s = s & "       , s.[ApprovedCOTotal]  AS CO_AP_Total  " & vbCrLf
s = s & "       --, s.[COTotal]  " & vbCrLf
s = s & "       , s.[ApprovedCOTotalExempt] COApExcempt  " & vbCrLf
s = s & "       , s.[PendingCOTotalExempt]  CoPdExempt  " & vbCrLf
s = s & "          , s.[COTotalExempt] TotalExempt  " & vbCrLf
s = s & "       , 1 AS AddedToServiceModule  " & vbCrLf
s = s & "       , isnull(s.EstimateIndex,0) " & vbCrLf
s = s & "      , s.CondoFees" & vbCrLf
s = s & "      , s.UnitFactor" & vbCrLf
s = s & "      , s.[COLocked]" & vbCrLf
s = s & "      , s.Parking1ID" & vbCrLf
s = s & "      , s.Parking1Cost" & vbCrLf
s = s & "      , s.Parking1SellingPrice" & vbCrLf
s = s & "      , s.Parking2ID" & vbCrLf
s = s & "      , s.Parking2Cost" & vbCrLf
s = s & "      , s.Parking2SellingPrice" & vbCrLf
s = s & "      , s.Parking3ID" & vbCrLf
s = s & "      , s.Parking3Cost" & vbCrLf
s = s & "      , s.Parking3SellingPrice    " & vbCrLf
s = s & "                FROM    @tblSalesContract s  " & vbCrLf
s = s & "                        LEFT OUTER JOIN tblCustomers cust ON (cust.CRMID = s.QuoteID or cust.Customer_no  =s.Customer_No)  " & vbCrLf
s = s & "                        LEFT OUTER JOIN tblLocality c ON c.Area = s.Community  " & vbCrLf
s = s & "      LEFT OUTER JOIN tbllotinventory l on l.community = s.community AND l.Lot_No = s.LotNo  " & vbCrLf
s = s & "                WHERE   cust.CRMID IS NULL  " & vbCrLf
s = s & "                        AND s.CRMType = 2  " & vbCrLf
s = s & "                        AND s.IsCancelled = 0 " & vbCrLf
s = s & "END" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "--Update Schedule B CRMID value for any options created in HF that were pushed from ClientDB to CRM " & vbCrLf
s = s & "" & vbCrLf
s = s & "Update CRM" & vbCrLf
s = s & "set CRMID = SCA.SalesContractAddendumID" & vbCrLf
s = s & "FROM    tblScheduleB CRM  " & vbCrLf
s = s & "                 INNER JOIN @tblSalesContractAddendum SCA ON (SCA.CRMSeq = CRM.seq AND SCA.Customer_no = CRM.Customer_No) " & vbCrLf
s = s & "               INNER JOIN tblCustomers c ON CRM.Customer_No = c.Customer_No " & vbCrLf
s = s & "                WHERE c.SendTOCRM = 1 " & vbCrLf
s = s & "                AND SCA.CRMType = 2  " & vbCrLf
s = s & "                AND CRM.CRMID is Null    " & vbCrLf
s = s & "                " & vbCrLf
s = s & "" & vbCrLf
s = s & " " & vbCrLf
s = s & "-- Do the options operations before modifying tblCustomers, otherwise the updates in the options will never happen.  " & vbCrLf
s = s & "-- We are using tblCustomer's ModifiedDate  " & vbCrLf
s = s & "" & vbCrLf
s = s & "set @sqlStatement = 'Insert ScheduleB from @tblSalesContractAddendum'" & vbCrLf
s = s & "" & vbCrLf
s = s & " INSERT  INTO tblScheduleB  " & vbCrLf
s = s & "                (CRMID  " & vbCrLf
s = s & "               , Customer_No  " & vbCrLf
s = s & "               , OPT  " & vbCrLf
s = s & "               , Option_Type  " & vbCrLf
s = s & "               , Custom  " & vbCrLf
s = s & "               , Require_Quote  " & vbCrLf
s = s & "               , Description  " & vbCrLf
s = s & "               , QTY  " & vbCrLf
s = s & "               , Cost_Amount  " & vbCrLf
s = s & "               , Rate  " & vbCrLf
s = s & "               , GrandTotal  " & vbCrLf
s = s & "               , Override_Price  " & vbCrLf
s = s & "               , comments  " & vbCrLf
s = s & "               , UOM  " & vbCrLf
s = s & "               , Major_Group -- This is the Category  " & vbCrLf
s = s & "              , Category  -- This is the subCategory  " & vbCrLf
s = s & "               , Assembly  " & vbCrLf
s = s & "               , SalesWorksheet  " & vbCrLf
s = s & "               , TAX  " & vbCrLf
s = s & "               , AddFromSpec  " & vbCrLf
s = s & "               , SpecDetailReference  " & vbCrLf
s = s & "               , EstimatorNotes  " & vbCrLf
s = s & "               , CONumber  " & vbCrLf
s = s & "               , Bldr_Declined  " & vbCrLf
s = s & "               , Bldr_Declined_Reason  " & vbCrLf
s = s & "               , Color  " & vbCrLf
s = s & "               , Location  " & vbCrLf
s = s & "               , Style  " & vbCrLf
s = s & "               , Finish  " & vbCrLf
s = s & "               , Other  " & vbCrLf
s = s & "               , CRMType  " & vbCrLf
s = s & "              , ApplyIncentive  " & vbCrLf
s = s & "              , EstimateIndex  " & vbCrLf
s = s & "              , InventoryHomeCRMID" & vbCrLf
s = s & "      )  " & vbCrLf
s = s & "                SELECT  sca.SalesContractAddendumID  " & vbCrLf
s = s & "                      , sca.Customer_no  " & vbCrLf
s = s & "                      , sca.[Option]  " & vbCrLf
s = s & "                      , sca.OptionType  " & vbCrLf
s = s & "                      , sca.IsCustom  " & vbCrLf
s = s & "                      , isnull(sca.RequireQuote,0)  " & vbCrLf
s = s & "                      , sca.[Description]  " & vbCrLf
s = s & "                      , sca.QTY  " & vbCrLf
s = s & "                      , sca.CostAmount  " & vbCrLf
s = s & "                      , sca.Rate  " & vbCrLf
s = s & "                      , isnull(sca.TotalPrice,0.00) TotalPrice  " & vbCrLf
s = s & "                      , CASE WHEN sca.OverridePrice = 0 THEN sca.TotalPrice ELSE sca.OverridePrice END  " & vbCrLf
s = s & "                      , sca.Comments  " & vbCrLf
s = s & "                      , sca.UOM  " & vbCrLf
s = s & "                      , sca.CategoryName  " & vbCrLf
s = s & "                      , sca.SubCategoryName  " & vbCrLf
s = s & "                      , sca.Assembly  " & vbCrLf
s = s & "                      , sca.SalesWorksheet  " & vbCrLf
s = s & "                      , isnull(sca.TaxAmount, 0.00) TaxAmount   " & vbCrLf
s = s & "     , sca.AddFromSpec  " & vbCrLf
s = s & "                      , sca.SpecDetailReference  " & vbCrLf
s = s & "                      , sca.EstimatorNotes  " & vbCrLf
s = s & "                      , sca.CONumber  " & vbCrLf
s = s & "                      , sca.IsBldrDeclined  " & vbCrLf
s = s & "                      , sca.BldrDeclinedReason  " & vbCrLf
s = s & "                      , sca.Color  " & vbCrLf
s = s & "                      , sca.Location  " & vbCrLf
s = s & "                      , sca.Style  " & vbCrLf
s = s & "                      , sca.Finish  " & vbCrLf
s = s & "                      , sca.Other  " & vbCrLf
s = s & "                      , sca.CRMType  " & vbCrLf
s = s & "                     , sca.ApplyIncentive  " & vbCrLf
s = s & "                     , sca.EstimateIndex" & vbCrLf
s = s & "                     , sca.InventoryHomeCRMID" & vbCrLf
s = s & "                FROM    @tblSalesContractAddendum sca  " & vbCrLf
s = s & "                        LEFT JOIN tblScheduleB sb ON   ((SCA.CRMSeq = sb.seq AND SCA.Customer_no = sb.Customer_No) OR (SCA.SalesContractAddendumID = sb.CRMID))  " & vbCrLf
s = s & "                WHERE   sb.CRMID IS NULL  " & vbCrLf
s = s & "                        AND sca.CRMType = 2  " & vbCrLf
s = s & "                       AND sca.IsOptDelete = 0  " & vbCrLf
s = s & "        " & vbCrLf
s = s & "       set @sqlStatement = 'Update ScheduleB from @tblSalesContractAddendum'" & vbCrLf
s = s & "-- Update for custom options previously downloaded.  " & vbCrLf
s = s & "        UPDATE  CRM  " & vbCrLf
s = s & "        SET     CRM.OPT = SCA.[Option]  " & vbCrLf
s = s & "              , CRM.Option_Type = SCA.OptionType  " & vbCrLf
s = s & "              , CRM.Custom = SCA.IsCustom  " & vbCrLf
s = s & "              , CRM.Require_Quote = isnull(SCA.RequireQuote,0)  " & vbCrLf
s = s & "              , CRM.Description = SCA.[Description]  " & vbCrLf
s = s & "              , CRM.QTY = SCA.QTY  " & vbCrLf
s = s & "              , CRM.Cost_Amount = SCA.CostAmount  " & vbCrLf
s = s & "              , CRM.Rate = SCA.Rate  " & vbCrLf
s = s & "              , CRM.Override_Price = CASE WHEN sca.OverridePrice = 0 THEN sca.TotalPrice ELSE sca.OverridePrice END  " & vbCrLf
s = s & "              , CRM.GrandTotal = SCA.TotalPrice  " & vbCrLf
s = s & "              , CRM.comments = SCA.Comments  " & vbCrLf
s = s & "              , CRM.UOM = SCA.UOM  " & vbCrLf
s = s & "              , CRM.Category =  SCA.SubCategoryName  " & vbCrLf
s = s & "              , CRM.Major_Group = SCA.CategoryName  -- Major Group is the Category  " & vbCrLf
s = s & "              , CRM.Assembly = SCA.Assembly  " & vbCrLf
s = s & "              , CRM.SalesWorksheet = SCA.SalesWorksheet  " & vbCrLf
s = s & "              , CRM.TAX = SCA.TaxAmount  " & vbCrLf
s = s & "              , CRM.AddFromSpec = SCA.AddFromSpec  " & vbCrLf
s = s & "              , CRM.SpecDetailReference = SCA.SpecDetailReference  " & vbCrLf
s = s & "              , CRM.EstimatorNotes = SCA.EstimatorNotes  " & vbCrLf
s = s & "              , CRM.CONumber = SCA.CONumber  " & vbCrLf
s = s & "              , CRM.Bldr_Declined = SCA.IsBldrDeclined  " & vbCrLf
s = s & "              , CRM.Bldr_Declined_Reason = SCA.BldrDeclinedReason  " & vbCrLf
s = s & "              , CRM.Color = SCA.Color  " & vbCrLf
s = s & "              , CRM.Location = SCA.Location  " & vbCrLf
s = s & "              , CRM.Style = SCA.Style  " & vbCrLf
s = s & "              , CRM.Finish = SCA.Finish  " & vbCrLf
s = s & "              , CRM.Other = SCA.Other  " & vbCrLf
s = s & "              , CRM.CRMType = SCA.CRMType  " & vbCrLf
s = s & "             , CRM.CRMID = SCA.SalesContractAddendumID  " & vbCrLf
s = s & "             , CRM.ApplyIncentive = SCA.ApplyIncentive  " & vbCrLf
s = s & "             , CRM.EstimateIndex = isnull(nullif(isnull(SCA.EstimateIndex,0),0),crm.EstimateIndex)" & vbCrLf
s = s & "        FROM    tblScheduleB CRM  " & vbCrLf
s = s & "                 INNER JOIN @tblSalesContractAddendum SCA ON (SCA.CRMSeq = CRM.seq AND SCA.Customer_no = CRM.Customer_No) OR (SCA.SalesContractAddendumID = CRM.CRMID)  " & vbCrLf
s = s & "               INNER JOIN tblCustomers c ON CRM.Customer_No = c.Customer_No  " & vbCrLf
s = s & "                WHERE  SCA.CRMType = 2  " & vbCrLf
s = s & "                AND CRM.CRMID is NOT Null    " & vbCrLf
s = s & "               -- AND CRM.ModifiedDate <= DATEADD(HOUR, @TimeZoneUTCOffSet, SCA.ModifiedDate)  " & vbCrLf
s = s & "       " & vbCrLf
s = s & "" & vbCrLf
s = s & "       DELETE b from tblScheduleB b" & vbCrLf
s = s & "       join @tblQuoteDetails a on a.QuoteDetailsID = b.CRMID" & vbCrLf
s = s & "       where a.IsOptDelete = 1 and a.QuoteDetailsID is not null" & vbCrLf
s = s & "" & vbCrLf
s = s & "       DELETE b from tblScheduleB b" & vbCrLf
s = s & "       join @tblSalesContractAddendum a on a.SalesContractAddendumID = b.CRMID" & vbCrLf
s = s & "       where a.IsOptDelete = 1 and a.SalesContractAddendumID is not null" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "----Update for existing options  " & vbCrLf
s = s & "--        UPDATE  CRM  " & vbCrLf
s = s & "--        SET     CRM.OPT = SCA.[Option]  " & vbCrLf
s = s & "--              , CRM.Option_Type = SCA.OptionType  " & vbCrLf
s = s & "--              , CRM.Custom = SCA.IsCustom  " & vbCrLf
s = s & "--              , CRM.Require_Quote = SCA.RequireQuote  " & vbCrLf
s = s & "--              , CRM.Description = SCA.[Description]  " & vbCrLf
s = s & "--              , CRM.QTY = SCA.QTY  " & vbCrLf
s = s & "--              , CRM.Cost_Amount = SCA.CostAmount  " & vbCrLf
s = s & "--              , CRM.Rate = SCA.Rate  " & vbCrLf
s = s & "--              , CRM.Override_Price = CASE WHEN sca.OverridePrice = 0 THEN sca.TotalPrice ELSE sca.OverridePrice END  " & vbCrLf
s = s & "--              , CRM.GrandTotal = SCA.TotalPrice  " & vbCrLf
s = s & "--              , CRM.comments = SCA.Comments  " & vbCrLf
s = s & "--              , CRM.UOM = SCA.UOM  " & vbCrLf
s = s & "--              , CRM.Category =  SCA.SubCategoryName  " & vbCrLf
s = s & "--              , CRM.Major_Group = SCA.CategoryName  -- Major Group is the Category  " & vbCrLf
s = s & "--              , CRM.Assembly = SCA.Assembly  " & vbCrLf
s = s & "--              , CRM.SalesWorksheet = SCA.SalesWorksheet  " & vbCrLf
s = s & "--              , CRM.TAX = SCA.TaxAmount  " & vbCrLf
s = s & "--              , CRM.AddFromSpec = SCA.AddFromSpec  " & vbCrLf
s = s & "--              , CRM.SpecDetailReference = SCA.SpecDetailReference  " & vbCrLf
s = s & "--              , CRM.EstimatorNotes = SCA.EstimatorNotes  " & vbCrLf
s = s & "--              , CRM.CONumber = SCA.CONumber  " & vbCrLf
s = s & "--              , CRM.Bldr_Declined = SCA.IsBldrDeclined  " & vbCrLf
s = s & "--              , CRM.Bldr_Declined_Reason = SCA.BldrDeclinedReason  " & vbCrLf
s = s & "--              , CRM.Color = SCA.Color  " & vbCrLf
s = s & "--              , CRM.Location = SCA.Location  " & vbCrLf
s = s & "--              , CRM.Style = SCA.Style  " & vbCrLf
s = s & "--              , CRM.Finish = SCA.Finish  " & vbCrLf
s = s & "--              , CRM.Other = SCA.Other  " & vbCrLf
s = s & "--              , CRM.CRMType = SCA.CRMType  " & vbCrLf
s = s & "--     , CRM.CRMID = SCA.SalesContractAddendumID  " & vbCrLf
s = s & "--        FROM    tblScheduleB CRM  " & vbCrLf
s = s & "--                 INNER JOIN @tblSalesContractAddendum SCA ON SCA.Customer_no = CRM.Customer_No AND SCA.SalesContractAddendumID = CRM.CRMID  " & vbCrLf
s = s & "--       WHERE  SCA.CRMType = 2  " & vbCrLf
s = s & "--     AND CRM.CRMID is NOT Null  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "-- Update CRMID on customer records there were converted and not yet assigned." & vbCrLf
s = s & "" & vbCrLf
s = s & "Update CRM" & vbCrLf
s = s & "set CRMID = s.QuoteID" & vbCrLf
s = s & ",CRMType =2" & vbCrLf
s = s & "FROM    tblCustomers CRM  " & vbCrLf
s = s & "           INNER JOIN @tblSalesContract s ON CRM.customer_no = s.ClientDBCustomerNo   " & vbCrLf
s = s & "           where isnull(s.ClientDBCustomerNo,'') <>'' and CRM.CRMID is null" & vbCrLf
s = s & "           AND s.CRMType = 2  AND s.IsCancelled = 0" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "declare @Reccount int" & vbCrLf
s = s & "" & vbCrLf
s = s & "  --set @Reccount = isnull((select count(*) from @tblSalesContract),0)" & vbCrLf
s = s & "  --if @Reccount>0 begin" & vbCrLf
s = s & "  -- select * into tmptblSalesContract from @tblSalesContract" & vbCrLf
s = s & "  --end" & vbCrLf
s = s & "" & vbCrLf
s = s & "                                                                                     " & vbCrLf
s = s & "--Insert Lot Premium record if premium is set" & vbCrLf
s = s & "set @sqlStatement = 'Insert LotPremium value if applicable from @tblSalesContract'  " & vbCrLf
s = s & "" & vbCrLf
s = s & "insert into LotPremium(Lot_No,PremiumID,Description,PremiumValue,PremiumOverride,UseOverride)" & vbCrLf
s = s & "select s.LotNo,'PREM','Premium',s.LotPremiumValue,s.LotPremiumValue,1 from @tblSalesContract s " & vbCrLf
s = s & "left outer join LotPremium p on p.Lot_No = s.LotNo" & vbCrLf
s = s & "where p.Lot_No is null and isnull(s.LotNo,'') <>'' and isnull(s.LotPremiumValue,0)<>0" & vbCrLf
s = s & "" & vbCrLf
s = s & "--Update Lot premium if premium is set " & vbCrLf
s = s & "" & vbCrLf
s = s & "set @sqlStatement = 'Update LotPremium value if applicable from @tblSalesContract'  " & vbCrLf
s = s & "" & vbCrLf
s = s & "update p" & vbCrLf
s = s & "set PremiumOverride=s.LotPremiumValue,UseOverride = 1" & vbCrLf
s = s & "from LotPremium p " & vbCrLf
s = s & "join @tblSalesContract s  on p.Lot_No = s.LotNo" & vbCrLf
s = s & "where p.Lot_No is null and isnull(s.LotNo,'') <>'' and isnull(s.LotPremiumValue,0)<>0" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "" & vbCrLf
s = s & " " & vbCrLf
s = s & "--" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  set @sqlStatement = 'Update tblCustomers from @tblSalesContract'      " & vbCrLf
s = s & "" & vbCrLf
s = s & " if (@IgnoreApprovedDate='True') BEGIN" & vbCrLf
s = s & " UPDATE  CRM  " & vbCrLf
s = s & "           SET     Community = s.Community  " & vbCrLf
s = s & "           , DivisionID = s.ClientDivisionID   " & vbCrLf
s = s & "           , Customer_Name = s.BuyerFirstName  " & vbCrLf
s = s & "           , Customer_LName = s.BuyerLastName  " & vbCrLf
s = s & "           , Phone =  s.Phone  " & vbCrLf
s = s & "           , Fax  =  s.Fax  " & vbCrLf
s = s & "           , Email =  s.Email  " & vbCrLf
s = s & "           , Address1 = s.Address1  " & vbCrLf
s = s & "           , Address2 = s.Address2  " & vbCrLf
s = s & "           , City = s.City  " & vbCrLf
s = s & "           , Province = s.[State]  " & vbCrLf
s = s & "           , Zip = s.Zip  " & vbCrLf
s = s & "           , CoBuyer_Name = s.CoBuyerFirstName  " & vbCrLf
s = s & "           , CoBuyer_LName = s.CoBuyerLastName  " & vbCrLf
s = s & "           , C_Phone = s.CoBuyerPhone  " & vbCrLf
s = s & "           , C_Fax = s.CoBuyerFax  " & vbCrLf
s = s & "           , C_Email = s.CoBuyerEmail  " & vbCrLf
s = s & "           , C_Address1 = s.CoBuyerAddress1  " & vbCrLf
s = s & "           , C_Address2 = s.CoBuyerAddress2  " & vbCrLf
s = s & "           , C_City  = s.CoBuyerCity  " & vbCrLf
s = s & "           , C_Province = s.CoBuyerState  " & vbCrLf
s = s & "           , C_Zip = s.CoBuyerZip  " & vbCrLf
s = s & "           , Model = s.Model  " & vbCrLf
s = s & "           , Series = s.Series  " & vbCrLf
s = s & "           , Sales_Person_ID = s.SalesPerson  " & vbCrLf
s = s & "           , [Description] = s.[Description]  " & vbCrLf
s = s & "           , Job_No =  CASE WHEN ISNULL(s.JobNo,'') = '' THEN CRM.Job_No ELSE s.[JobNo] END  " & vbCrLf
s = s & "           , Expected_Occupancy = s.[ExpectedOccupancy]  " & vbCrLf
s = s & "           , Lot_No = s.[LotNo]  " & vbCrLf
s = s & "           , Lot = s.Lot  " & vbCrLf
s = s & "           , Block = s.[Block]  " & vbCrLf
s = s & "           , Phase = s.[Phase]  " & vbCrLf
s = s & "           , LOTPLAN = s.[Plan]  " & vbCrLf
s = s & "           , Approved_By = s.ApprovedBy  " & vbCrLf
s = s & "           , Contract_Assigned = CASE WHEN s.[IsCancelled] = 0 THEN s.[IsContractSigned] ELSE 0 END  " & vbCrLf
s = s & "           , Contract_Signed_Date = CASE WHEN isnull(Contract_Signed_Date,'') = '' THEN  s.[ContractSignedDate] ELSE Contract_Signed_Date END  " & vbCrLf
s = s & "           , Pos_Date = CASE WHEN ISNULL(Pos_Date,'') = '' then s.[ActualClosingDate] ELSE Pos_Date END  " & vbCrLf
s = s & "            " & vbCrLf
s = s & "           , Comments = s.[Comments]  " & vbCrLf
s = s & "           , Municipal_Address = s.[MunicipalAddress]  " & vbCrLf
s = s & "           , County = s.[County]  " & vbCrLf
s = s & "           , Township = s.[Township]  " & vbCrLf
s = s & "           , LegalAddress = s.[LegalAddress]  " & vbCrLf
s = s & "           , Square_Footage = s.[SquareFootage]  " & vbCrLf
s = s & "           , ChkGST = s.[GSTOverride]  " & vbCrLf
s = s & "           , ChkGstRebate = s.[GSTRebateOverride]  " & vbCrLf
s = s & "           , Last_CO = s.[LastCO]  " & vbCrLf
s = s & "           , Last_Add_Chng = s.[LastAddChng]  " & vbCrLf
s = s & "           , Cancelled = s.[IsCancelled]  " & vbCrLf
s = s & "           , Home_Selection = s.HomeSelection  " & vbCrLf
s = s & "           , PreSale_Selection = s.[PurchaseType]  " & vbCrLf
s = s & "           , Sold = s.[IsSold]  " & vbCrLf
s = s & "           , Last_DC = s.[LastDC]  " & vbCrLf
s = s & "           , IntUID = s.[IntUID]  " & vbCrLf
s = s & "           , IntPwd = s.[IntPwd]  " & vbCrLf
s = s & "           , Referral = s.[Referral]  " & vbCrLf
s = s & "            " & vbCrLf
s = s & "           , Purchased = s.Purchased  " & vbCrLf
s = s & "           -- , Purchased_Date = IIF (Purchased_Date = NULL, s.[PurchasedDate], Purchased_Date)  " & vbCrLf
s = s & "           , Purchased_Date = CASE WHEN Purchased_Date = NULL THEN s.[PurchasedDate] ELSE Purchased_Date END  " & vbCrLf
s = s & "           , Unit_No = s.[UnitNo]  " & vbCrLf
s = s & "           , Legal_Unit_No = s.[LegalUnit]  " & vbCrLf
s = s & "           , [Floor] = s.[Floor]  " & vbCrLf
s = s & "           --, DirectionID   = -, s.DirectionID ??  " & vbCrLf
s = s & "           , LotNetTAX = s.[LotNetTAX]  " & vbCrLf
s = s & "           , ModelAssembly = s.[ModelAssembly]  " & vbCrLf
s = s & "           , ModelSalesWorksheet = s.[ModelSalesWorksheet]  " & vbCrLf
s = s & "           , ModelSpecDoc = s.[ModelSpecDoc]  " & vbCrLf
s = s & "           , ModelSpecDate = s.[ModelSpecDate]  " & vbCrLf
s = s & "           , RealtorID = s.[RealtorID]  " & vbCrLf
s = s & "           , SellingAgent = s.[SellingAgent]  " & vbCrLf
s = s & "           , CancelledDate = s.[CancelledDate]  " & vbCrLf
s = s & "           , CreatedDate = s.[CreatedDate]  " & vbCrLf
s = s & "           , ModifiedDate = s.[ModifiedDate]  " & vbCrLf
s = s & "           , PST = s.[PST]  " & vbCrLf
s = s & "           , PSTRebate = s.[PSTRebate]  " & vbCrLf
s = s & "           , GST = s.[GST]  " & vbCrLf
s = s & "           , GST_Rebate = s.[GSTRebate]  " & vbCrLf
s = s & "           , GST_Rate = s.[GSTRate]  " & vbCrLf
s = s & "           , PST_Rate = s.[PSTRate]  " & vbCrLf
s = s & "           , Base_House = s.BaseHome  " & vbCrLf
s = s & "           , Override_Price = s.BaseHomeSalesPrice  " & vbCrLf
s = s & "           , Lot_Price = s.LotPrice  " & vbCrLf
s = s & "           , LotPremiumValue =  ISNULL(s.LotPremiumValue,0.00)  " & vbCrLf
s = s & "           , LotPremiumPreTax =  ISNULL(s.LotPremiumValue,0.00)  " & vbCrLf
s = s & "           , ScheduleB_total = ISNULL(s.ContractAddendumTotal,0.00)  " & vbCrLf
s = s & "           , CRMType = s.CRMType  " & vbCrLf
s = s & "           , LeadStatus = 3  " & vbCrLf
s = s & "           --, Sold_To_Customer = s.ClientInventoryHomeID  " & vbCrLf
s = s & "           , Closed = s.Closed  " & vbCrLf
s = s & "           , Lot_City = s.LotCity  " & vbCrLf
s = s & "           , Lot_Province = s.LotProvince  " & vbCrLf
s = s & "           , Lot_Zip = s.LotZip  " & vbCrLf
s = s & "           , Law_Firm = s.LawFirm  " & vbCrLf
s = s & "           , Lending_Institution = s.LendingInstitution  " & vbCrLf
s = s & "           , ModelIncentive = s.ModelIncentive  " & vbCrLf
s = s & "           , Sales_Initiative = s.SalesIncentive  " & vbCrLf
s = s & "           , Sales_Deduction = s.SalesDeduction  " & vbCrLf
s = s & "           , TotalIncentives = s.TotalIncentives  " & vbCrLf
s = s & "           , Bank_Contact = s.LenderContactName  " & vbCrLf
s = s & "           , WO_Issued_Date = CASE WHEN s.[IsCancelled] = 0 THEN s.[ContractSignedDate] ELSE 0 END  " & vbCrLf
s = s & "           , CRM.Cost_Amount = s.ModelCostAmount " & vbCrLf
s = s & "           , LOCKCO = s.CoLocked" & vbCrLf
s = s & "           , CondoFees = s.CondoFees" & vbCrLf
s = s & "           , UnitFactor = s.UnitFactor " & vbCrLf
s = s & "           , Parking1 = s.Parking1ID" & vbCrLf
s = s & "             , Parking1_Cost = s.Parking1Cost" & vbCrLf
s = s & "             , Parking1_Price = s.Parking1SellingPrice" & vbCrLf
s = s & "             , Parking2 = s.Parking2ID" & vbCrLf
s = s & "             , Parking2_Cost = s.Parking2Cost" & vbCrLf
s = s & "             , Parking2_Price = s.Parking2SellingPrice" & vbCrLf
s = s & "             , Parking3 = s.Parking3ID" & vbCrLf
s = s & "             , Parking3_Cost = s.Parking3Cost" & vbCrLf
s = s & "             , Parking3_Price  = s.Parking3SellingPrice        " & vbCrLf
s = s & "             ,Bal_Sheet_Prefix=case when isnull(crm.Bal_Sheet_Prefix,'')<>'' then crm.Bal_Sheet_Prefix else c.Prefix1 end" & vbCrLf
s = s & "             ,Income_Prefix=case when isnull(crm.Income_Prefix,'')<>'' then crm.Income_Prefix else c.Prefix2 end" & vbCrLf
s = s & "           FROM    tblCustomers CRM  " & vbCrLf
s = s & "           INNER JOIN @tblSalesContract s ON CRM.CRMID = s.QuoteID   " & vbCrLf
s = s & "           join tbllocality c on c.Area = crm.Community" & vbCrLf
s = s & "           WHERE   s.CRMType = 2 " & vbCrLf
s = s & " END ELSE BEGIN" & vbCrLf
s = s & "        UPDATE  CRM  " & vbCrLf
s = s & "           SET     Community = s.Community  " & vbCrLf
s = s & "           , DivisionID = s.ClientDivisionID   " & vbCrLf
s = s & "           , Customer_Name = s.BuyerFirstName  " & vbCrLf
s = s & "           , Customer_LName = s.BuyerLastName  " & vbCrLf
s = s & "           , Phone =  s.Phone  " & vbCrLf
s = s & "           , Fax  =  s.Fax  " & vbCrLf
s = s & "           , Email =  s.Email  " & vbCrLf
s = s & "           , Address1 = s.Address1  " & vbCrLf
s = s & "           , Address2 = s.Address2  " & vbCrLf
s = s & "           , City = s.City  " & vbCrLf
s = s & "           , Province = s.[State]  " & vbCrLf
s = s & "           , Zip = s.Zip  " & vbCrLf
s = s & "           , CoBuyer_Name = s.CoBuyerFirstName  " & vbCrLf
s = s & "           , CoBuyer_LName = s.CoBuyerLastName  " & vbCrLf
s = s & "           , C_Phone = s.CoBuyerPhone  " & vbCrLf
s = s & "           , C_Fax = s.CoBuyerFax  " & vbCrLf
s = s & "           , C_Email = s.CoBuyerEmail  " & vbCrLf
s = s & "           , C_Address1 = s.CoBuyerAddress1  " & vbCrLf
s = s & "           , C_Address2 = s.CoBuyerAddress2  " & vbCrLf
s = s & "           , C_City  = s.CoBuyerCity  " & vbCrLf
s = s & "           , C_Province = s.CoBuyerState  " & vbCrLf
s = s & "           , C_Zip = s.CoBuyerZip  " & vbCrLf
s = s & "           , Model = s.Model  " & vbCrLf
s = s & "           , Series = s.Series  " & vbCrLf
s = s & "           , Sales_Person_ID = s.SalesPerson  " & vbCrLf
s = s & "           , [Description] = s.[Description]  " & vbCrLf
s = s & "           , Job_No =  CASE WHEN ISNULL(s.JobNo,'') = '' THEN CRM.Job_No ELSE s.[JobNo] END  " & vbCrLf
s = s & "           , Expected_Occupancy = s.[ExpectedOccupancy]  " & vbCrLf
s = s & "           , Lot_No = s.[LotNo]  " & vbCrLf
s = s & "           , Lot = s.Lot  " & vbCrLf
s = s & "           , Block = s.[Block]  " & vbCrLf
s = s & "           , Phase = s.[Phase]  " & vbCrLf
s = s & "           , LOTPLAN = s.[Plan]  " & vbCrLf
s = s & "           , Approved_By = s.ApprovedBy  " & vbCrLf
s = s & "           , Approved_Date = CASE WHEN s.[IsCancelled] = 0 THEN ( CASE WHEN Approved_Date is null then s.[ApprovedDate] ELSE Approved_Date END) ELSE NULL END  " & vbCrLf
s = s & "           , Contract_Assigned = CASE WHEN s.[IsCancelled] = 0 THEN s.[IsContractSigned] ELSE 0 END  " & vbCrLf
s = s & "           , Contract_Signed_Date = CASE WHEN isnull(Contract_Signed_Date,'') = '' THEN  s.[ContractSignedDate] ELSE Contract_Signed_Date END  " & vbCrLf
s = s & "           , Pos_Date = CASE WHEN ISNULL(Pos_Date,'') = '' then s.[ActualClosingDate] ELSE Pos_Date END  " & vbCrLf
s = s & "           --, Construction_Status = s.[ConstructionStatus]  " & vbCrLf
s = s & "           , Comments = s.[Comments]  " & vbCrLf
s = s & "           , Municipal_Address = s.[MunicipalAddress]  " & vbCrLf
s = s & "           , County = s.[County]  " & vbCrLf
s = s & "           , Township = s.[Township]  " & vbCrLf
s = s & "           , LegalAddress = s.[LegalAddress]  " & vbCrLf
s = s & "           , Square_Footage = s.[SquareFootage]  " & vbCrLf
s = s & "           , ChkGST = s.[GSTOverride]  " & vbCrLf
s = s & "           , ChkGstRebate = s.[GSTRebateOverride]  " & vbCrLf
s = s & "           , Last_CO = s.[LastCO]  " & vbCrLf
s = s & "           , Last_Add_Chng = s.[LastAddChng]  " & vbCrLf
s = s & "           , Cancelled = s.[IsCancelled]  " & vbCrLf
s = s & "           , Home_Selection = s.HomeSelection  " & vbCrLf
s = s & "           , PreSale_Selection = s.[PurchaseType]  " & vbCrLf
s = s & "           , Sold = s.[IsSold]  " & vbCrLf
s = s & "           , Last_DC = s.[LastDC]  " & vbCrLf
s = s & "           , IntUID = s.[IntUID]  " & vbCrLf
s = s & "           , IntPwd = s.[IntPwd]  " & vbCrLf
s = s & "           , Referral = s.[Referral]  " & vbCrLf
s = s & "           , Approved = CASE WHEN s.[IsCancelled] = 0 THEN s.[Ratified] ELSE 0 END  " & vbCrLf
s = s & "           , Purchased = s.Purchased  " & vbCrLf
s = s & "           -- , Purchased_Date = IIF (Purchased_Date = NULL, s.[PurchasedDate], Purchased_Date)  " & vbCrLf
s = s & "           , Purchased_Date = CASE WHEN Purchased_Date = NULL THEN s.[PurchasedDate] ELSE Purchased_Date END  " & vbCrLf
s = s & "           , Unit_No = s.[UnitNo]  " & vbCrLf
s = s & "           , Legal_Unit_No = s.[LegalUnit]  " & vbCrLf
s = s & "           , [Floor] = s.[Floor]  " & vbCrLf
s = s & "           --, DirectionID   = -, s.DirectionID ??  " & vbCrLf
s = s & "           , LotNetTAX = s.[LotNetTAX]  " & vbCrLf
s = s & "           , ModelAssembly = s.[ModelAssembly]  " & vbCrLf
s = s & "           , ModelSalesWorksheet = s.[ModelSalesWorksheet]  " & vbCrLf
s = s & "           , ModelSpecDoc = s.[ModelSpecDoc]  " & vbCrLf
s = s & "           , ModelSpecDate = s.[ModelSpecDate]  " & vbCrLf
s = s & "           , RealtorID = s.[RealtorID]  " & vbCrLf
s = s & "           , SellingAgent = s.[SellingAgent]  " & vbCrLf
s = s & "           , CancelledDate = s.[CancelledDate]  " & vbCrLf
s = s & "           , CreatedDate = s.[CreatedDate]  " & vbCrLf
s = s & "           , ModifiedDate = s.[ModifiedDate]  " & vbCrLf
s = s & "           , PST = s.[PST]  " & vbCrLf
s = s & "           , PSTRebate = s.[PSTRebate]  " & vbCrLf
s = s & "           , GST = s.[GST]  " & vbCrLf
s = s & "           , GST_Rebate = s.[GSTRebate]  " & vbCrLf
s = s & "           , GST_Rate = s.[GSTRate]  " & vbCrLf
s = s & "           , PST_Rate = s.[PSTRate]  " & vbCrLf
s = s & "           , Base_House = s.BaseHome  " & vbCrLf
s = s & "           , Override_Price = s.BaseHomeSalesPrice  " & vbCrLf
s = s & "           , Lot_Price = s.LotPrice  " & vbCrLf
s = s & "           , LotPremiumValue =  ISNULL(s.LotPremiumValue,0.00)  " & vbCrLf
s = s & "           , LotPremiumPreTax =  ISNULL(s.LotPremiumValue,0.00)  " & vbCrLf
s = s & "           , ScheduleB_total = ISNULL(s.ContractAddendumTotal,0.00)  " & vbCrLf
s = s & "           , CRMType = s.CRMType  " & vbCrLf
s = s & "           , LeadStatus = 3  " & vbCrLf
s = s & "           --, Sold_To_Customer = s.ClientInventoryHomeID  " & vbCrLf
s = s & "           , Closed = s.Closed  " & vbCrLf
s = s & "           , Lot_City = s.LotCity  " & vbCrLf
s = s & "           , Lot_Province = s.LotProvince  " & vbCrLf
s = s & "           , Lot_Zip = s.LotZip  " & vbCrLf
s = s & "           , Law_Firm = s.LawFirm  " & vbCrLf
s = s & "           , Lending_Institution = s.LendingInstitution  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "           , ModelIncentive = s.ModelIncentive  " & vbCrLf
s = s & "           , Sales_Initiative = s.SalesIncentive  " & vbCrLf
s = s & "           , Sales_Deduction = s.SalesDeduction  " & vbCrLf
s = s & "           , TotalIncentives = s.TotalIncentives  " & vbCrLf
s = s & "           , Bank_Contact = s.LenderContactName  " & vbCrLf
s = s & "           , WO_Issued_Date = CASE WHEN s.[IsCancelled] = 0 THEN s.[ContractSignedDate] ELSE 0 END  " & vbCrLf
s = s & "           , CRM.Cost_Amount = s.ModelCostAmount  " & vbCrLf
s = s & "           , LOCKCO = s.CoLocked" & vbCrLf
s = s & "           , CondoFees = s.CondoFees" & vbCrLf
s = s & "           , UnitFactor = s.UnitFactor" & vbCrLf
s = s & "           , Parking1 = s.Parking1ID" & vbCrLf
s = s & "             , Parking1_Cost = s.Parking1Cost" & vbCrLf
s = s & "             , Parking1_Price = s.Parking1SellingPrice" & vbCrLf
s = s & "             , Parking2 = s.Parking2ID" & vbCrLf
s = s & "             , Parking2_Cost = s.Parking2Cost" & vbCrLf
s = s & "             , Parking2_Price = s.Parking2SellingPrice" & vbCrLf
s = s & "             , Parking3 = s.Parking3ID" & vbCrLf
s = s & "             , Parking3_Cost = s.Parking3Cost" & vbCrLf
s = s & "             , Parking3_Price  = s.Parking3SellingPrice" & vbCrLf
s = s & "             ,Bal_Sheet_Prefix=case when isnull(crm.Bal_Sheet_Prefix,'')<>'' then crm.Bal_Sheet_Prefix else c.Prefix1 end" & vbCrLf
s = s & "             ,Income_Prefix=case when isnull(crm.Income_Prefix,'')<>'' then crm.Income_Prefix else c.Prefix2 end" & vbCrLf
s = s & "           FROM    tblCustomers CRM  " & vbCrLf
s = s & "           INNER JOIN @tblSalesContract s ON CRM.CRMID = s.QuoteID   " & vbCrLf
s = s & "           join tbllocality c on c.Area = crm.Community" & vbCrLf
s = s & "           WHERE   s.CRMType = 2 " & vbCrLf
s = s & "           --AND CRM.ModifiedDate <= DATEADD( HOUR, @TimeZoneUTCOffSet,s.ModifiedDate)  " & vbCrLf
s = s & "END" & vbCrLf
s = s & " " & vbCrLf
s = s & "--Update Estimateindex from Spec Record on customer for solds specs" & vbCrLf
s = s & " Update CRM" & vbCrLf
s = s & "set EstimateIndex = c.EstimateIndex" & vbCrLf
s = s & "FROM    tblCustomers CRM  " & vbCrLf
s = s & "           INNER JOIN @tblSalesContract s ON CRM.customer_no = s.Customer_No   " & vbCrLf
s = s & "           inner join tblCustomers c on c.Customer_No=s.ClientInventoryHomeID" & vbCrLf
s = s & "           where CRM.CRMID =s.QuoteID" & vbCrLf
s = s & "           AND s.CRMType = 2  AND isnull(s.ClientInventoryHomeID,'')<>'' and s.IsCancelled = 0" & vbCrLf
s = s & " " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  set @Reccount = 0" & vbCrLf
s = s & "          " & vbCrLf
s = s & "  ----set @Reccount = isnull((select count(*) from @tblQuotes),0)" & vbCrLf
s = s & "  ----if @Reccount>0 begin" & vbCrLf
s = s & "  ---- select * into tmptblQuotes from @tblQuotes" & vbCrLf
s = s & "  ----end" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "-- Insert into from Quotes Table.  " & vbCrLf
s = s & "-- This one inserts into the tblCustomer Table.  " & vbCrLf
s = s & "" & vbCrLf
s = s & "  set @sqlStatement = 'Insert tblCustomers from @tblQuotes'" & vbCrLf
s = s & "" & vbCrLf
s = s & "  INSERT  INTO tblCustomers  " & vbCrLf
s = s & "                (CRMID  " & vbCrLf
s = s & "               , Customer_No  " & vbCrLf
s = s & "               , Community  " & vbCrLf
s = s & "               , DivisionID  " & vbCrLf
s = s & "               , Customer_Name  " & vbCrLf
s = s & "               , Customer_LName  " & vbCrLf
s = s & "               , Address1  " & vbCrLf
s = s & "               , Address2  " & vbCrLf
s = s & "               , City  " & vbCrLf
s = s & "               , Province  " & vbCrLf
s = s & "               , Zip  " & vbCrLf
s = s & "               , CoBuyer_Name  " & vbCrLf
s = s & "               , CoBuyer_LName  " & vbCrLf
s = s & "               , Model  " & vbCrLf
s = s & "               , Series  " & vbCrLf
s = s & "               , Sales_Person_ID  " & vbCrLf
s = s & "               , [Description]  " & vbCrLf
s = s & "               , Job_No  " & vbCrLf
s = s & "               , Expected_Occupancy  " & vbCrLf
s = s & "               , Lot  " & vbCrLf
s = s & "               , Lot_No  " & vbCrLf
s = s & "               , Block  " & vbCrLf
s = s & "               , Phase  " & vbCrLf
s = s & "               , LOTPLAN  " & vbCrLf
s = s & "               , Approved_By  " & vbCrLf
s = s & "               , Approved_Date  " & vbCrLf
s = s & "               , Contract_Assigned  " & vbCrLf
s = s & "               , Contract_Signed_Date  " & vbCrLf
s = s & "               , Pos_Date  " & vbCrLf
s = s & "               , Construction_Status  " & vbCrLf
s = s & "               , Comments  " & vbCrLf
s = s & "               , Municipal_Address  " & vbCrLf
s = s & "               , County  " & vbCrLf
s = s & "               , Township  " & vbCrLf
s = s & "               , LegalAddress  " & vbCrLf
s = s & "               , Square_Footage  " & vbCrLf
s = s & "               , ChkGST  " & vbCrLf
s = s & "               , ChkGstRebate  " & vbCrLf
s = s & "               , Last_CO  " & vbCrLf
s = s & "               , Last_Add_Chng  " & vbCrLf
s = s & "               , Cancelled  " & vbCrLf
s = s & "               , Home_Selection  " & vbCrLf
s = s & "               , PreSale_Selection  " & vbCrLf
s = s & "               , Sold  " & vbCrLf
s = s & "               , Last_DC  " & vbCrLf
s = s & "               , IntUID  " & vbCrLf
s = s & "               , IntPwd  " & vbCrLf
s = s & "               , Referral  " & vbCrLf
s = s & "               , Approved  " & vbCrLf
s = s & "               , Purchased  " & vbCrLf
s = s & "               , Purchased_Date  " & vbCrLf
s = s & "               , Unit_No  " & vbCrLf
s = s & "               , Legal_Unit_No  " & vbCrLf
s = s & "               , [Floor]  " & vbCrLf
s = s & " --, DirectionID  " & vbCrLf
s = s & "               , LotNetTAX  " & vbCrLf
s = s & "               , ModelAssembly  " & vbCrLf
s = s & "               , ModelSalesWorksheet  " & vbCrLf
s = s & "               , ModelSpecDoc  " & vbCrLf
s = s & "               , ModelSpecDate  " & vbCrLf
s = s & "               , RealtorID  " & vbCrLf
s = s & "               , SellingAgent  " & vbCrLf
s = s & "               , CancelledDate  " & vbCrLf
s = s & "               , CreatedDate  " & vbCrLf
s = s & "               , ModifiedDate  " & vbCrLf
s = s & "               , PST  " & vbCrLf
s = s & "               , PSTRebate  " & vbCrLf
s = s & "               , GST  " & vbCrLf
s = s & "               , GST_Rebate  " & vbCrLf
s = s & "               , GST_Rate  " & vbCrLf
s = s & "               , PST_Rate  " & vbCrLf
s = s & "               , Base_House  " & vbCrLf
s = s & "               , Override_Price  " & vbCrLf
s = s & "               , Lot_Price  " & vbCrLf
s = s & "               , LotPremiumValue  " & vbCrLf
s = s & "      , LotPremiumPreTax  " & vbCrLf
s = s & "               , ScheduleB_total  " & vbCrLf
s = s & "               , CRMType  " & vbCrLf
s = s & "               , LeadStatus  " & vbCrLf
s = s & "               , Bal_Sheet_Prefix  " & vbCrLf
s = s & "               , Income_Prefix  " & vbCrLf
s = s & "               , Sold_To_Customer  " & vbCrLf
s = s & "      , Lot_City  " & vbCrLf
s = s & "      , Lot_Province  " & vbCrLf
s = s & "      , Lot_Zip  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "      , ModelIncentive  " & vbCrLf
s = s & "      , Sales_Initiative  " & vbCrLf
s = s & "      , Sales_Deduction  " & vbCrLf
s = s & "      , TotalIncentives  " & vbCrLf
s = s & "      , Cost_Amount  " & vbCrLf
s = s & "      , AddedToServiceModule)  " & vbCrLf
s = s & "                SELECT  q.[QuoteID]  " & vbCrLf
s = s & "                      , q.Customer_No  " & vbCrLf
s = s & "                      , q.Community  " & vbCrLf
s = s & "                      , q.ClientDivisionID  " & vbCrLf
s = s & "                      , q.BuyerFirstName  " & vbCrLf
s = s & "                      , q.BuyerLastName  " & vbCrLf
s = s & "                      , q.Address1  " & vbCrLf
s = s & "                      , q.Address2  " & vbCrLf
s = s & "                      , q.City  " & vbCrLf
s = s & "                      , q.[State]  " & vbCrLf
s = s & "                      , q.Zip  " & vbCrLf
s = s & "                      , q.CoBuyerFirstName  " & vbCrLf
s = s & "                      , q.CoBuyerLastName  " & vbCrLf
s = s & "                      , q.Model  " & vbCrLf
s = s & "                      , q.Series  " & vbCrLf
s = s & "                      , q.SalesPerson  " & vbCrLf
s = s & "                      , q.[Description]  " & vbCrLf
s = s & "                      , q.[JobNo] AS [JobNo]  --The sp_downloadtoCRM will send JobNo when is Pending Sale and enpty when is not.  " & vbCrLf
s = s & "                      , q.[ExpectedOccupancy]  " & vbCrLf
s = s & "                      , q.[Lot]  " & vbCrLf
s = s & "                      , q.[LotNo]  " & vbCrLf
s = s & "                      , q.[Block]  " & vbCrLf
s = s & "                      , q.[Phase]  " & vbCrLf
s = s & "                      , q.[Plan]  " & vbCrLf
s = s & "                      , q.ApprovedBy  " & vbCrLf
s = s & "                      , q.[ApprovedDate]  " & vbCrLf
s = s & "                      , q.[IsContractSigned]  " & vbCrLf
s = s & "                      , q.[ContractSignedDate]  " & vbCrLf
s = s & "                      , q.[PossessionDate]  " & vbCrLf
s = s & "                      , q.[ConstructionStatus]  " & vbCrLf
s = s & "                      , q.[Comments]  " & vbCrLf
s = s & "                      , q.[MunicipalAddress]  " & vbCrLf
s = s & "                      , q.[County]  " & vbCrLf
s = s & "                      , q.[Township]  " & vbCrLf
s = s & "                      , q.[LegalAddress]  " & vbCrLf
s = s & "                      , q.[SquareFootage]  " & vbCrLf
s = s & "                      , q.[GSTOverride]  " & vbCrLf
s = s & "                      , q.[GSTRebateOverride]  " & vbCrLf
s = s & "                      , q.[LastCO]  " & vbCrLf
s = s & "                      , q.[LastAddChng]  " & vbCrLf
s = s & "                      , q.[IsCancelled]  " & vbCrLf
s = s & "                      , q.HomeSelection  " & vbCrLf
s = s & "                      , q.[PurchaseType]  " & vbCrLf
s = s & "                      , q.[IsSold]  " & vbCrLf
s = s & "                      , q.[LastDC]  " & vbCrLf
s = s & "                      , q.[IntUID]  " & vbCrLf
s = s & "                      , q.[IntPwd]  " & vbCrLf
s = s & "                      , q.[Referral]  " & vbCrLf
s = s & "                      , q.[Ratified]  " & vbCrLf
s = s & "                      , q.Purchased  " & vbCrLf
s = s & "                      , q.[PurchasedDate]  " & vbCrLf
s = s & "                      , q.[UnitNo]  " & vbCrLf
s = s & "                      , q.[LegalUnit]  " & vbCrLf
s = s & "                      , q.[Floor]  " & vbCrLf
s = s & "   --, q.DirectionID ??  " & vbCrLf
s = s & "                      , q.[LotNetTAX]  " & vbCrLf
s = s & "                      , q.[ModelAssembly]  " & vbCrLf
s = s & "                      , q.[ModelSalesWorksheet]  " & vbCrLf
s = s & "                      , q.[ModelSpecDoc]  " & vbCrLf
s = s & "                      , q.[ModelSpecDate]  " & vbCrLf
s = s & "                      , q.[RealtorID]  " & vbCrLf
s = s & "                      , q.[SellingAgent]  " & vbCrLf
s = s & "                      , q.[CancelledDate]  " & vbCrLf
s = s & "                      , q.[CreatedDate]  " & vbCrLf
s = s & "                      , q.[ModifiedDate]  " & vbCrLf
s = s & "                      , q.[PST]  " & vbCrLf
s = s & "                      , q.[PSTRebate]  " & vbCrLf
s = s & "                      , q.[GST]  " & vbCrLf
s = s & "                      , q.[GSTRebate]  " & vbCrLf
s = s & "                      , q.[GSTRate]  " & vbCrLf
s = s & "                      , q.[PSTRate]  " & vbCrLf
s = s & "                      , q.BaseHome  " & vbCrLf
s = s & "                      , q.BaseHomeSalesPrice  " & vbCrLf
s = s & "                      , q.LotPrice  " & vbCrLf
s = s & "                      , ISNULL(q.LotPremiumValue,0.00)  " & vbCrLf
s = s & "       , ISNULL(q.LotPremiumValue,0.00)  " & vbCrLf
s = s & "                      , ISNULL(q.ContractAddendumTotal,0.00)  " & vbCrLf
s = s & "                      , q.CRMType  " & vbCrLf
s = s & "                      , 2  " & vbCrLf
s = s & "                      , (CASE WHEN ISNULL(l.Bal_Sheet_Prefix, '') !='' THEN l.Bal_Sheet_Prefix ELSE c.Prefix1 END) AS Prefix1  " & vbCrLf
s = s & "                      , c.Prefix2  " & vbCrLf
s = s & "                      , q.ClientInventoryHomeID  " & vbCrLf
s = s & "       , q.LotCity  " & vbCrLf
s = s & "       , q.LotProvince  " & vbCrLf
s = s & "       , q.LotZip  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "       , q.ModelIncentive  " & vbCrLf
s = s & "       , q.SalesIncentive  " & vbCrLf
s = s & "       , q.SalesDeduction  " & vbCrLf
s = s & "       , q.TotalIncentives  " & vbCrLf
s = s & "       , q.ModelCostAmount  " & vbCrLf
s = s & "       , (CASE WHEN ISNULL(q.ClientInventoryHomeID, '') = '' THEN 0 ELSE 1 END) AS AddedToServiceModule " & vbCrLf
s = s & "       " & vbCrLf
s = s & "                FROM    @tblQuotes q  " & vbCrLf
s = s & "                        LEFT OUTER JOIN tblCustomers cust ON (cust.CRMID = q.QuoteID or cust.Customer_No=q.Customer_No )" & vbCrLf
s = s & "                        LEFT OUTER JOIN tblLocality c ON c.Area = q.Community  " & vbCrLf
s = s & "      LEFT OUTER JOIN tbllotinventory l on l.community = q.community AND l.Lot_No = q.LotNo  " & vbCrLf
s = s & "                WHERE   cust.CRMID IS NULL  " & vbCrLf
s = s & "                        AND q.CRMType = 1  " & vbCrLf
s = s & "                        AND q.IsCancelled = 0 " & vbCrLf
s = s & "" & vbCrLf
s = s & "-- Do the options operations before modifying tblCustomers, otherwise the updates in the options will never happen.  " & vbCrLf
s = s & "-- We are using tblCustomer's ModifiedDate  " & vbCrLf
s = s & "" & vbCrLf
s = s & "  set @sqlStatement = 'Insert tblscheduleb from @tblQuoteDetails'" & vbCrLf
s = s & "" & vbCrLf
s = s & "INSERT  INTO tblScheduleB  " & vbCrLf
s = s & "                (CRMID  " & vbCrLf
s = s & "               , Customer_No  " & vbCrLf
s = s & "               , OPT  " & vbCrLf
s = s & "               , Option_Type  " & vbCrLf
s = s & "               , Custom  " & vbCrLf
s = s & "               , Require_Quote  " & vbCrLf
s = s & "               , Description  " & vbCrLf
s = s & "               , QTY  " & vbCrLf
s = s & "               , Cost_Amount  " & vbCrLf
s = s & "               , Rate  " & vbCrLf
s = s & "               , GrandTotal  " & vbCrLf
s = s & "               , Override_Price  " & vbCrLf
s = s & "               , comments  " & vbCrLf
s = s & "               , UOM  " & vbCrLf
s = s & "               , Major_Group -- This is the Category  " & vbCrLf
s = s & "               , Category  -- This is the SubCategory  " & vbCrLf
s = s & "              , Assembly  " & vbCrLf
s = s & "               , SalesWorksheet  " & vbCrLf
s = s & "               , TAX  " & vbCrLf
s = s & "               , AddFromSpec  " & vbCrLf
s = s & "               , SpecDetailReference  " & vbCrLf
s = s & "               , EstimatorNotes  " & vbCrLf
s = s & "               , CONumber  " & vbCrLf
s = s & "               , Bldr_Declined  " & vbCrLf
s = s & "               , Bldr_Declined_Reason  " & vbCrLf
s = s & "               , Color  " & vbCrLf
s = s & "               , Location  " & vbCrLf
s = s & "               , Style  " & vbCrLf
s = s & "               , Finish  " & vbCrLf
s = s & "               , Other  " & vbCrLf
s = s & "               , CRMType  " & vbCrLf
s = s & "              , ApplyIncentive" & vbCrLf
s = s & "              , InventoryHomeCRMID)  " & vbCrLf
s = s & "                SELECT  qd.QuoteDetailsID  " & vbCrLf
s = s & "                      , qd.Customer_no  " & vbCrLf
s = s & "                      , qd.[Option]  " & vbCrLf
s = s & "                      , qd.OptionType  " & vbCrLf
s = s & "                      , qd.IsCustom  " & vbCrLf
s = s & "                      , isnull(qd.RequireQuote,0)  " & vbCrLf
s = s & "                      , qd.[Description]  " & vbCrLf
s = s & "                      , qd.QTY  " & vbCrLf
s = s & "                  , isnull(qd.CostAmount,0.00) as CostAmount  " & vbCrLf
s = s & "                      , qd.Rate  " & vbCrLf
s = s & "                      , isnull(qd.TotalPrice,0.00) as TotalPrice  " & vbCrLf
s = s & "                      , CASE WHEN qd.OverridePrice = 0 THEN qd.TotalPrice ELSE qd.OverridePrice END  " & vbCrLf
s = s & "                      , qd.Comments  " & vbCrLf
s = s & "                      , qd.UOM  " & vbCrLf
s = s & "                      , qd.CategoryName  " & vbCrLf
s = s & "                      , qd.SubCategoryName  " & vbCrLf
s = s & "                      , qd.Assembly  " & vbCrLf
s = s & "                      , qd.SalesWorksheet  " & vbCrLf
s = s & "                      , qd.TaxAmount  " & vbCrLf
s = s & "                      , qd.AddFromSpec  " & vbCrLf
s = s & "                      , qd.SpecDetailReference  " & vbCrLf
s = s & "                      , qd.EstimatorNotes  " & vbCrLf
s = s & "                      , qd.CONumber  " & vbCrLf
s = s & "                      , qd.IsBldrDeclined  " & vbCrLf
s = s & "                      , qd.BldrDeclinedReason  " & vbCrLf
s = s & "                      , qd.Color  " & vbCrLf
s = s & "                      , qd.Location  " & vbCrLf
s = s & "                      , qd.Style  " & vbCrLf
s = s & "                      , qd.Finish  " & vbCrLf
s = s & "                      , qd.Other  " & vbCrLf
s = s & "                      , qd.CRMType  " & vbCrLf
s = s & "                     , qd.ApplyIncentive  " & vbCrLf
s = s & "                     , qd.InventoryHomeCRMID" & vbCrLf
s = s & "                FROM    @tblQuoteDetails qd  " & vbCrLf
s = s & "                        LEFT JOIN tblScheduleB sb ON qd.QuoteDetailsID = sb.CRMID  " & vbCrLf
s = s & "                       WHERE   sb.CRMID IS NULL  AND qd.CRMType = 1  AND qd.IsOptDelete = 0 " & vbCrLf
s = s & "                       and qd.requireQuote = 1" & vbCrLf
s = s & "" & vbCrLf
s = s & "    set @sqlStatement = 'Update tblscheduleb from @tblQuoteDetails'" & vbCrLf
s = s & "  " & vbCrLf
s = s & "        UPDATE  CRM  " & vbCrLf
s = s & "        SET     CRM.OPT = qd.[Option]  " & vbCrLf
s = s & "              , CRM.Option_Type = qd.OptionType  " & vbCrLf
s = s & "              , CRM.Custom = qd.IsCustom  " & vbCrLf
s = s & "              , CRM.Require_Quote = isnull(qd.RequireQuote,0)  " & vbCrLf
s = s & "              , CRM.Description = qd.[Description]  " & vbCrLf
s = s & "              , CRM.QTY = qd.QTY  " & vbCrLf
s = s & "              , CRM.Cost_Amount = qd.CostAmount  " & vbCrLf
s = s & "              , CRM.Rate = qd.Rate  " & vbCrLf
s = s & "              , CRM.GrandTotal = qd.TotalPrice  " & vbCrLf
s = s & "              , CRM.Override_Price = CASE WHEN qd.OverridePrice = 0 THEN qd.TotalPrice ELSE Qd.OverridePrice END  " & vbCrLf
s = s & "              , CRM.comments = qd.Comments  " & vbCrLf
s = s & "              , CRM.UOM = qd.UOM  " & vbCrLf
s = s & "              , CRM.Category = qd.SubCategoryName   " & vbCrLf
s = s & "              , CRM.Major_Group = qd.CategoryName -- Major Group is the Category  " & vbCrLf
s = s & "             -- , CRM.Assembly = qd.Assembly  " & vbCrLf
s = s & "              , CRM.SalesWorksheet = qd.SalesWorksheet  " & vbCrLf
s = s & "              , CRM.TAX = qd.TaxAmount  " & vbCrLf
s = s & "              , CRM.AddFromSpec = qd.AddFromSpec  " & vbCrLf
s = s & "              , CRM.SpecDetailReference = qd.SpecDetailReference  " & vbCrLf
s = s & "              , CRM.EstimatorNotes = qd.EstimatorNotes  " & vbCrLf
s = s & "              , CRM.CONumber = qd.CONumber  " & vbCrLf
s = s & "              , CRM.Bldr_Declined = qd.IsBldrDeclined  " & vbCrLf
s = s & "              , CRM.Bldr_Declined_Reason = qd.BldrDeclinedReason  " & vbCrLf
s = s & "              , CRM.Color = qd.Color  " & vbCrLf
s = s & "              , CRM.Location = qd.Location  " & vbCrLf
s = s & "              , CRM.Style = qd.Style  " & vbCrLf
s = s & "              , CRM.Finish = qd.Finish  " & vbCrLf
s = s & "              , CRM.Other = qd.Other  " & vbCrLf
s = s & "              , CRM.CRMType = qd.CRMType  " & vbCrLf
s = s & "             , CRM.ApplyIncentive = qd.ApplyIncentive   " & vbCrLf
s = s & "        FROM    tblScheduleB CRM  " & vbCrLf
s = s & "               INNER JOIN dbo.tblCustomers c ON CRM.Customer_No = c.Customer_No  " & vbCrLf
s = s & "                INNER JOIN @tblQuoteDetails qd ON qd.[QuoteDetailsID] = CRM.CRMID AND qd.CRMType = 1  " & vbCrLf
s = s & "               WHERE qd.requireQuote = 1 and c.Crmtype = 1 and qd.IsOptDelete = 0" & vbCrLf
s = s & "               --AND CRM.ModifiedDate  <= DATEADD( HOUR, @TimeZoneUTCOffSet,qd.LastUpdatedDate )   " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "" & vbCrLf
s = s & "         " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "    set @sqlStatement = 'Update tblCustomers from @tblQuotes'" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        UPDATE  CRM  " & vbCrLf
s = s & "        SET     Community = q.Community  " & vbCrLf
s = s & "              , DivisionID = q.ClientDivisionID  " & vbCrLf
s = s & "              , Customer_Name = q.BuyerFirstName  " & vbCrLf
s = s & "              , Customer_LName = q.BuyerLastName  " & vbCrLf
s = s & "              , Address1 = q.Address1  " & vbCrLf
s = s & "              , Address2 = q.Address2  " & vbCrLf
s = s & "              , City = q.City  " & vbCrLf
s = s & "              , Province = q.[State]  " & vbCrLf
s = s & "              , Zip = q.Zip  " & vbCrLf
s = s & "              , CoBuyer_Name = q.CoBuyerFirstName  " & vbCrLf
s = s & "              , CoBuyer_LName = q.CoBuyerLastName  " & vbCrLf
s = s & "              , Model = q.Model  " & vbCrLf
s = s & "              , Series = q.Series  " & vbCrLf
s = s & "              , Sales_Person_ID = q.SalesPerson  " & vbCrLf
s = s & "              , [Description] = q.[Description]  " & vbCrLf
s = s & "              , Job_No =  CASE WHEN ISNULL(q.JobNo,'') = '' THEN CRM.Job_No ELSE q.[JobNo] END    " & vbCrLf
s = s & "              , Expected_Occupancy = q.[ExpectedOccupancy]  " & vbCrLf
s = s & "              , Lot_No = q.[LotNo]  " & vbCrLf
s = s & "              , Lot = q.Lot  " & vbCrLf
s = s & "              , Block = q.[Block]  " & vbCrLf
s = s & "              , Phase = q.[Phase]  " & vbCrLf
s = s & "              , LOTPLAN = q.[Plan]  " & vbCrLf
s = s & "              , Approved_By = q.ApprovedBy  " & vbCrLf
s = s & "              , Approved_Date = q.[ApprovedDate]  " & vbCrLf
s = s & "              , Contract_Assigned = q.[IsContractSigned]  " & vbCrLf
s = s & "              , Contract_Signed_Date = q.[ContractSignedDate]  " & vbCrLf
s = s & "              , Pos_Date = q.[PossessionDate]  " & vbCrLf
s = s & "             -- , Construction_Status = q.[ConstructionStatus]  " & vbCrLf
s = s & "              , Comments = q.[Comments]  " & vbCrLf
s = s & "              , Municipal_Address = q.[MunicipalAddress]  " & vbCrLf
s = s & "              , County = q.[County]  " & vbCrLf
s = s & "              , Township = q.[Township]  " & vbCrLf
s = s & "              , LegalAddress = q.[LegalAddress]  " & vbCrLf
s = s & "              , Square_Footage = q.[SquareFootage]  " & vbCrLf
s = s & "              , ChkGST = q.[GSTOverride]  " & vbCrLf
s = s & "              , ChkGstRebate = q.[GSTRebateOverride]  " & vbCrLf
s = s & "              , Last_CO = q.[LastCO]  " & vbCrLf
s = s & "              , Last_Add_Chng = q.[LastAddChng]  " & vbCrLf
s = s & "              , Cancelled = q.[IsCancelled]  " & vbCrLf
s = s & "              , Home_Selection = q.HomeSelection  " & vbCrLf
s = s & "    , PreSale_Selection = q.[PurchaseType]  " & vbCrLf
s = s & "              , Sold = q.[IsSold]  " & vbCrLf
s = s & "              , Last_DC = q.[LastDC]  " & vbCrLf
s = s & "              , IntUID = q.[IntUID]  " & vbCrLf
s = s & "              , IntPwd = q.[IntPwd]  " & vbCrLf
s = s & "              , Referral = q.[Referral]  " & vbCrLf
s = s & "              , Approved = q.[Ratified]  " & vbCrLf
s = s & "              , Purchased = q.Purchased  " & vbCrLf
s = s & "              , Purchased_Date = q.[PurchasedDate]  " & vbCrLf
s = s & "              , Unit_No = q.[UnitNo]  " & vbCrLf
s = s & "              , Legal_Unit_No = q.[LegalUnit]  " & vbCrLf
s = s & "              , [Floor] = q.[Floor]  " & vbCrLf
s = s & " --, DirectionID   = -, q.DirectionID ??  " & vbCrLf
s = s & "              , LotNetTAX = q.[LotNetTAX]  " & vbCrLf
s = s & "              , ModelAssembly = q.[ModelAssembly]  " & vbCrLf
s = s & "              , ModelSalesWorksheet = q.[ModelSalesWorksheet]  " & vbCrLf
s = s & "              , ModelSpecDoc = q.[ModelSpecDoc]  " & vbCrLf
s = s & "              , ModelSpecDate = q.[ModelSpecDate]  " & vbCrLf
s = s & "              , RealtorID = q.[RealtorID]  " & vbCrLf
s = s & "              , SellingAgent = q.[SellingAgent]  " & vbCrLf
s = s & "              , CancelledDate = q.[CancelledDate]  " & vbCrLf
s = s & "              , CreatedDate = q.[CreatedDate]  " & vbCrLf
s = s & "              , ModifiedDate = q.[ModifiedDate]  " & vbCrLf
s = s & "              , PST = q.[PST]  " & vbCrLf
s = s & "              , PSTRebate = q.[PSTRebate]  " & vbCrLf
s = s & "              , GST = q.[GST]  " & vbCrLf
s = s & "              , GST_Rebate = q.[GSTRebate]  " & vbCrLf
s = s & "              , GST_Rate = q.[GSTRate]  " & vbCrLf
s = s & "              , PST_Rate = q.[PSTRate]  " & vbCrLf
s = s & "              , Base_House = q.BaseHome  " & vbCrLf
s = s & "              , Override_Price = q.BaseHomeSalesPrice  " & vbCrLf
s = s & "              , Lot_Price = q.LotPrice  " & vbCrLf
s = s & "              , LotPremiumValue =  isnull(q.LotPremiumValue,0.00)  " & vbCrLf
s = s & "     , LotPremiumPreTax = isnull(q.LotPremiumValue,0.00)  " & vbCrLf
s = s & "              , ScheduleB_total =  isnull(q.ContractAddendumTotal,0.00)  " & vbCrLf
s = s & "             -- , CRMType = q.CRMType  " & vbCrLf
s = s & "              , Sold_To_Customer = CASE WHEN q.IsCancelled = 0 THEN q.ClientInventoryHomeID ELSE NULL END  " & vbCrLf
s = s & "     , Lot_City = q.LotCity  " & vbCrLf
s = s & "     , Lot_Province = q.LotProvince  " & vbCrLf
s = s & "     , Lot_Zip = q.LotZip  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "     , ModelIncentive = q.ModelIncentive  " & vbCrLf
s = s & "     , Sales_Initiative = q.SalesIncentive  " & vbCrLf
s = s & "     , Sales_Deduction = q.SalesDeduction  " & vbCrLf
s = s & "     , TotalIncentives = q.TotalIncentives  " & vbCrLf
s = s & "     , Cost_Amount = q.ModelCostAmount  " & vbCrLf
s = s & "    , Bal_Sheet_Prefix=case when isnull(crm.Bal_Sheet_Prefix,'')<>'' then crm.Bal_Sheet_Prefix else c.Prefix1 end" & vbCrLf
s = s & "    , Income_Prefix=case when isnull(crm.Income_Prefix,'')<>'' then crm.Income_Prefix else c.Prefix2 end" & vbCrLf
s = s & "        FROM    tblCustomers CRM  " & vbCrLf
s = s & "                INNER JOIN @tblQuotes q ON CRM.CRMID = q.QuoteID   AND q.CRMType = 1  and crm.CRMType = 1 " & vbCrLf
s = s & "               join tbllocality c on c.Area = crm.Community" & vbCrLf
s = s & "  WHERE CRM.ModifiedDate <=  DATEADD( HOUR, @TimeZoneUTCOffSet,q.ModifiedDate )  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "      " & vbCrLf
s = s & "" & vbCrLf
s = s & "--Change Orders  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "    set @sqlStatement = 'Insert ChangeOrderMaster from @@tblChangeOrderMaster'" & vbCrLf
s = s & "" & vbCrLf
s = s & "        INSERT  INTO ChangeOrderMaster  " & vbCrLf
s = s & "                (Customer_No  " & vbCrLf
s = s & "               , Change_Order_No  " & vbCrLf
s = s & "               , Change_Date  " & vbCrLf
s = s & "               , Bldr_Approved_By  " & vbCrLf
s = s & "               , Bldr_Approved_Date  " & vbCrLf
s = s & "               , Cust_Approved_Date  " & vbCrLf
s = s & "               , Paid_Date  " & vbCrLf
s = s & "               , Next_Item_No  " & vbCrLf
s = s & "               , Cust_Approved  " & vbCrLf
s = s & "               , Bldr_Approved  " & vbCrLf
s = s & "               , Total_Price  " & vbCrLf
s = s & "               , Amount_Paid  " & vbCrLf
s = s & "               , Addition_Change  " & vbCrLf
s = s & "               , Declined  " & vbCrLf
s = s & "               , Declined_By  " & vbCrLf
s = s & "               , Notice_Estimating  " & vbCrLf
s = s & "               , Office_Change  " & vbCrLf
s = s & "               , Declined_Notice  " & vbCrLf
s = s & "               , UserID  " & vbCrLf
s = s & "               , NoCommTotal  " & vbCrLf
s = s & "               , ChargedFee  " & vbCrLf
s = s & "               , COLocked  " & vbCrLf
s = s & "               , CreationDate  " & vbCrLf
s = s & "               , RebateExempt  " & vbCrLf
s = s & "               , WebUpdated  " & vbCrLf
s = s & "               , NoDataChanged  " & vbCrLf
s = s & "               , Sales_Person_ID  " & vbCrLf
s = s & "               , contractexempt  " & vbCrLf
s = s & "               , EnerGuideRatingSum  " & vbCrLf
s = s & "               , GreenHouseRatingSum  " & vbCrLf
s = s & "               , FuelCostSum  " & vbCrLf
s = s & "               , REMOVEITEM  " & vbCrLf
s = s & "               , SumOfTotalAmount  " & vbCrLf
s = s & "               , SumOfNetTAX  " & vbCrLf
s = s & "               , NotTaxedTotal  " & vbCrLf
s = s & "               , ExportedToBMT  " & vbCrLf
s = s & "               , DataChangedBMTExport  " & vbCrLf
s = s & "               , EstimateIndex  " & vbCrLf
s = s & "               , ApplyIncentive  " & vbCrLf
s = s & "               , BillingReady  " & vbCrLf
s = s & "               , AmountBilled  " & vbCrLf
s = s & "               , Invoice  " & vbCrLf
s = s & "               , UseAltJCCodes  " & vbCrLf
s = s & "               , PurchasingCO  " & vbCrLf
s = s & "               , comments  " & vbCrLf
s = s & "               , Accounted_For  " & vbCrLf
s = s & "               , CRMID)  " & vbCrLf
s = s & "       SELECT  C.Customer_No  " & vbCrLf
s = s & "                      , C.Change_Order_No  " & vbCrLf
s = s & "                      , C.Change_Date  " & vbCrLf
s = s & "                      , C.Bldr_Approved_By  " & vbCrLf
s = s & "                      , C.Bldr_Approved_Date  " & vbCrLf
s = s & "                      , C.Cust_Approved_Date  " & vbCrLf
s = s & "                      , C.Paid_Date  " & vbCrLf
s = s & "                      , C.Next_Item_No  " & vbCrLf
s = s & "                      , C.Cust_Approved  " & vbCrLf
s = s & "                      , C.Bldr_Approved  " & vbCrLf
s = s & "                      --, 0" & vbCrLf
s = s & "       , C.Total_Price  " & vbCrLf
s = s & "                      , C.Amount_Paid  " & vbCrLf
s = s & "                      , C.Addition_Change  " & vbCrLf
s = s & "                      , C.Declined  " & vbCrLf
s = s & "                      , C.Declined_By  " & vbCrLf
s = s & "                      , C.Notice_Estimating  " & vbCrLf
s = s & "                      , C.Office_Change  " & vbCrLf
s = s & "                      , C.Declined_Notice  " & vbCrLf
s = s & "                      , C.UserID  " & vbCrLf
s = s & "                      , C.NoCommTotal  " & vbCrLf
s = s & "                      , C.ChargedFee  " & vbCrLf
s = s & "                      , C.COLocked  " & vbCrLf
s = s & "                      , C.CreationDate  " & vbCrLf
s = s & "                      , C.RebateExempt  " & vbCrLf
s = s & "                      , C.WebUpdated  " & vbCrLf
s = s & "                      , C.NoDataChanged  " & vbCrLf
s = s & "                      , C.Sales_Person_ID  " & vbCrLf
s = s & "                      , C.contractexempt  " & vbCrLf
s = s & "                      , C.EnerGuideRatingSum  " & vbCrLf
s = s & "                      , C.GreenHouseRatingSum  " & vbCrLf
s = s & "                      , C.FuelCostSum  " & vbCrLf
s = s & "                      , C.REMOVEITEM  " & vbCrLf
s = s & "                      , isnull(C.SumOfTotalAmount, 0.00)   " & vbCrLf
s = s & "                      , C.SumOfNetTAX  " & vbCrLf
s = s & "                      , C.NotTaxedTotal  " & vbCrLf
s = s & "                      , C.ExportedToBMT  " & vbCrLf
s = s & "                      , C.DataChangedBMTExport  " & vbCrLf
s = s & "                      , C.EstimateIndex  " & vbCrLf
s = s & "                      , 1 -- The value has to go hardcode for now -- C.ApplyIncentive  " & vbCrLf
s = s & "                      , C.BillingReady  " & vbCrLf
s = s & "                      , C.AmountBilled  " & vbCrLf
s = s & "                      , C.Invoice  " & vbCrLf
s = s & "                      , C.UseAltJCCodes  " & vbCrLf
s = s & "                      , C.PurchasingCO  " & vbCrLf
s = s & "                      , C.comments  " & vbCrLf
s = s & "                      , C.Accounted_For  " & vbCrLf
s = s & "                      , C.CRMID  " & vbCrLf
s = s & "                FROM    @tblChangeOrderMaster C  " & vbCrLf
s = s & "                       JOIN tblCustomers cust on cust.Customer_no = C.Customer_No" & vbCrLf
s = s & "                        LEFT JOIN ChangeOrderMaster CO ON (C.Customer_No = CO.Customer_No  " & vbCrLf
s = s & "                                                          AND C.Change_Order_No = CO.Change_Order_No)  " & vbCrLf
s = s & "     -- This has to be uncommented for any other client  " & vbCrLf
s = s & "    OR (C.CRMID = CO.CRMID)  " & vbCrLf
s = s & "                WHERE   CO.Customer_No IS NULL  " & vbCrLf
s = s & "" & vbCrLf
s = s & " " & vbCrLf
s = s & "     set @sqlStatement = 'Update ChangeOrderMaster from @@tblChangeOrderMaster' " & vbCrLf
s = s & "" & vbCrLf
s = s & "        UPDATE  CO  " & vbCrLf
s = s & "        SET     CO.Customer_No = C.Customer_No  " & vbCrLf
s = s & "              , CO.Change_Order_No = C.Change_Order_No  " & vbCrLf
s = s & "              , CO.Change_Date = C.Change_Date  " & vbCrLf
s = s & "              , CO.Bldr_Approved_By = C.Bldr_Approved_By  " & vbCrLf
s = s & "              , CO.Bldr_Approved_Date = C.Bldr_Approved_Date  " & vbCrLf
s = s & "              , CO.Cust_Approved_Date = C.Cust_Approved_Date  " & vbCrLf
s = s & "              , CO.Paid_Date = C.Paid_Date  " & vbCrLf
s = s & "              , CO.Next_Item_No = C.Next_Item_No  " & vbCrLf
s = s & "              , CO.Cust_Approved = C.Cust_Approved  " & vbCrLf
s = s & "              , CO.Bldr_Approved = C.Bldr_Approved  " & vbCrLf
s = s & "              , CO.Total_Price = C.Total_Price  " & vbCrLf
s = s & "              , CO.Amount_Paid = C.Amount_Paid  " & vbCrLf
s = s & "              , CO.Addition_Change = C.Addition_Change  " & vbCrLf
s = s & "              , CO.Declined = C.Declined  " & vbCrLf
s = s & "              , CO.Declined_By = C.Declined_By  " & vbCrLf
s = s & "              , CO.Notice_Estimating = C.Notice_Estimating  " & vbCrLf
s = s & "              , CO.Office_Change = C.Office_Change  " & vbCrLf
s = s & "              , CO.Declined_Notice = C.Declined_Notice  " & vbCrLf
s = s & "              , CO.UserID = C.UserID  " & vbCrLf
s = s & "              , CO.NoCommTotal = C.NoCommTotal  " & vbCrLf
s = s & "              , CO.ChargedFee = C.ChargedFee  " & vbCrLf
s = s & "              , CO.COLocked = C.COLocked  " & vbCrLf
s = s & "              , CO.CreationDate = C.CreationDate  " & vbCrLf
s = s & "              , CO.RebateExempt = C.RebateExempt  " & vbCrLf
s = s & "              , CO.WebUpdated = C.WebUpdated  " & vbCrLf
s = s & "              , CO.NoDataChanged = C.NoDataChanged  " & vbCrLf
s = s & "              , CO.Sales_Person_ID = C.Sales_Person_ID  " & vbCrLf
s = s & "              , CO.contractexempt = C.contractexempt  " & vbCrLf
s = s & "              , CO.EnerGuideRatingSum = C.EnerGuideRatingSum  " & vbCrLf
s = s & "              , CO.GreenHouseRatingSum = C.GreenHouseRatingSum  " & vbCrLf
s = s & "              , CO.FuelCostSum = C.FuelCostSum  " & vbCrLf
s = s & "              , CO.REMOVEITEM = C.REMOVEITEM  " & vbCrLf
s = s & "              , CO.SumOfTotalAmount = isnull(c.SumOfTotalAmount,0.00)  " & vbCrLf
s = s & "              , CO.SumOfNetTAX = C.SumOfNetTAX  " & vbCrLf
s = s & "              , CO.NotTaxedTotal = C.NotTaxedTotal  " & vbCrLf
s = s & "              , CO.ExportedToBMT = C.ExportedToBMT  " & vbCrLf
s = s & "              , CO.DataChangedBMTExport = C.DataChangedBMTExport  " & vbCrLf
s = s & "              , CO.ApplyIncentive = 1 -- the value is harcoded by now C.ApplyIncentive CCordoba  " & vbCrLf
s = s & "              , CO.BillingReady = C.BillingReady  " & vbCrLf
s = s & "              , CO.AmountBilled = C.AmountBilled  " & vbCrLf
s = s & "              , CO.Invoice = C.Invoice  " & vbCrLf
s = s & "              , CO.UseAltJCCodes = C.UseAltJCCodes  " & vbCrLf
s = s & "              , CO.PurchasingCO = C.PurchasingCO  " & vbCrLf
s = s & "              , CO.comments = C.comments  " & vbCrLf
s = s & "              , CO.Accounted_For = C.Accounted_For  " & vbCrLf
s = s & "              , CO.CRMID = C.CRMID  " & vbCrLf
s = s & "        FROM    @tblChangeOrderMaster C  " & vbCrLf
s = s & "                INNER JOIN ChangeOrderMaster CO ON (C.Customer_No = CO.Customer_No  " & vbCrLf
s = s & "                                                   AND C.Change_Order_No = CO.Change_Order_No)  " & vbCrLf
s = s & "     -- This has to be uncommented for any other client  " & vbCrLf
s = s & "     OR (C.CRMID = CO.CRMID )  " & vbCrLf
s = s & "  WHERE CO.TSTMP <= DATEADD( HOUR, @TimeZoneUTCOffSet,C.ModifiedDate )   " & vbCrLf
s = s & "  --AND (C.Bldr_Approved = 1   " & vbCrLf
s = s & "  AND (CO.Bldr_Approved = 0  or C.Declined =1)" & vbCrLf
s = s & "  --@LastPullDate OR CO.TSTMP IS NULL  " & vbCrLf
s = s & "     --NOTE: The data transfer should not update unless the HF.Bldr_approve is 0   " & vbCrLf
s = s & "     -- (None Approved) and the CRM is approved. If the client approves in HF then the DT would not update profit.     " & vbCrLf
s = s & "     --Daryl Email 04/20/2016  " & vbCrLf
s = s & "  --NOTE: Carlos C. Correction What happens if the CO is declined? Do we need to update that?   " & vbCrLf
s = s & "  --I removed the C.Bldr_Approve=1 to update the CO as long as the CO.Bldr_Approved is not 1  " & vbCrLf
s = s & "" & vbCrLf
s = s & "     " & vbCrLf
s = s & "   Update CRM" & vbCrLf
s = s & "set CRMID = SCA.SalesContractAddendumID" & vbCrLf
s = s & "FROM    tblScheduleB CRM  " & vbCrLf
s = s & "                 INNER JOIN @tblSalesContractAddendum SCA ON (SCA.CRMSeq = CRM.seq AND SCA.Customer_no = CRM.Customer_No) " & vbCrLf
s = s & "               INNER JOIN tblCustomers c ON CRM.Customer_No = c.Customer_No   " & vbCrLf
s = s & "                WHERE c.SendTOCRM = 1 " & vbCrLf
s = s & "                AND SCA.CRMType = 2  " & vbCrLf
s = s & "                AND CRM.CRMID is Null  " & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "    set @sqlStatement = 'Insert ChangeOrderDetails from @@tblChangeOrderDetails'" & vbCrLf
s = s & "" & vbCrLf
s = s & "        INSERT  INTO ChangeOrderDetails  " & vbCrLf
s = s & "                (Customer_No  " & vbCrLf
s = s & "               , Change_Order_No  " & vbCrLf
s = s & "               , Bldr_Approved_By  " & vbCrLf
s = s & "               , Bldr_Approved_Date  " & vbCrLf
s = s & "               , Paid_Date  " & vbCrLf
s = s & "               , Item_No  " & vbCrLf
s = s & "               , Custom  " & vbCrLf
s = s & "               , Structural_Change  " & vbCrLf
s = s & "               , Require_Quote  " & vbCrLf
s = s & "               , Description  " & vbCrLf
s = s & "               , QTY  " & vbCrLf
s = s & "               , Option_Type  " & vbCrLf
s = s & "               , Email_Sent  " & vbCrLf
s = s & "               , Cost_Amount  " & vbCrLf
s = s & "               , Base_price  " & vbCrLf
s = s & "               , Amount_Paid  " & vbCrLf
s = s & "               , Rate  " & vbCrLf
s = s & "               , Override_Price  " & vbCrLf
s = s & "               , TL_extra  " & vbCrLf
s = s & "               , Extra_Amount  " & vbCrLf
s = s & "               , gstrebate_exempt  " & vbCrLf
s = s & "               , UOM  " & vbCrLf
s = s & "               , OPT  " & vbCrLf
s = s & "               , Category  " & vbCrLf
s = s & "               , UserID  " & vbCrLf
s = s & "               , Notice_Estimating  " & vbCrLf
s = s & "               , NoCommission  " & vbCrLf
s = s & "               , Major_Group  -- This is the Category  " & vbCrLf
s = s & "               , CreationDate  " & vbCrLf
s = s & "               , UnSentEmail  " & vbCrLf
s = s & "               , item_seq  " & vbCrLf
s = s & "               , ref1  " & vbCrLf
s = s & "               , WebUpdated  " & vbCrLf
s = s & "               , NoDataChanged  " & vbCrLf
s = s & "               , PROCESSPO  " & vbCrLf
s = s & "               , POCREATED  " & vbCrLf
s = s & "               , OPTDELETED  " & vbCrLf
s = s & "               , EnerGuideRating  " & vbCrLf
s = s & "               , GreenHouseRating  " & vbCrLf
s = s & "               , FuelCost  " & vbCrLf
s = s & "               , REMOVEITEM  " & vbCrLf
s = s & "               , UseTax  " & vbCrLf
s = s & "               , NetTax  " & vbCrLf
s = s & "               , TotalAmount  " & vbCrLf
s = s & "               , GrandTotal  " & vbCrLf
s = s & "               , ExportedToBMT  " & vbCrLf
s = s & "               , DataChangedBMTExport  " & vbCrLf
s = s & "               , StoredFile  " & vbCrLf
s = s & "               , FileExtension  " & vbCrLf
s = s & "               , EstimateIndex  " & vbCrLf
s = s & "               , Assembly  " & vbCrLf
s = s & "               , SalesWorksheet  " & vbCrLf
s = s & "               , TAX  " & vbCrLf
s = s & "               , ApplyIncentive  " & vbCrLf
s = s & "               , IncentiveValue  " & vbCrLf
s = s & "               , Declined  " & vbCrLf
s = s & "               , EstimatorNotes  " & vbCrLf
s = s & "               , AddendumSeq  " & vbCrLf
s = s & "               , Bldr_Declined  " & vbCrLf
s = s & "               , Bldr_Declined_Reason  " & vbCrLf
s = s & "               , comments  " & vbCrLf
s = s & "               , Color  " & vbCrLf
s = s & "               , Location  " & vbCrLf
s = s & "               , SpecSeq  " & vbCrLf
s = s & "               , Style  " & vbCrLf
s = s & "               , Finish  " & vbCrLf
s = s & "               , Other  " & vbCrLf
s = s & "               , ColorListID  " & vbCrLf
s = s & "               , StyleListID  " & vbCrLf
s = s & "               , FinishListID  " & vbCrLf
s = s & "         , OtherListID  " & vbCrLf
s = s & "               , quote_entered_date  " & vbCrLf
s = s & "               , CRMID)  " & vbCrLf
s = s & "                SELECT  D.Customer_No  " & vbCrLf
s = s & "                      , D.Change_Order_No  " & vbCrLf
s = s & "                      , D.Bldr_Approved_By  " & vbCrLf
s = s & "                      , D.Bldr_Approved_Date  " & vbCrLf
s = s & "                      , D.Paid_Date  " & vbCrLf
s = s & "                      , D.Item_No  " & vbCrLf
s = s & "                      , D.Custom  " & vbCrLf
s = s & "                      , D.Structural_Change  " & vbCrLf
s = s & "                      , D.Require_Quote  " & vbCrLf
s = s & "                      , D.Description  " & vbCrLf
s = s & "                      , D.QTY  " & vbCrLf
s = s & "                      , D.Option_Type  " & vbCrLf
s = s & "                      , D.Email_Sent  " & vbCrLf
s = s & "                      , D.Cost_Amount  " & vbCrLf
s = s & "                      , D.Base_price  " & vbCrLf
s = s & "                      , D.Amount_Paid  " & vbCrLf
s = s & "                      , D.Rate  " & vbCrLf
s = s & "                      , D.Override_Price  " & vbCrLf
s = s & "                      , D.TL_extra  " & vbCrLf
s = s & "                      , D.Extra_Amount  " & vbCrLf
s = s & "                      , isnull(D.gstrebate_exempt,0)  " & vbCrLf
s = s & "                      , D.UOM  " & vbCrLf
s = s & "                      , D.OPT  " & vbCrLf
s = s & "                      , D.Category  " & vbCrLf

    DT_Set_DataFromCRM_a = s
End Function
Private Function DT_Set_DataFromCRM_b() As String
    Dim s As String

s = ""
s = s & "                      , D.UserID  " & vbCrLf
s = s & "                      , D.Notice_Estimating  " & vbCrLf
s = s & "                      , D.NoCommission  " & vbCrLf
s = s & "                      , D.Major_Group -- This is the Category  " & vbCrLf
s = s & "                      , D.CreationDate  " & vbCrLf
s = s & "                      , D.UnSentEmail  " & vbCrLf
s = s & "                      , D.item_seq  " & vbCrLf
s = s & "                      , D.ref1  " & vbCrLf
s = s & "                      , D.WebUpdated  " & vbCrLf
s = s & "                      , D.NoDataChanged  " & vbCrLf
s = s & "                      , D.PROCESSPO  " & vbCrLf
s = s & "                      , D.POCREATED  " & vbCrLf
s = s & "                      , D.OPTDELETED  " & vbCrLf
s = s & "                      , D.EnerGuideRating  " & vbCrLf
s = s & "                      , D.GreenHouseRating  " & vbCrLf
s = s & "                      , D.FuelCost  " & vbCrLf
s = s & "                      , D.REMOVEITEM  " & vbCrLf
s = s & "                      , D.UseTax  " & vbCrLf
s = s & "                      , D.NetTax  " & vbCrLf
s = s & "                      , isnull(D.TotalAmount,0.00) as TotalAmount  " & vbCrLf
s = s & "                      , D.GrandTotal  " & vbCrLf
s = s & "                      , D.ExportedToBMT  " & vbCrLf
s = s & "                      , D.DataChangedBMTExport  " & vbCrLf
s = s & "                      , D.StoredFile  " & vbCrLf
s = s & "                      , D.FileExtension  " & vbCrLf
s = s & "                      , D.EstimateIndex  " & vbCrLf
s = s & "                      , D.Assembly  " & vbCrLf
s = s & "                      , D.SalesWorksheet  " & vbCrLf
s = s & "                      , D.TAX  " & vbCrLf
s = s & "                      , 1   --, D.ApplyIncentive The value has be be harcoded for now May 03 2016 CCordoba  " & vbCrLf
s = s & "                      , D.IncentiveValue  " & vbCrLf
s = s & "                      , D.Declined  " & vbCrLf
s = s & "                      , D.EstimatorNotes  " & vbCrLf
s = s & "                      , D.AddendumSeq  " & vbCrLf
s = s & "                      , D.Bldr_Declined  " & vbCrLf
s = s & "                      , D.Bldr_Declined_Reason  " & vbCrLf
s = s & "                      , D.comments  " & vbCrLf
s = s & "                      , D.Color  " & vbCrLf
s = s & "                      , D.Location  " & vbCrLf
s = s & "                      , D.SpecSeq  " & vbCrLf
s = s & "                      , D.Style  " & vbCrLf
s = s & "                      , D.Finish  " & vbCrLf
s = s & "                      , D.Other  " & vbCrLf
s = s & "                      , D.ColorListID  " & vbCrLf
s = s & "                      , D.StyleListID  " & vbCrLf
s = s & "                      , D.FinishListID  " & vbCrLf
s = s & "                      , D.OtherListID  " & vbCrLf
s = s & "                      , D.quote_entered_date  " & vbCrLf
s = s & "                      , D.CRMID  " & vbCrLf
s = s & "              FROM    @tblChangeOrderDetails D  " & vbCrLf
s = s & "                   INNER JOIN ChangeOrderMaster com ON D.Customer_No=COM.Customer_No   AND D.Change_Order_No=COM.Change_Order_No  " & vbCrLf
s = s & "                    LEFT JOIN ChangeOrderDetails COD ON ((D.Customer_No = COD.Customer_No AND D.ClientSeq = COD.seq AND D.Change_Order_No = COD.Change_Order_No) OR D.CRMID = COD.CRMID)  " & vbCrLf
s = s & "                   WHERE COD.seq IS NULL AND  ISNULL(D.OPTDELETED,0) = 0  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "      set @sqlStatement = 'Update ChangeOrderDetails from @@tblChangeOrderDetails' " & vbCrLf
s = s & "    " & vbCrLf
s = s & "        UPDATE  COD  " & vbCrLf
s = s & "        SET     COD.Customer_No = D.Customer_No  " & vbCrLf
s = s & "              , COD.Change_Order_No = D.Change_Order_No  " & vbCrLf
s = s & "              , COD.Bldr_Approved_By = D.Bldr_Approved_By  " & vbCrLf
s = s & "              , COD.Bldr_Approved_Date = D.Bldr_Approved_Date  " & vbCrLf
s = s & "              , COD.Paid_Date = D.Paid_Date  " & vbCrLf
s = s & "              , COD.Item_No = D.Item_No  " & vbCrLf
s = s & "              , COD.Custom = D.Custom  " & vbCrLf
s = s & "              , COD.Structural_Change = D.Structural_Change  " & vbCrLf
s = s & "              , COD.Require_Quote = D.Require_Quote  " & vbCrLf
s = s & "              , COD.Description = D.Description  " & vbCrLf
s = s & "              , COD.QTY = D.QTY  " & vbCrLf
s = s & "              , COD.Option_Type = D.Option_Type  " & vbCrLf
s = s & "              , COD.Email_Sent = D.Email_Sent  " & vbCrLf
s = s & "              , COD.Cost_Amount = D.Cost_Amount  " & vbCrLf
s = s & "              , COD.Base_price = D.Base_price  " & vbCrLf
s = s & "              , COD.Amount_Paid = D.Amount_Paid  " & vbCrLf
s = s & "              , COD.Rate = D.Rate  " & vbCrLf
s = s & "              , COD.Override_Price = D.Override_Price  " & vbCrLf
s = s & "              , COD.TL_extra = D.TL_extra  " & vbCrLf
s = s & "              , COD.Extra_Amount = D.Extra_Amount  " & vbCrLf
s = s & "              , COD.gstrebate_exempt = isnull(D.gstrebate_exempt  ,0)" & vbCrLf
s = s & "              , COD.UOM = D.UOM  " & vbCrLf
s = s & "              , COD.OPT = D.OPT  " & vbCrLf
s = s & "              , COD.Category = D.Category  " & vbCrLf
s = s & "              , COD.UserID = D.UserID  " & vbCrLf
s = s & "              , COD.Notice_Estimating = D.Notice_Estimating  " & vbCrLf
s = s & "              , COD.NoCommission = D.NoCommission  " & vbCrLf
s = s & "              , COD.Major_Group = D.Major_Group  " & vbCrLf
s = s & "              , COD.CreationDate = D.CreationDate  " & vbCrLf
s = s & "              , COD.UnSentEmail = D.UnSentEmail  " & vbCrLf
s = s & "              , COD.item_seq = D.item_seq  " & vbCrLf
s = s & "              , COD.ref1 = D.ref1  " & vbCrLf
s = s & "              , COD.WebUpdated = D.WebUpdated  " & vbCrLf
s = s & "              , COD.NoDataChanged = D.NoDataChanged  " & vbCrLf
s = s & "              , COD.PROCESSPO = D.PROCESSPO  " & vbCrLf
s = s & "              , COD.POCREATED = D.POCREATED   " & vbCrLf
s = s & "              , COD.EnerGuideRating = D.EnerGuideRating  " & vbCrLf
s = s & "              , COD.GreenHouseRating = D.GreenHouseRating  " & vbCrLf
s = s & "              , COD.FuelCost = D.FuelCost  " & vbCrLf
s = s & "              , COD.REMOVEITEM = D.REMOVEITEM  " & vbCrLf
s = s & "              , COD.UseTax = D.UseTax  " & vbCrLf
s = s & "              , COD.NetTax = D.NetTax  " & vbCrLf
s = s & "              , COD.TotalAmount = isnull( D.TotalAmount, 0.00)   " & vbCrLf
s = s & "              , COD.GrandTotal = D.GrandTotal  " & vbCrLf
s = s & "              , COD.ExportedToBMT = D.ExportedToBMT  " & vbCrLf
s = s & "              , COD.DataChangedBMTExport = D.DataChangedBMTExport  " & vbCrLf
s = s & "              , COD.StoredFile = D.StoredFile  " & vbCrLf
s = s & "              , COD.FileExtension = D.FileExtension   " & vbCrLf
s = s & "              , COD.Assembly = D.Assembly  " & vbCrLf
s = s & "              , COD.SalesWorksheet = D.SalesWorksheet  " & vbCrLf
s = s & "              , COD.TAX = D.TAX  " & vbCrLf
s = s & "              , COD.ApplyIncentive = 1 --D.ApplyIncentive the value is Harcoded by now.   " & vbCrLf
s = s & "              , COD.IncentiveValue = D.IncentiveValue  " & vbCrLf
s = s & "              , COD.Declined = D.Declined  " & vbCrLf
s = s & "              , COD.EstimatorNotes = D.EstimatorNotes  " & vbCrLf
s = s & "              , COD.AddendumSeq = D.AddendumSeq  " & vbCrLf
s = s & "              , COD.Bldr_Declined = D.Bldr_Declined  " & vbCrLf
s = s & "              , COD.Bldr_Declined_Reason = D.Bldr_Declined_Reason  " & vbCrLf
s = s & "              , COD.comments = D.comments  " & vbCrLf
s = s & "              , COD.Color = D.Color  " & vbCrLf
s = s & "              , COD.Location = D.Location  " & vbCrLf
s = s & "              , COD.SpecSeq = D.SpecSeq  " & vbCrLf
s = s & "              , COD.Style = D.Style  " & vbCrLf
s = s & "              , COD.Finish = D.Finish  " & vbCrLf
s = s & "              , COD.Other = D.Other  " & vbCrLf
s = s & "              , COD.ColorListID = D.ColorListID  " & vbCrLf
s = s & "              , COD.StyleListID = D.StyleListID  " & vbCrLf
s = s & "              , COD.FinishListID = D.FinishListID  " & vbCrLf
s = s & "              , COD.OtherListID = D.OtherListID  " & vbCrLf
s = s & "              , COD.quote_entered_date = D.quote_entered_date  " & vbCrLf
s = s & "              , COD.CRMID = D.CRMID  " & vbCrLf
s = s & "            FROM    @tblChangeOrderDetails D  " & vbCrLf
s = s & "    INNER JOIN ChangeOrderMaster com ON D.Customer_No=COM.Customer_No   " & vbCrLf
s = s & "       AND D.Change_Order_No=COM.Change_Order_No  " & vbCrLf
s = s & "    INNER JOIN ChangeOrderDetails COD ON ((D.Customer_No = COD.Customer_No   " & vbCrLf
s = s & "       AND D.ClientSeq = COD.seq   " & vbCrLf
s = s & "       AND D.Change_Order_No = COD.Change_Order_No) OR D.CRMID = COD.CRMID)  " & vbCrLf
s = s & "  WHERE   COD.TSTMP <=  DATEADD( HOUR, @TimeZoneUTCOffSet,D.ModifiedDate )   " & vbCrLf
s = s & "  --TODO: The CODetails should not update if the CO is approved already.   " & vbCrLf
s = s & "  --@LastPullDate OR COD.TSTMP IS NULL  " & vbCrLf
s = s & "  --COD.Require_Quote = 1  -- CCordoba This column is not used any more Michael email 12/11/2015      " & vbCrLf
s = s & "  --WHERE ISNULL(D.OPTDELETED,0) = 0  " & vbCrLf
s = s & "  --Deleted Options need to be updated as well to a deleted status  " & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "      set @sqlStatement = 'Update ChangeOrderMaster BldrApproved from @@tblChangeOrderMaster'" & vbCrLf
s = s & "   -- Restore BuilderApproved to original state  " & vbCrLf
s = s & "  UPDATE  CO  " & vbCrLf
s = s & "        SET     CO.Bldr_Approved = C.Bldr_Approved,  " & vbCrLf
s = s & "    CO.UserID = C.UserID  " & vbCrLf
s = s & "        FROM    @tblChangeOrderMaster C  " & vbCrLf
s = s & "     INNER JOIN ChangeOrderMaster CO ON (C.Customer_No = CO.Customer_No  " & vbCrLf
s = s & "                                                   AND C.Change_Order_No = CO.Change_Order_No)  " & vbCrLf
s = s & "     -- This has to be uncommented for any other client  " & vbCrLf
s = s & "     OR ( C.CRMID = CO.CRMID)  " & vbCrLf
s = s & "  AND (C.Bldr_Approved = 1 AND CO.Bldr_Approved = 0)  " & vbCrLf
s = s & "     --NOTE: The data transfer should not update unless the HF.Bldr_approve is 0   " & vbCrLf
s = s & "     -- (None Approved) and the CRM is approved. If the client approves in HF then the DT would not update profit.     " & vbCrLf
s = s & "     --Daryl Email 04/20/2016  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "   -- Deposit dates  " & vbCrLf
s = s & " " & vbCrLf
s = s & "" & vbCrLf
s = s & "     set @sqlStatement = 'Update DepositDates from @tblDepositDates' " & vbCrLf
s = s & "" & vbCrLf
s = s & "        UPDATE  CD  " & vbCrLf
s = s & "        SET     CD.Description = D.Description  " & vbCrLf
s = s & "              , CD.Order_index = D.OrderIndex  " & vbCrLf
s = s & "              , CD.CustDate = D.DueDate  " & vbCrLf
s = s & "              , CD.Date_Completed = D.ReceivedDate  " & vbCrLf
s = s & "              , CD.PercentComplete = CASE WHEN D.IsCompleted = 1 THEN 100 ELSE CD.PercentComplete END  " & vbCrLf
s = s & "              , CD.Check_Box = D.IsCompleted   " & vbCrLf
s = s & "              , CD.notice_reqd = D.IsNoticeRequired  " & vbCrLf
s = s & "              , CD.Amount = ISNULL(D.Amount,0)  " & vbCrLf
s = s & "              , CD.AmountPaid = ISNULL(D.AmountPaid, 0)  " & vbCrLf
s = s & "              , CD.PaidDate = D.AmountPaidDate  " & vbCrLf
s = s & "              , CD.Comments = D.Comments  " & vbCrLf
s = s & "              , CD.sales_view = D.IsSalesView  " & vbCrLf
s = s & "              , CD.vendor_webview = D.IsVendorWebView  " & vbCrLf
s = s & "              , CD.Web_View = D.IsContractWebView  " & vbCrLf
s = s & "        FROM    @tblDepositDates D  " & vbCrLf
s = s & "                INNER JOIN dbo.Customer_Date CD ON D.CustomerNo = CD.Customer_No AND D.DateField = CD.Date_Field" & vbCrLf
s = s & "               ;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "       set @sqlStatement = 'Insert DepositDates from @tblDepositDates' " & vbCrLf
s = s & "" & vbCrLf
s = s & "        INSERT  INTO dbo.Customer_Date  " & vbCrLf
s = s & "                (Customer_No  " & vbCrLf
s = s & "               , Date_Field  " & vbCrLf
s = s & "               , Comments  " & vbCrLf
s = s & "               , Description  " & vbCrLf
s = s & "               , notice_reqd  " & vbCrLf
s = s & "               , Amount  " & vbCrLf
s = s & "               , Order_index  " & vbCrLf
s = s & "               , CustDate  " & vbCrLf
s = s & "               , Date_Completed  " & vbCrLf
s = s & "               , PercentComplete  " & vbCrLf
s = s & "               , AmountPaid  " & vbCrLf
s = s & "               , PaidDate  " & vbCrLf
s = s & "               , sales_view  " & vbCrLf
s = s & "               , Web_View  " & vbCrLf
s = s & "               , vendor_webview  " & vbCrLf
s = s & "               , DateType  " & vbCrLf
s = s & "      , Check_Box)  " & vbCrLf
s = s & "                SELECT  D.CustomerNo  " & vbCrLf
s = s & "                      , D.DateField  " & vbCrLf
s = s & "                      , D.Comments  " & vbCrLf
s = s & "                      , D.Description  " & vbCrLf
s = s & "                      , D.IsNoticeRequired  " & vbCrLf
s = s & "                      , ISNULL(D.Amount, 0)  " & vbCrLf
s = s & "                      , D.OrderIndex  " & vbCrLf
s = s & "                      , D.DueDate  " & vbCrLf
s = s & "                      , D.ReceivedDate  " & vbCrLf
s = s & "                      , CASE WHEN D.IsCompleted = 1 THEN 100  " & vbCrLf
s = s & "                             ELSE 0  " & vbCrLf
s = s & "                        END  " & vbCrLf
s = s & "                      , isnull(D.AmountPaid,0)  " & vbCrLf
s = s & "                      , D.AmountPaidDate  " & vbCrLf
s = s & "                      , D.IsSalesView  " & vbCrLf
s = s & "                      , D.IsContractWebView  " & vbCrLf
s = s & "                      , D.IsVendorWebView  " & vbCrLf
s = s & "                      , 3  " & vbCrLf
s = s & "       , D.IsCompleted  " & vbCrLf
s = s & "                FROM    @tblDepositDates D  " & vbCrLf
s = s & "                        LEFT JOIN dbo.Customer_Date CD ON D.CustomerNo = CD.Customer_No AND D.DateField = CD.Date_Field  " & vbCrLf
s = s & "                       inner join dbo.tblCustomers c on c.customer_no = D.CustomerNo" & vbCrLf
s = s & "                WHERE   CD.Date_Field IS NULL AND CD.Customer_No IS NULL;  " & vbCrLf
s = s & "" & vbCrLf
s = s & "    -- Scheduling Dates  " & vbCrLf
s = s & " " & vbCrLf
s = s & "      set @sqlStatement = 'Update ScheduleDates from @tblSchedulingDates'  " & vbCrLf
s = s & "        UPDATE  CD  " & vbCrLf
s = s & "        SET     CD.Description = D.Description  " & vbCrLf
s = s & "              , CD.Order_index = D.OrderIndex  " & vbCrLf
s = s & "              , CD.CustDate = D.StartDate  " & vbCrLf
s = s & "              , CD.Date_Completed = D.FinishDate  " & vbCrLf
s = s & "              , CD.PercentComplete = CASE WHEN D.IsCompleted = 1 THEN 100  " & vbCrLf
s = s & "                                          ELSE CD.PercentComplete  " & vbCrLf
s = s & "                                     END  " & vbCrLf
s = s & "              , CD.Amount = ISNULL(D.Amount, 0)  " & vbCrLf
s = s & "              , CD.AmountPaid = ISNULL(D.AmountPaid,0)  " & vbCrLf
s = s & "              , CD.PaidDate = D.AmountPaidDate  " & vbCrLf
s = s & "              , CD.Comments = D.Comments  " & vbCrLf
s = s & "              , CD.sales_view = D.IsSalesView  " & vbCrLf
s = s & "              , CD.vendor_webview = D.IsVendorWebView  " & vbCrLf
s = s & "              , CD.Web_View = D.IsContractWebView  " & vbCrLf
s = s & "     , CD.Check_Box = D.IsCompleted  " & vbCrLf
s = s & "        FROM    @tblSchedulingDates D  " & vbCrLf
s = s & "                INNER JOIN dbo.Customer_Date CD ON D.CustomerNo = CD.Customer_No AND D.DateField = CD.Date_Field" & vbCrLf
s = s & "               ;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "" & vbCrLf
s = s & "   set @sqlStatement = 'Insert ScheduleDates from @tblSchedulingDates'  " & vbCrLf
s = s & "" & vbCrLf
s = s & "     INSERT  INTO dbo.Customer_Date  " & vbCrLf
s = s & "                (Customer_No  " & vbCrLf
s = s & "               , Date_Field  " & vbCrLf
s = s & "               , Comments  " & vbCrLf
s = s & "               , Description  " & vbCrLf
s = s & "               , Amount  " & vbCrLf
s = s & "               , Order_index  " & vbCrLf
s = s & "               , CustDate  " & vbCrLf
s = s & "               , Date_Completed  " & vbCrLf
s = s & "               , AmountPaid  " & vbCrLf
s = s & "               , PaidDate  " & vbCrLf
s = s & "               , sales_view  " & vbCrLf
s = s & "               , Web_View  " & vbCrLf
s = s & "               , Supplier  " & vbCrLf
s = s & "               , vendor_webview  " & vbCrLf
s = s & "               , DateType  " & vbCrLf
s = s & "               , PercentComplete  " & vbCrLf
s = s & "               , notice_reqd  " & vbCrLf
s = s & "      , Check_Box)  " & vbCrLf
s = s & "                SELECT  D.CustomerNo  " & vbCrLf
s = s & "                      , D.DateField  " & vbCrLf
s = s & "                      , D.Comments  " & vbCrLf
s = s & "                      , D.Description  " & vbCrLf
s = s & "                      , ISNULL(D.Amount, 0)  " & vbCrLf
s = s & "                      , D.OrderIndex  " & vbCrLf
s = s & "                      , D.StartDate  " & vbCrLf
s = s & "                      , D.FinishDate  " & vbCrLf
s = s & "                      , ISNULL(D.AmountPaid,0)  " & vbCrLf
s = s & "                      , D.AmountPaidDate  " & vbCrLf
s = s & "                      , D.IsSalesView  " & vbCrLf
s = s & "                      , D.IsContractWebView  " & vbCrLf
s = s & "                      , D.Supplier  " & vbCrLf
s = s & "                      , D.IsVendorWebView  " & vbCrLf
s = s & "                      , 2  " & vbCrLf
s = s & "                      , CASE WHEN D.IsCompleted = 1 THEN 100  " & vbCrLf
s = s & "                             ELSE 0  " & vbCrLf
s = s & "                        END  " & vbCrLf
s = s & "                      , 0  " & vbCrLf
s = s & "       , D.IsCompleted  " & vbCrLf
s = s & "                FROM    @tblSchedulingDates D  " & vbCrLf
s = s & "                        LEFT JOIN dbo.Customer_Date CD ON D.CustomerNo = CD.Customer_No AND D.DateField = CD.Date_Field " & vbCrLf
s = s & "                       inner join dbo.tblCustomers c on c.customer_no = D.CustomerNo " & vbCrLf
s = s & "                WHERE   CD.Date_Field IS NULL  " & vbCrLf
s = s & "                        AND CD.Customer_No IS NULL;  " & vbCrLf
s = s & "" & vbCrLf
s = s & "    -- Condition Dates  " & vbCrLf
s = s & " set @sqlStatement = 'Update ConditionDates from @tblConditionDates'    " & vbCrLf
s = s & "        UPDATE  CD  " & vbCrLf
s = s & "        SET     CD.Description = D.Description  " & vbCrLf
s = s & "              , CD.Order_index = D.OrderIndex  " & vbCrLf
s = s & "              , CD.CustDate = D.ExpectedRemovalDate  " & vbCrLf
s = s & "              , CD.Date_Completed = D.ActualRemovalDate  " & vbCrLf
s = s & "              , CD.PercentComplete = CASE WHEN D.IsCompleted = 1 THEN 100  " & vbCrLf
s = s & "                                          ELSE CD.PercentComplete  " & vbCrLf
s = s & "                                     END  " & vbCrLf
s = s & "              , CD.Amount = ISNULL(D.Amount, 0)  " & vbCrLf
s = s & "              , CD.Comments = D.Comments  " & vbCrLf
s = s & "              , CD.sales_view = D.IsSalesView  " & vbCrLf
s = s & "              , CD.vendor_webview = D.IsVendorWebView  " & vbCrLf
s = s & "              , CD.Web_View = D.IsContractWebView  " & vbCrLf
s = s & "     , CD.Check_Box = D.IsCompleted  " & vbCrLf
s = s & "        FROM    @tblConditionDates D  " & vbCrLf
s = s & "                INNER JOIN dbo.Customer_Date CD ON D.CustomerNo = CD.Customer_No  " & vbCrLf
s = s & "                                                   AND D.DateField = CD.Date_Field;  " & vbCrLf
s = s & "" & vbCrLf
s = s & "      " & vbCrLf
s = s & "set @sqlStatement = 'Insert ConditionDates from @tblConditionDates' " & vbCrLf
s = s & "" & vbCrLf
s = s & "        INSERT  INTO dbo.Customer_Date  " & vbCrLf
s = s & "                ( Customer_No ,  " & vbCrLf
s = s & "                  Date_Field ,  " & vbCrLf
s = s & "                  Comments ,  " & vbCrLf
s = s & "                  Description ,  " & vbCrLf
s = s & "                  Amount ,  " & vbCrLf
s = s & "                  Order_index ,  " & vbCrLf
s = s & "                  CustDate ,  " & vbCrLf
s = s & "                  Date_Completed ,  " & vbCrLf
s = s & "                  sales_view ,  " & vbCrLf
s = s & "                  Web_View ,  " & vbCrLf
s = s & "                  vendor_webview ,  " & vbCrLf
s = s & "                  DateType ,  " & vbCrLf
s = s & "                  PercentComplete ,  " & vbCrLf
s = s & "                  notice_reqd ,  " & vbCrLf
s = s & "      Check_Box  " & vbCrLf
s = s & "                )  " & vbCrLf
s = s & "                SELECT  D.CustomerNo ,  " & vbCrLf
s = s & "                        D.DateField ,  " & vbCrLf
s = s & "                        D.Comments ,  " & vbCrLf
s = s & "                        D.Description ,  " & vbCrLf
s = s & "                        ISNULL(D.Amount, 0) ,  " & vbCrLf
s = s & "                        D.OrderIndex ,  " & vbCrLf
s = s & "                        D.ExpectedRemovalDate ,  " & vbCrLf
s = s & "                        D.ActualRemovalDate ,  " & vbCrLf
s = s & "                        D.IsSalesView ,  " & vbCrLf
s = s & "                        D.IsContractWebView ,  " & vbCrLf
s = s & "                        D.IsVendorWebView ,  " & vbCrLf
s = s & "                        4 ,  " & vbCrLf
s = s & "                        CASE WHEN D.IsCompleted = 1 THEN 100  " & vbCrLf
s = s & "                             ELSE 0  " & vbCrLf
s = s & "                        END ,  " & vbCrLf
s = s & "                        0 ,  " & vbCrLf
s = s & "      D.IsCompleted  " & vbCrLf
s = s & "                FROM    @tblConditionDates D  " & vbCrLf
s = s & "                        LEFT JOIN dbo.Customer_Date CD ON D.CustomerNo = CD.Customer_No  AND D.DateField = CD.Date_Field " & vbCrLf
s = s & "                       inner join dbo.tblCustomers c on c.customer_no = D.CustomerNo " & vbCrLf
s = s & "                WHERE   CD.Date_Field IS NULL  " & vbCrLf
s = s & "                        AND CD.Customer_No IS NULL;  " & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  -- Inventory Homes migration (w/custom options in Addendum or CO)  " & vbCrLf
s = s & "  -- Do the options operations before tblCustomer updates, otherwise the updates will never execute due to using tblCustomer's Modified Date  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "    set @Reccount = 0" & vbCrLf
s = s & "          " & vbCrLf
s = s & "  --set @Reccount = isnull((select count(*) from @tblInventoryHome),0)" & vbCrLf
s = s & "  --if @Reccount>0 begin" & vbCrLf
s = s & "  -- select * into tmptblInventoryHome from @tblInventoryHome" & vbCrLf
s = s & "  --end" & vbCrLf
s = s & "" & vbCrLf
s = s & "  set @sqlStatement = 'Insert tblcustomers from @tblInventoryHome' " & vbCrLf
s = s & "" & vbCrLf
s = s & "  INSERT  INTO dbo.tblCustomers  " & vbCrLf
s = s & "                ( Customer_No ,  " & vbCrLf
s = s & "                  [Description] ,  " & vbCrLf
s = s & "                  Job_No ,  " & vbCrLf
s = s & "                  ReservedExpiry ,  " & vbCrLf
s = s & "                  Community ,  " & vbCrLf
s = s & "                  [Floor] ,  " & vbCrLf
s = s & "                  Unit_No ,  " & vbCrLf
s = s & "                  Legal_Unit_No ,  " & vbCrLf
s = s & "                  DirectionID ,  " & vbCrLf
s = s & "                  UnitFactor ,  " & vbCrLf
s = s & "                  CondoFees ,  " & vbCrLf
s = s & "                  Model ,  " & vbCrLf
s = s & "                 Base_House ,  " & vbCrLf
s = s & "                  Override_Price ,  " & vbCrLf
s = s & "                  ScheduleB_total ,  " & vbCrLf
s = s & "                  GST_Rate ,  " & vbCrLf
s = s & "                  GST ,  " & vbCrLf
s = s & "                  GST_Rebate ,  " & vbCrLf
s = s & "                  RevenueAdjustment ,  " & vbCrLf
s = s & "                  Lot ,  " & vbCrLf
s = s & "                  Lot_No ,  " & vbCrLf
s = s & "                  Lot_Price ,  " & vbCrLf
s = s & "                  Lot_Cost ,  " & vbCrLf
s = s & "                  LotPremiumValue ,  " & vbCrLf
s = s & "                  LotPremiumPreTax ,  " & vbCrLf
s = s & "                  Parking1 ,  " & vbCrLf
s = s & "                  Parking1_Cost,  " & vbCrLf
s = s & "                  Parking1_Price,  " & vbCrLf
s = s & "                  Parking2 ,  " & vbCrLf
s = s & "                  Parking2_Cost,  " & vbCrLf
s = s & "                  Parking2_Price,  " & vbCrLf
s = s & "                  Parking3 ,  " & vbCrLf
s = s & "                  Parking3_Cost,  " & vbCrLf
s = s & "                  Parking3_Price,  " & vbCrLf
s = s & "                  PST_Rate ,  " & vbCrLf
s = s & "                  PST ,  " & vbCrLf
s = s & "                  PSTRebate ,  " & vbCrLf
s = s & "                  Comments ,  " & vbCrLf
s = s & "                  ModelAssembly ,  " & vbCrLf
s = s & "                  ModelSalesWorksheet ,  " & vbCrLf
s = s & "                  ModelSpecDoc ,  " & vbCrLf
s = s & "                  ModelSpecDate ,  " & vbCrLf
s = s & "                  CO_PD_Total ,  " & vbCrLf
s = s & "                  NotAvailableforSale ,  " & vbCrLf
s = s & "                  Sales_Person_ID ,  " & vbCrLf
s = s & "                  PM ,  " & vbCrLf
s = s & "                  Phase ,  " & vbCrLf
s = s & "                  Series ,  " & vbCrLf
s = s & "                  Square_Footage ,  " & vbCrLf
s = s & "                  --Expected_Occupancy ,  " & vbCrLf
s = s & "                  Construction_Status ,  " & vbCrLf
s = s & "                  Total_Sales_Price ,  " & vbCrLf
s = s & "                  LastEstimateIndex ,  " & vbCrLf
s = s & "                  EstimateIndex ,  " & vbCrLf
s = s & "                  DivisionID ,  " & vbCrLf
s = s & "                 Home_Selection ,  " & vbCrLf
s = s & "                 PreSale_Selection ,  " & vbCrLf
s = s & "                  CRMType ,  " & vbCrLf
s = s & "                  CRMID ,  " & vbCrLf
s = s & "                  ModelIncentive ,  " & vbCrLf
s = s & "                  Sales_Initiative ,  " & vbCrLf
s = s & "                  Sales_Deduction ,  " & vbCrLf
s = s & "                  TotalIncentives ,  " & vbCrLf
s = s & "                 CreatedDate ,  " & vbCrLf
s = s & "                 ModifiedDate,  " & vbCrLf
s = s & "                 Cost_Amount  " & vbCrLf
s = s & "                 , AddedToServiceModule  " & vbCrLf
s = s & "                 , Municipal_Address  " & vbCrLf
s = s & "                 , LegalAddress  " & vbCrLf
s = s & "                 , County  " & vbCrLf
s = s & "                 , Zip  " & vbCrLf
s = s & "                 , City  " & vbCrLf
s = s & "                 , Province  " & vbCrLf
s = s & "                 , Lot_Zip  " & vbCrLf
s = s & "                 , Lot_City  " & vbCrLf
s = s & "                 , Lot_Province " & vbCrLf
s = s & "                 ,Cancelled " & vbCrLf
s = s & "                 ,InventoryHomeApproved" & vbCrLf
s = s & "                 ,Bal_Sheet_Prefix" & vbCrLf
s = s & "                 ,Income_Prefix" & vbCrLf
s = s & "     )  " & vbCrLf
s = s & "                SELECT  ih.CustomerNo ,  " & vbCrLf
s = s & "                        ih.[Description] ,  " & vbCrLf
s = s & "                        ih.Job ,  " & vbCrLf
s = s & "                        ih.ReservedExpiry ,  " & vbCrLf
s = s & "                        ih.Community ,  " & vbCrLf
s = s & "                        ih.[Floor] ,  " & vbCrLf
s = s & "                        ih.UnitNo ,  " & vbCrLf
s = s & "                        ih.LegalUnit ,  " & vbCrLf
s = s & "                        ih.[DirectionFacing] ,  " & vbCrLf
s = s & "                        ih.UnitFactor ,  " & vbCrLf
s = s & "                        ih.[StrataFees] ,  " & vbCrLf
s = s & "                        ih.Model ,  " & vbCrLf
s = s & "                        ih.ModelPrice ,  " & vbCrLf
s = s & "                       ih.ModelPrice ,  " & vbCrLf
s = s & "                        ih.AddendumTotal ,  " & vbCrLf
s = s & "                        ih.GSTRate ,  " & vbCrLf
s = s & "                        ih.GST ,  " & vbCrLf
s = s & "                        ih.GSTRebate ,  " & vbCrLf
s = s & "                        ISNULL(ih.RevenueAdj,0.00) AS RevenueAdj ,  " & vbCrLf
s = s & "                        ih.Lot ,  " & vbCrLf
s = s & "                        ih.LotNo ,  " & vbCrLf
s = s & "                        ih.LotPrice ,  " & vbCrLf
s = s & "                        ih.LotCost ,  " & vbCrLf
s = s & "                        ISNULL(ih.PremiumAmount,0.00 ) as PremiumAmount ,  " & vbCrLf
s = s & "                        ISNULL(ih.PremiumAmount,0.00) as PremiumPreTax , -- ih.PremiumPreTax ,  " & vbCrLf
s = s & "                        ih.InventoryHomesParking1 ,  " & vbCrLf
s = s & "                        ISNULL(ih.Parking1Cost ,0.00 ) as Parking1Cost,  " & vbCrLf
s = s & "                        ISNULL(ih.Parking1SellingPrice ,0.00 ) as Parking1SellingPrice,  " & vbCrLf
s = s & "                        ih.InventoryHomesParking2 ,  " & vbCrLf
s = s & "                        ISNULL(ih.Parking2Cost  ,0.00 ) as Parking2Cost,  " & vbCrLf
s = s & "                        ISNULL(ih.Parking2SellingPrice ,0.00 ) as Parking2SellingPrice,  " & vbCrLf
s = s & "                        ih.InventoryHomesParking3 ,  " & vbCrLf
s = s & "                        ISNULL(ih.Parking3Cost ,0.00 ) as Parking3Cost,  " & vbCrLf
s = s & "                        ISNULL(ih.Parking3SellingPrice  ,0.00 ) asParking3SellingPrice,  " & vbCrLf
s = s & "                        ih.PSTRate ,  " & vbCrLf
s = s & "                        ih.PST ,  " & vbCrLf
s = s & "                        ih.PSTRebate ,  " & vbCrLf
s = s & "                        ih.Comments ,  " & vbCrLf
s = s & "                        ih.ModelAssembly ,  " & vbCrLf
s = s & "                        ih.ModelSalesWorksheet ,  " & vbCrLf
s = s & "                        ih.ModelSpecDoc ,  " & vbCrLf
s = s & "                        ih.ModelSpecDate ,  " & vbCrLf
s = s & "                        ih.ChangeOrderTotal ,  " & vbCrLf
s = s & "                        CASE WHEN ih.IsAvailableForSale = 1 THEN 0  " & vbCrLf
s = s & "                             ELSE 1  " & vbCrLf
s = s & "                        END ,  " & vbCrLf
s = s & "                        ih.SalesPerson ,  " & vbCrLf
s = s & "                        ih.SiteSupervisor ,  " & vbCrLf
s = s & "                        '' ,  " & vbCrLf
s = s & "                        ih.Series ,  " & vbCrLf
s = s & "                        ih.SquareFoot ,  " & vbCrLf
s = s & "                        --ih.CompletionDate ,  " & vbCrLf
s = s & "                        ih.ConstructionStatus ,  " & vbCrLf
s = s & "                        ih.TotalPrice ,  " & vbCrLf
s = s & "                        ih.LastEstimateIndex ,  " & vbCrLf
s = s & "                        ih.EstimateIndex ,  " & vbCrLf
s = s & "                        ih.Division ,  " & vbCrLf
s = s & "                       ih.HomeSelection ,  " & vbCrLf
s = s & "                       ih.PreSaleSelection ,  " & vbCrLf
s = s & "                        ih.CRMType ,  " & vbCrLf
s = s & "                        ih.CRMID ,  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                         ih.ModelIncentive ,  " & vbCrLf
s = s & "                         ih.SalesIncentive ,  " & vbCrLf
s = s & "                         ih.SalesDeduction ,  " & vbCrLf
s = s & "                         ih.TotalIncentives ,  " & vbCrLf
s = s & "                         ih.CreatedDate ,  " & vbCrLf
s = s & "                         ih.ModifiedDate,  " & vbCrLf
s = s & "                         ih.ModelCostAmount,  " & vbCrLf
s = s & "                         1 AS AddedToServiceModule,  " & vbCrLf
s = s & "                         ih.MunicipalAddress   " & vbCrLf
s = s & "                         , l.LegalAddress   " & vbCrLf
s = s & "                         , l.County  " & vbCrLf
s = s & "                         , l.Zip  " & vbCrLf
s = s & "                         , l.City  " & vbCrLf
s = s & "                         , l.Province  " & vbCrLf
s = s & "                         , l.zip As LotZip  " & vbCrLf
s = s & "                         , l.city As LotCity  " & vbCrLf
s = s & "                         , l.Province As Province  " & vbCrLf
s = s & "                         , ih.[IsCancelled]" & vbCrLf
s = s & "                         ,1" & vbCrLf
s = s & "                         , b.Prefix1" & vbCrLf
s = s & "                         , b.Prefix2" & vbCrLf
s = s & "                FROM    @tblInventoryHome ih  " & vbCrLf
s = s & "                        LEFT JOIN dbo.tblCustomers c ON ih.CustomerNo = c.Customer_No  " & vbCrLf
s = s & "                       join tbllocality b on b.Area = ih.Community" & vbCrLf
s = s & "      LEFT JOIN DBO.tblLotInventory l on l.Lot = ih.Lot and l.Lot_no = ih.LotNo and l.Job_no = ih.Job  " & vbCrLf
s = s & "                WHERE   c.Customer_No IS NULL;  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & " set @sqlStatement = 'Update tblcustomers from @tblInventoryHome' " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "UPDATE  c  " & vbCrLf
s = s & "        SET     c.[Description] = ih.[Description] ,  " & vbCrLf
s = s & "                c.Job_No = Case When c.sold=1 then c.Job_No else ih.Job end ,  " & vbCrLf
s = s & "                c.ReservedExpiry = ih.ReservedExpiry ,  " & vbCrLf
s = s & "                c.Community = ih.Community ,  " & vbCrLf
s = s & "                c.[Floor] = ih.[Floor] ,  " & vbCrLf
s = s & "                c.Unit_No = ih.UnitNo ,  " & vbCrLf
s = s & "                c.Legal_Unit_No = ih.LegalUnit ,  " & vbCrLf
s = s & "                c.DirectionID = ih.DirectionFacing ,  " & vbCrLf
s = s & "                c.UnitFactor = ih.UnitFactor ,  " & vbCrLf
s = s & "                c.CondoFees = ih.StrataFees ,  " & vbCrLf
s = s & "                c.Model = ih.Model ,  " & vbCrLf
s = s & "               c.Series = ih.Series,   " & vbCrLf
s = s & "                c.Override_Price = ih.ModelPrice ,  " & vbCrLf
s = s & "               c.Base_House = ih.ModelPrice ,  " & vbCrLf
s = s & "                c.ScheduleB_total = ih.AddendumTotal ,  " & vbCrLf
s = s & "                c.GST_Rate = ih.GSTRate ,  " & vbCrLf
s = s & "                c.GST = ih.GST ,  " & vbCrLf
s = s & "                c.GST_Rebate = ih.GSTRebate ,  " & vbCrLf
s = s & "                c.RevenueAdjustment = ISNULL(ih.RevenueAdj,0.00) ,  " & vbCrLf
s = s & "                c.Lot = ih.Lot ,  " & vbCrLf
s = s & "                c.Lot_No = ih.LotNo ,  " & vbCrLf
s = s & "                c.Lot_Price = ih.LotPrice ,  " & vbCrLf
s = s & "                c.Lot_Cost = ih.LotCost ,  " & vbCrLf
s = s & "                c.LotPremiumValue = isnull(ih.PremiumAmount,0.00) ,  " & vbCrLf
s = s & "                c.LotPremiumPreTax =isnull(ih.PremiumAmount,0.00) ,  --isnull(ih.PremiumPreTax,0.00) ,  " & vbCrLf
s = s & "                c.Parking1 = ih.InventoryHomesParking1 ,  " & vbCrLf
s = s & "                c.Parking1_Cost = ISNULL(ih.Parking1Cost,0.00) ,  " & vbCrLf
s = s & "                c.Parking1_Price = ISNULL(ih.Parking1SellingPrice,0.00) ,  " & vbCrLf
s = s & "                c.Parking2 = ih.InventoryHomesParking2 ,  " & vbCrLf
s = s & "                c.Parking2_Cost = ISNULL(ih.Parking2Cost,0.00) ,  " & vbCrLf
s = s & "                c.Parking2_Price = ISNULL(ih.Parking2SellingPrice,0.00) ,  " & vbCrLf
s = s & "                c.Parking3 = ih.InventoryHomesParking3 ,  " & vbCrLf
s = s & "                c.Parking3_Cost = ISNULL(ih.Parking3Cost,0.00) ,  " & vbCrLf
s = s & "                c.Parking3_Price = ISNULL(ih.Parking3SellingPrice,0.00) ,  " & vbCrLf
s = s & "                c.PST_Rate = ih.PSTRate ,  " & vbCrLf
s = s & "                c.PST = ih.PST ,  " & vbCrLf
s = s & "                c.PSTRebate = ih.PSTRebate ,  " & vbCrLf
s = s & "                c.Comments = ih.Comments ,  " & vbCrLf
s = s & "                c.ModelAssembly = ih.ModelAssembly ,  " & vbCrLf
s = s & "                c.ModelSalesWorksheet = ih.ModelSalesWorksheet ,  " & vbCrLf
s = s & "                c.ModelSpecDoc = ih.ModelSpecDoc ,  " & vbCrLf
s = s & "                c.ModelSpecDate = ih.ModelSpecDate ,  " & vbCrLf
s = s & "                c.CO_PD_Total = ih.ChangeOrderTotal ,  " & vbCrLf
s = s & "                c.NotAvailableforSale = CASE WHEN ih.IsAvailableForSale = 1  " & vbCrLf
s = s & "                                             THEN 0  " & vbCrLf
s = s & "                                             ELSE 1  " & vbCrLf
s = s & "                                        END ,  " & vbCrLf
s = s & "                c.Sales_Person_ID = ih.SalesPerson ,  " & vbCrLf
s = s & "                c.PM = ih.SiteSupervisor ,  " & vbCrLf
s = s & "        -- c.Phase  = ih.,  " & vbCrLf
s = s & "        -- c.Series  = ih.,  " & vbCrLf
s = s & "                c.Square_Footage = ih.SquareFoot ,  " & vbCrLf
s = s & "                --c.Expected_Occupancy = ih.CompletionDate ,  " & vbCrLf
s = s & "                --c.Construction_Status = ih.ConstructionStatus ,  " & vbCrLf
s = s & "                c.Total_Sales_Price = ih.TotalPrice ,                                          " & vbCrLf
s = s & "                c.DivisionID = ih.Division ,  " & vbCrLf
s = s & "                c.CRMType = ih.CRMType ,  " & vbCrLf
s = s & "                c.CRMID = ih.CRMID ,  " & vbCrLf
s = s & "               c.ModelIncentive = ih.ModelIncentive ,  " & vbCrLf
s = s & "               c.Sales_Initiative = ih.SalesIncentive ,  " & vbCrLf
s = s & "               c.Sales_Deduction = ih.SalesDeduction ,  " & vbCrLf
s = s & "               c.TotalIncentives = ih.TotalIncentives ,  " & vbCrLf
s = s & "               c.CreatedDate = ih.CreatedDate ,   " & vbCrLf
s = s & "               c.ModifiedDate = ih.ModifiedDate,  " & vbCrLf
s = s & "               c.Municipal_Address = ih.MunicipalAddress,  " & vbCrLf
s = s & "               c.Cost_Amount = ih.ModelCostAmount  ," & vbCrLf
s = s & "               c.Cancelled = ih.[IsCancelled]," & vbCrLf
s = s & "               c.InventoryHomeApproved=1" & vbCrLf
s = s & "             , Bal_Sheet_Prefix=case when isnull(c.Bal_Sheet_Prefix,'')<>'' then c.Bal_Sheet_Prefix else b.Prefix1 end" & vbCrLf
s = s & "             , Income_Prefix=case when isnull(c.Income_Prefix,'')<>'' then c.Income_Prefix else b.Prefix2 end" & vbCrLf
s = s & "        FROM    @tblInventoryHome ih  " & vbCrLf
s = s & "                INNER JOIN dbo.tblCustomers c ON ih.CustomerNo = c.Customer_No OR c.CRMID = ih.CRMID  " & vbCrLf
s = s & "               JOIN tbllocality b on b.area = c.community" & vbCrLf
s = s & "  WHERE c.ModifiedDate <= DATEADD( HOUR, @TimeZoneUTCOffSet,ih.ModifiedDate )   " & vbCrLf
s = s & "  --@LastPullDate   " & vbCrLf
s = s & "  OR c.ModifiedDate IS NULL;  " & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & " set @sqlStatement = 'Insert tblScheduleB from @tblInventoryHomeAddendum' " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  INSERT  INTO dbo.tblScheduleB  " & vbCrLf
s = s & "                ( Customer_No ,  " & vbCrLf
s = s & "                  OPT ,  " & vbCrLf
s = s & "                  Option_Type ,  " & vbCrLf
s = s & "                  Custom ,  " & vbCrLf
s = s & "                  QTY ,  " & vbCrLf
s = s & "                  UOM ,  " & vbCrLf
s = s & "                  Base_Price ,  " & vbCrLf
s = s & "                  Rate ,  " & vbCrLf
s = s & "                  TAX ,  " & vbCrLf
s = s & "                  Cost_Amount ,  " & vbCrLf
s = s & "                  Override_Price ,  " & vbCrLf
s = s & "                  TotalAmount ,  " & vbCrLf
s = s & "                  GrandTotal ,  " & vbCrLf
s = s & "                  [Description] ,  " & vbCrLf
s = s & "                  comments ,  " & vbCrLf
s = s & "                  Major_Group , -- This is the Category  " & vbCrLf
s = s & "                  Category ,  " & vbCrLf
s = s & "                  [Assembly] ,  " & vbCrLf
s = s & "                  Color ,  " & vbCrLf
s = s & "                  Style ,  " & vbCrLf
s = s & "                  Finish ,  " & vbCrLf
s = s & "                  SalesWorksheet ,  " & vbCrLf
s = s & "                  ApplyIncentive ,  " & vbCrLf
s = s & "                  IncentiveValue ,  " & vbCrLf
s = s & "                  Location ,  " & vbCrLf
s = s & "                  EstimateIndex ,  " & vbCrLf
s = s & "                  Other ,  " & vbCrLf
s = s & "                  CRMType ,  " & vbCrLf
s = s & "                  CRMID,  " & vbCrLf
s = s & "                 Require_Quote ," & vbCrLf
s = s & "                 UserID" & vbCrLf
s = s & "                )  " & vbCrLf
s = s & "                SELECT  ia.CustomerNo ,  " & vbCrLf
s = s & "                        ia.[Option] ,  " & vbCrLf
s = s & "                        ia.OptionType ,  " & vbCrLf
s = s & "                        ia.IsCustom ,  " & vbCrLf
s = s & "                        ia.QTY ,  " & vbCrLf
s = s & "                        ia.UOM ,  " & vbCrLf
s = s & "                        ia.BasePrice ,  " & vbCrLf
s = s & "                        ia.Rate ,  " & vbCrLf
s = s & "                        ia.Tax ,  " & vbCrLf
s = s & "                        ia.CostRate ,  " & vbCrLf
s = s & "                        ia.OverridePrice ,  " & vbCrLf
s = s & "                        isnull(ia.TaxInTotal,0.00) TaxInTotal ,  " & vbCrLf
s = s & "                        isnull(ia.GrandTotal,0.00) GrandTotal ,  " & vbCrLf
s = s & "                        ia.[Description] ,  " & vbCrLf
s = s & "                        ia.Comments ,  " & vbCrLf
s = s & "                        ia.Category ,  " & vbCrLf
s = s & "                        ia.SubCategory ,  " & vbCrLf
s = s & "                        ia.[Assembly] ,  " & vbCrLf
s = s & "                        ia.Color ,  " & vbCrLf
s = s & "                        ia.Style ,  " & vbCrLf
s = s & "                        ia.Finish ,  " & vbCrLf
s = s & "                        ia.SalesWorksheet ,  " & vbCrLf
s = s & "                        ia.IsApplyIncentive ,  " & vbCrLf
s = s & "                        ia.IncentiveValue ,  " & vbCrLf
s = s & "                        ia.Location ,  " & vbCrLf
s = s & "                        ia.EstimateIndex ,  " & vbCrLf
s = s & "                        ia.Other ,  " & vbCrLf
s = s & "                        ia.CRMType ,  " & vbCrLf
s = s & "                        ia.CRMID,  " & vbCrLf
s = s & "                       isnull(ia.IsRequireQuote, 0)," & vbCrLf
s = s & "                       ia.SalesPerson" & vbCrLf
s = s & "      " & vbCrLf
s = s & "                FROM    @tblInventoryHomesAddendum ia  " & vbCrLf
s = s & "                        LEFT JOIN dbo.tblScheduleB sb ON ia.CRMID = sb.CRMID  " & vbCrLf
s = s & "                                                         AND ia.CRMType = sb.CRMType  " & vbCrLf
s = s & "                WHERE   sb.Customer_No IS NULL AND ia.ClientSeq IS NULL  " & vbCrLf
s = s & "    AND ia.IsOptDelete = 0;  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & " set @sqlStatement = 'Update tblScheduleB from @tblInventoryHomeAddendum' " & vbCrLf
s = s & "" & vbCrLf
s = s & " " & vbCrLf
s = s & "        UPDATE  sb  " & vbCrLf
s = s & "        SET     sb.OPT = ia.[Option] ,  " & vbCrLf
s = s & "                sb.Option_Type = ia.OptionType ,  " & vbCrLf
s = s & "                sb.Custom = ia.IsCustom ,  " & vbCrLf
s = s & "                sb.QTY = ia.QTY ,  " & vbCrLf
s = s & "                sb.UOM = ia.UOM ,  " & vbCrLf
s = s & "                sb.Base_Price = ia.BasePrice ,  " & vbCrLf
s = s & "                sb.Rate = ia.Rate ,  " & vbCrLf
s = s & "                sb.TAX = ia.Tax ,  " & vbCrLf
s = s & "                sb.Cost_Amount = ia.CostRate ,  " & vbCrLf
s = s & "                sb.Override_Price = ia.OverridePrice ,  " & vbCrLf
s = s & "                sb.TotalAmount = ISNULL(ia.TaxInTotal,0.00) ,  " & vbCrLf
s = s & "                sb.GrandTotal = ISNULL(ia.GrandTotal,0.00) ,  " & vbCrLf
s = s & "                sb.[Description] = ia.[Description] ,  " & vbCrLf
s = s & "                sb.comments = ia.Comments ,  " & vbCrLf
s = s & "                sb.Major_Group = ia.Category ,  " & vbCrLf
s = s & "                sb.Category = ia.SubCategory ,  " & vbCrLf
s = s & "                sb.[Assembly] = ia.[Assembly] ,  " & vbCrLf
s = s & "                sb.Color = ia.Color ,  " & vbCrLf
s = s & "                sb.Style = ia.Style ,  " & vbCrLf
s = s & "                sb.Finish = ia.Finish ,  " & vbCrLf
s = s & "                sb.SalesWorksheet = ia.SalesWorksheet ,  " & vbCrLf
s = s & "                sb.ApplyIncentive = ia.IsApplyIncentive ,  " & vbCrLf
s = s & "                sb.IncentiveValue = ia.IncentiveValue ,  " & vbCrLf
s = s & "                sb.Location = ia.Location ,    " & vbCrLf
s = s & "                sb.Other = ia.Other,  " & vbCrLf
s = s & "    sb.Require_Quote = isnull(ia.IsRequireQuote,0)  " & vbCrLf
s = s & "        FROM    @tblInventoryHomesAddendum ia  " & vbCrLf
s = s & "                INNER JOIN dbo.tblScheduleB sb ON (ia.CRMID = sb.CRMID OR ia.ClientSeq = sb.seq)  " & vbCrLf
s = s & "                                                  AND ia.CRMType = sb.CRMType  " & vbCrLf
s = s & "    INNER JOIN dbo.tblCustomers c ON sb.Customer_No = c.Customer_No  " & vbCrLf
s = s & "  WHERE sb.ModifiedDate <= DATEADD( HOUR, @TimeZoneUTCOffSet,ia.ModifiedDate )   " & vbCrLf
s = s & "  --@LastPullDate;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "   " & vbCrLf
s = s & "" & vbCrLf
s = s & "        " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  set @sqlStatement = 'Insert ChangeOrderMaster from @tblInventoryHomeCOMaster' " & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "        INSERT  INTO dbo.ChangeOrderMaster  " & vbCrLf
s = s & "                ( Customer_No   " & vbCrLf
s = s & "                  ,Change_Order_No ,  " & vbCrLf
s = s & "                  Change_Date ,  " & vbCrLf
s = s & "                  Bldr_Approved_By ,  " & vbCrLf
s = s & "                  Bldr_Approved_Date ,  " & vbCrLf
s = s & "                  Paid_Date ,  " & vbCrLf
s = s & "                  Next_Item_No ,  " & vbCrLf
s = s & "                  Bldr_Approved ,  " & vbCrLf
s = s & "                  Total_Price ,  " & vbCrLf
s = s & "                  Amount_Paid ,  " & vbCrLf
s = s & "                  Addition_Change ,  " & vbCrLf
s = s & "                  Declined ,  " & vbCrLf
s = s & "                  Declined_By ,  " & vbCrLf
s = s & "                  Declined_Notice ,  " & vbCrLf
s = s & "                  UserID ,  " & vbCrLf
s = s & "                  NoCommTotal ,  " & vbCrLf
s = s & "                 ChargedFee ,  " & vbCrLf
s = s & "                  COLocked ,  " & vbCrLf
s = s & "                  CreationDate ,  " & vbCrLf
s = s & "                  RebateExempt ,  " & vbCrLf
s = s & "                  Sales_Person_ID ,  " & vbCrLf
s = s & "                  contractexempt ,  " & vbCrLf
s = s & "                  SumOfTotalAmount ,  " & vbCrLf
s = s & "                  SumOfNetTAX ,  " & vbCrLf
s = s & "                  NotTaxedTotal ,  " & vbCrLf
s = s & "                  EstimateIndex ,  " & vbCrLf
s = s & "                  ApplyIncentive ,  " & vbCrLf
s = s & "                  comments ,  " & vbCrLf
s = s & "                  CRMID ,  " & vbCrLf
s = s & "                 CRMType  " & vbCrLf
s = s & "                )  " & vbCrLf
s = s & "                SELECT  icm.CustomerNo   " & vbCrLf
s = s & "                        ,icm.ChangeOrderNo ,  " & vbCrLf
s = s & "                        icm.ChangeDate ,  " & vbCrLf
s = s & "                        icm.BldrApprovedBy ,  " & vbCrLf
s = s & "                        icm.BldrApprovedDate ,  " & vbCrLf
s = s & "                        icm.PaidDate ,  " & vbCrLf
s = s & "                        icm.NextItemNo ,  " & vbCrLf
s = s & "                        icm.IsBldrApproved ,  " & vbCrLf
s = s & "                       --0 ,  " & vbCrLf
s = s & "                        icm.TotalPrice ,  " & vbCrLf
s = s & "                        ISNULL(icm.AmountPaid,0) ,  " & vbCrLf
s = s & "                        icm.IsAdditionChange ,  " & vbCrLf
s = s & "                        icm.IsDeclined ,  " & vbCrLf
s = s & "                        icm.DeclinedBy ,  " & vbCrLf
s = s & "                        icm.IsDeclinedNotice ,  " & vbCrLf
s = s & "                        icm.UserName ,  " & vbCrLf
s = s & "                        icm.NoCommTotal ,  " & vbCrLf
s = s & "                        icm.ChangeFee ,  " & vbCrLf
s = s & "                        icm.IsCOLocked ,  " & vbCrLf
s = s & "                        icm.CreationDate ,  " & vbCrLf
s = s & "                        icm.IsRebateExempt ,  " & vbCrLf
s = s & "                        icm.SalesPerson ,  " & vbCrLf
s = s & "                        icm.IsContractExempt ,  " & vbCrLf
s = s & "                        isnull(icm.SumOfTotalAmount,0.00) SumOfTotalAmount  ,  " & vbCrLf
s = s & "                        icm.SumOfNetTAX ,  " & vbCrLf
s = s & "                        icm.NotTaxedTotal ,  " & vbCrLf
s = s & "                        icm.EstimateIndex ,  " & vbCrLf
s = s & "                        icm.IsApplyIncentive ,  " & vbCrLf
s = s & "                        icm.Comments ,  " & vbCrLf
s = s & "                        icm.CRMID ,  " & vbCrLf
s = s & "                       icm.CRMType  " & vbCrLf
s = s & "                FROM    @tblInventoryHomesCOMaster icm " & vbCrLf
s = s & "                       JOIN tblcustomers cust on cust.customer_no = icm.CustomerNo" & vbCrLf
s = s & "                        LEFT JOIN dbo.ChangeOrderMaster cm ON (cm.Customer_No = icm.CustomerNo  " & vbCrLf
s = s & "                                                              AND cm.Change_Order_No = icm.ChangeOrderNo)  " & vbCrLf
s = s & "                                                              OR (icm.CRMID = cm.CRMID)  " & vbCrLf
s = s & "                WHERE   (cm.Customer_No IS NULL AND cm.Change_Order_No  IS NULL)   " & vbCrLf
s = s & "      AND cm.CRMID IS NULL;   " & vbCrLf
s = s & "" & vbCrLf
s = s & "  set @sqlStatement = 'Update ChangeOrderMaster from @tblInventoryHomeCOMaster' " & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "        UPDATE  cm  " & vbCrLf
s = s & "        SET     cm.Change_Date = icm.ChangeDate ,  " & vbCrLf
s = s & "                cm.Bldr_Approved_By = icm.BldrApprovedBy ,  " & vbCrLf
s = s & "                cm.Bldr_Approved_Date = icm.BldrApprovedDate ,  " & vbCrLf
s = s & "                cm.Paid_Date = icm.PaidDate ,  " & vbCrLf
s = s & "                cm.Next_Item_No = icm.NextItemNo ,  " & vbCrLf
s = s & "                cm.Bldr_Approved = icm.IsBldrApproved ,  " & vbCrLf
s = s & "                cm.Total_Price = isnull(icm.TotalPrice,0) ,  " & vbCrLf
s = s & "                cm.Amount_Paid = isnull(icm.AmountPaid,0) ,  " & vbCrLf
s = s & "                cm.Addition_Change = icm.IsAdditionChange ,  " & vbCrLf
s = s & "                cm.Declined = icm.IsDeclined ,  " & vbCrLf
s = s & "                cm.Declined_By = icm.DeclinedBy ,  " & vbCrLf
s = s & "                cm.Declined_Notice = icm.IsDeclinedNotice ,  " & vbCrLf
s = s & "    --  cm.UserID = icm.[User]  ,  " & vbCrLf
s = s & "                cm.NoCommTotal = icm.NoCommTotal ,  " & vbCrLf
s = s & "                cm.ChargedFee = icm.ChangeFee ,  " & vbCrLf
s = s & "                cm.COLocked = icm.IsCOLocked ,  " & vbCrLf
s = s & "                cm.CreationDate = icm.CreationDate ,  " & vbCrLf
s = s & "                cm.RebateExempt = icm.IsRebateExempt ,  " & vbCrLf
s = s & "                cm.Sales_Person_ID = icm.SalesPerson ,  " & vbCrLf
s = s & "                cm.contractexempt = icm.IsContractExempt ,  " & vbCrLf
s = s & "                cm.SumOfTotalAmount =  isnull( icm.SumOfTotalAmount,0.00)  ,  " & vbCrLf
s = s & "                cm.SumOfNetTAX = icm.SumOfNetTAX ,  " & vbCrLf
s = s & "                cm.NotTaxedTotal = icm.NotTaxedTotal ,  " & vbCrLf
s = s & "                                                         " & vbCrLf
s = s & "                cm.ApplyIncentive = icm.IsApplyIncentive ,  " & vbCrLf
s = s & "                cm.comments = icm.Comments  " & vbCrLf
s = s & "        FROM    @tblInventoryHomesCOMaster icm  " & vbCrLf
s = s & "                INNER JOIN dbo.ChangeOrderMaster cm ON (cm.Customer_No = icm.CustomerNo  " & vbCrLf
s = s & "                                                       AND cm.Change_Order_No = icm.ChangeOrderNo)  " & vbCrLf
s = s & "        OR( icm.CRMID = cm.CRMID)  " & vbCrLf
s = s & "  WHERE ( cm.TSTMP <=DATEADD( HOUR, @TimeZoneUTCOffSet,icm.ModifiedDate )   " & vbCrLf
s = s & "  -- @LastPullDate   " & vbCrLf
s = s & "  OR cm.TSTMP IS NULL)  " & vbCrLf
s = s & "  AND (icm.IsBldrApproved = 1 AND cm.Bldr_Approved = 0 or icm.isDeclined = 1)  " & vbCrLf
s = s & "     --NOTE: The data transfer should not update unless the HF.Bldr_approve is 0   " & vbCrLf
s = s & "     -- (None Approved) and the CRM is approved. If the client approves in HF then the DT would not update profit.     " & vbCrLf
s = s & "     --Daryl Email 04/20/2016;  " & vbCrLf
s = s & "" & vbCrLf
s = s & " set @sqlStatement = 'Insert ChangeOrderDetails from @tblInventoryHomeCODetails' " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "        INSERT  INTO dbo.ChangeOrderDetails  " & vbCrLf
s = s & "                ( Customer_No ,  " & vbCrLf
s = s & "                  Change_Order_No ,  " & vbCrLf
s = s & "                  Bldr_Approved_By ,  " & vbCrLf
s = s & "                  Bldr_Approved_Date ,  " & vbCrLf
s = s & "                  Item_No ,  " & vbCrLf
s = s & "                  Custom ,  " & vbCrLf
s = s & "                  Require_Quote ,  " & vbCrLf
s = s & "                  [Description] ,  " & vbCrLf
s = s & "                  QTY ,  " & vbCrLf
s = s & "                  Option_Type ,  " & vbCrLf
s = s & "                  Email_Sent ,  " & vbCrLf
s = s & "                  Cost_Amount ,  " & vbCrLf
s = s & "                  Base_price ,  " & vbCrLf
s = s & "                  Rate ,  " & vbCrLf
s = s & "                  Override_Price ,  " & vbCrLf
s = s & "                  UOM ,  " & vbCrLf
s = s & "                  OPT ,  " & vbCrLf
s = s & "                  Major_Group ,  -- This is the Category  " & vbCrLf
s = s & "                  Category ,  " & vbCrLf
s = s & "                  NoCommission ,  " & vbCrLf
s = s & "                  OPTDELETED ,  " & vbCrLf
s = s & "                  EnerGuideRating ,  " & vbCrLf
s = s & "                  GreenHouseRating ,  " & vbCrLf
s = s & "                  FuelCost ,  " & vbCrLf
s = s & "                  CreationDate ,  " & vbCrLf
s = s & "                  UseTax ,  " & vbCrLf
s = s & "                  NetTax ,  " & vbCrLf
s = s & "                  GrandTotal ,  " & vbCrLf
s = s & "                  EstimateIndex ,  " & vbCrLf
s = s & "                  [Assembly] ,  " & vbCrLf
s = s & "                  SalesWorksheet ,  " & vbCrLf
s = s & "                  TAX ,  " & vbCrLf
s = s & "                  ApplyIncentive ,  " & vbCrLf
s = s & "                  IncentiveValue ,  " & vbCrLf
s = s & "                  Declined ,  " & vbCrLf
s = s & "                  EstimatorNotes ,  " & vbCrLf
s = s & "                  Bldr_Declined ,  " & vbCrLf
s = s & "                  Bldr_Declined_Reason ,  " & vbCrLf
s = s & "                  comments ,  " & vbCrLf
s = s & "                  Color ,  " & vbCrLf
s = s & "                  Location ,  " & vbCrLf
s = s & "                  Style ,  " & vbCrLf
s = s & "                  Finish ,  " & vbCrLf
s = s & "                  Other ,  " & vbCrLf
s = s & "                  ColorListID ,  " & vbCrLf
s = s & "                  StyleListID ,  " & vbCrLf
s = s & "                  FinishListID ,  " & vbCrLf
s = s & "                  OtherListID ,  " & vbCrLf
s = s & "                  CRMID ,  " & vbCrLf
s = s & "      CRMType,USERID  " & vbCrLf
s = s & "                )  " & vbCrLf
s = s & "                SELECT  icd.CustomerNo ,  " & vbCrLf
s = s & "                        icd.ChangeOrderNo ,  " & vbCrLf
s = s & "                        icd.BldrApprovedBy ,  " & vbCrLf
s = s & "                        icd.BldrApprovedDate ,  " & vbCrLf
s = s & "                        icd.ItemNo ,  " & vbCrLf
s = s & "                        icd.IsCustom ,  " & vbCrLf
s = s & "                        isnull(icd.IsRequireQuote,0) ,  " & vbCrLf
s = s & "                        icd.[Description] ,  " & vbCrLf
s = s & "                        icd.QTY ,  " & vbCrLf
s = s & "                        icd.OptionType ,  " & vbCrLf
s = s & "                        icd.IsEmailSent ,  " & vbCrLf
s = s & "                        icd.Cost ,  " & vbCrLf
s = s & "                        icd.PreTaxSellingPrice ,  " & vbCrLf
s = s & "                        icd.Rate ,  " & vbCrLf
s = s & "                        icd.OverridePrice ,  " & vbCrLf
s = s & "                        icd.UOM ,  " & vbCrLf
s = s & "                        icd.[Option] ,  " & vbCrLf
s = s & "                        icd.Category ,  " & vbCrLf
s = s & "                        icd.SubCategory ,  " & vbCrLf
s = s & "                        icd.IsNoCommission ,  " & vbCrLf
s = s & "                        icd.IsOPTDeleted ,  " & vbCrLf
s = s & "                        icd.EnerGuideRating ,  " & vbCrLf
s = s & "                        icd.GreenHouseRating ,  " & vbCrLf
s = s & "                        icd.FuelCost ,  " & vbCrLf
s = s & "                        icd.CreationDate ,  " & vbCrLf
s = s & "                        icd.IsUseTax ,  " & vbCrLf
s = s & "                        icd.NetTax ,  " & vbCrLf
s = s & "                        icd.GrandTotal ,  " & vbCrLf
s = s & "                        icd.EstimateIndex ,  " & vbCrLf
s = s & "                        icd.[Assembly] ,  " & vbCrLf
s = s & "                        icd.SalesWorksheet ,  " & vbCrLf
s = s & "                        icd.TAX ,  " & vbCrLf
s = s & "                        icd.IsApplyIncentive ,  " & vbCrLf
s = s & "                        icd.IncentiveValue ,  " & vbCrLf
s = s & "                        icd.IsDeclined ,  " & vbCrLf
s = s & "                        icd.EstimatorNotes ,  " & vbCrLf
s = s & "                        icd.IsBldrDeclined ,  " & vbCrLf
s = s & "                        icd.BldrDeclinedReason ,  " & vbCrLf
s = s & "                        icd.Comments ,  " & vbCrLf
s = s & "                        icd.Color ,                          icd.Location ,  " & vbCrLf
s = s & "                        icd.Style ,  " & vbCrLf
s = s & "                        icd.Finish ,  " & vbCrLf
s = s & "                        icd.Other ,  " & vbCrLf
s = s & "                        icd.ColorListID ,  " & vbCrLf
s = s & "                        icd.StyleListID ,  " & vbCrLf
s = s & "                        icd.FinishListID ,  " & vbCrLf
s = s & "                        icd.OtherListID ,  " & vbCrLf
s = s & "                        icd.CRMID ,  " & vbCrLf
s = s & "                       icd.CRMType," & vbCrLf
s = s & "                       icd.SalesPerson" & vbCrLf
s = s & "                FROM    @tblInventoryHomesCODetails icd  " & vbCrLf
s = s & "                       INNER JOIN ChangeOrderMaster com ON ICD.CustomerNo=COM.Customer_No   " & vbCrLf
s = s & "                       AND ICD.ChangeOrderNo=COM.Change_Order_No  " & vbCrLf
s = s & "                        LEFT JOIN dbo.ChangeOrderDetails cd ON cd.Customer_No = icd.CustomerNo  " & vbCrLf
s = s & "                        AND cd.Change_Order_No = icd.ChangeOrderNo  " & vbCrLf
s = s & "                        AND cd.CRMID = icd.CRMID  " & vbCrLf
s = s & "                       WHERE   cd.CRMID IS NULL  " & vbCrLf
s = s & "                       AND icd.IsOPTDeleted = 0;  " & vbCrLf
s = s & " " & vbCrLf
s = s & "" & vbCrLf
s = s & " set @sqlStatement = 'Update ChangeOrderDetails from @tblInventoryHomeCODetails' " & vbCrLf
s = s & " " & vbCrLf
s = s & "        UPDATE  cd  " & vbCrLf
s = s & "        SET     cd.Bldr_Approved_By = icd.BldrApprovedBy ,  " & vbCrLf
s = s & "                cd.Bldr_Approved_Date = icd.BldrApprovedDate ,  " & vbCrLf
s = s & "                cd.Item_No = icd.ItemNo ,  " & vbCrLf
s = s & "                cd.Custom = icd.IsCustom ,  " & vbCrLf
s = s & "                cd.Require_Quote = isnull(icd.IsRequireQuote,0) ,  " & vbCrLf
s = s & "                cd.[Description] = icd.[Description] ,  " & vbCrLf
s = s & "                cd.QTY = icd.QTY ,  " & vbCrLf
s = s & "                cd.Option_Type = icd.OptionType ,  " & vbCrLf
s = s & "                cd.Email_Sent = icd.IsEmailSent ,  " & vbCrLf
s = s & "                cd.Cost_Amount = icd.Cost ,  " & vbCrLf
s = s & "                cd.Base_price = icd.PreTaxSellingPrice ,  " & vbCrLf
s = s & "                cd.Rate = icd.Rate ,  " & vbCrLf
s = s & "                cd.Override_Price = icd.OverridePrice ,  " & vbCrLf
s = s & "                cd.UOM = icd.UOM ,  " & vbCrLf
s = s & "                cd.OPT = icd.[Option] ,  " & vbCrLf
s = s & "                cd.Major_Group = icd.Category ,  " & vbCrLf
s = s & "                cd.Category = icd.SubCategory ,  " & vbCrLf
s = s & "                cd.NoCommission = icd.IsNoCommission ,  " & vbCrLf
s = s & "                cd.OPTDELETED = icd.IsOPTDeleted ,  " & vbCrLf
s = s & "                cd.EnerGuideRating = icd.EnerGuideRating ,  " & vbCrLf
s = s & "                cd.GreenHouseRating = icd.GreenHouseRating ,  " & vbCrLf
s = s & "                cd.FuelCost = icd.FuelCost ,  " & vbCrLf
s = s & "                cd.CreationDate = icd.CreationDate ,  " & vbCrLf
s = s & "                cd.UseTax = icd.IsUseTax ,  " & vbCrLf
s = s & "                cd.NetTax = icd.NetTax ,  " & vbCrLf
s = s & "                cd.GrandTotal = icd.GrandTotal ,  " & vbCrLf
s = s & "                                                         " & vbCrLf
s = s & "                cd.[Assembly] = icd.[Assembly] ,  " & vbCrLf
s = s & "                cd.SalesWorksheet = icd.SalesWorksheet ,  " & vbCrLf
s = s & "                cd.TAX = icd.TAX ,  " & vbCrLf
s = s & "                cd.ApplyIncentive = icd.IsApplyIncentive ,  " & vbCrLf
s = s & "                cd.IncentiveValue = icd.IncentiveValue ,  " & vbCrLf
s = s & "                cd.Declined = icd.IsDeclined ,  " & vbCrLf
s = s & "                cd.EstimatorNotes = icd.EstimatorNotes ,  " & vbCrLf
s = s & "                cd.Bldr_Declined = icd.IsBldrDeclined ,  " & vbCrLf
s = s & "                cd.Bldr_Declined_Reason = icd.BldrDeclinedReason ,  " & vbCrLf
s = s & "                cd.comments = icd.Comments ,  " & vbCrLf
s = s & "                cd.Color = icd.Color ,  " & vbCrLf
s = s & "                cd.Location = icd.Location ,  " & vbCrLf
s = s & "                cd.Style = icd.Style ,  " & vbCrLf
s = s & "                cd.Finish = icd.Finish ,  " & vbCrLf
s = s & "                cd.Other = icd.Other ,  " & vbCrLf
s = s & "                cd.ColorListID = icd.ColorListID ,  " & vbCrLf
s = s & "                cd.StyleListID = icd.StyleListID ,  " & vbCrLf
s = s & "                cd.FinishListID = icd.FinishListID ,  " & vbCrLf
s = s & "                cd.OtherListID = icd.OtherListID ,  " & vbCrLf
s = s & "               cd.CRMType = icd.CRMType  " & vbCrLf
s = s & "        FROM    @tblInventoryHomesCODetails icd  " & vbCrLf
s = s & "               INNER JOIN ChangeOrderMaster com ON ICD.CustomerNo=COM.Customer_No   " & vbCrLf
s = s & "               AND ICD.ChangeOrderNo=COM.Change_Order_No  " & vbCrLf
s = s & "               INNER JOIN dbo.ChangeOrderDetails cd ON cd.Customer_No = icd.CustomerNo  " & vbCrLf
s = s & "               AND cd.Change_Order_No = icd.ChangeOrderNo  " & vbCrLf
s = s & "               AND cd.CRMID = icd.CRMID  " & vbCrLf
s = s & "               WHERE  (cd.TSTMP <= DATEADD( HOUR, @TimeZoneUTCOffSet,icd.ModifiedDate ) OR cd.TSTMP IS NULL)  " & vbCrLf
s = s & "  --WHERE ISNULL(icd.IsOPTDeleted,0) = 0  " & vbCrLf
s = s & "  -- Deleted options need to be updated as well to the 'Deleted' status  " & vbCrLf
s = s & "    " & vbCrLf
s = s & "" & vbCrLf
s = s & "   --Restore BuilderApproved correct value after updating CO Details.  " & vbCrLf
s = s & "  UPDATE  cm  " & vbCrLf
s = s & "        SET     cm.Bldr_Approved = icm.IsBldrApproved   " & vbCrLf
s = s & "        FROM    @tblInventoryHomesCOMaster icm  " & vbCrLf
s = s & "                INNER JOIN dbo.ChangeOrderMaster cm ON (cm.Customer_No = icm.CustomerNo  " & vbCrLf
s = s & "                                                       AND cm.Change_Order_No = icm.ChangeOrderNo)  " & vbCrLf
s = s & "                                                       OR (icm.CRMID = cm.CRMID)  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "     WHERE (icm.IsBldrApproved = 1 AND cm.Bldr_Approved = 0)  " & vbCrLf
s = s & "     --NOTE: The data transfer should not update unless the HF.Bldr_approve is 0   " & vbCrLf
s = s & "     -- (None Approved) and the CRM is approved. If the client approves in HF then the DT would not update profit.     " & vbCrLf
s = s & "     --Daryl Email 04/20/2016  " & vbCrLf
s = s & "   " & vbCrLf
s = s & "  " & vbCrLf
s = s & " -- End Inventory  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  --   -- Update Inventory Homes status when used in SC & Release Inventory Homes on Cancelled SC  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  set @sqlStatement = 'Declare SalesContract_Cursor' " & vbCrLf
s = s & "" & vbCrLf
s = s & "        DECLARE SalesContract_Cursor CURSOR  " & vbCrLf
s = s & "        FOR  " & vbCrLf
s = s & "            SELECT  IsCancelled  " & vbCrLf
s = s & "                  , ClientInventoryHomeID  " & vbCrLf
s = s & "                  , Customer_No  " & vbCrLf
s = s & "            FROM    @tblSalesContract  " & vbCrLf
s = s & "          " & vbCrLf
s = s & "        DECLARE @SpecHomeJob NVARCHAR(12)  " & vbCrLf
s = s & "  DECLARE @SpecHomeID NVARCHAR(12)  " & vbCrLf
s = s & "        DECLARE @IsCancelled BIT  " & vbCrLf
s = s & "        DECLARE @ClientInventoryHomeID NVARCHAR(12)  " & vbCrLf
s = s & "        DECLARE @Customer_No NVARCHAR(12),  " & vbCrLf
s = s & "  @BalSheetPrefix NVARCHAR(20) =NULL,  " & vbCrLf
s = s & "  @IncomePrefix  NVARCHAR(20) =NULL  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        OPEN SalesContract_Cursor;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        FETCH NEXT FROM SalesContract_Cursor INTO @IsCancelled,  " & vbCrLf
s = s & "            @ClientInventoryHomeID, @Customer_No;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        WHILE @@FETCH_STATUS = 0  " & vbCrLf
s = s & "            BEGIN  " & vbCrLf
s = s & "    -- Use a try/catch to always close and deallocate the cursor  " & vbCrLf
s = s & "                BEGIN TRY  " & vbCrLf
s = s & "      " & vbCrLf
s = s & "    -- Move JobNo from Spec Home to SC  " & vbCrLf
s = s & "                    IF @IsCancelled = 0 -- IsPurchased = true  " & vbCrLf
s = s & "                        BEGIN  " & vbCrLf
s = s & "                            SET @SpecHomeJob = ( SELECT Job_No  " & vbCrLf
s = s & "                                                 FROM   tblCustomers  " & vbCrLf
s = s & "                                                 WHERE  Customer_No = @ClientInventoryHomeID  " & vbCrLf
s = s & "                                               )  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                           IF ISNULL(@SpecHomeJob, '') != ''  " & vbCrLf
s = s & "                                BEGIN  " & vbCrLf
s = s & "       -- Set JobNo in Sales Contract  " & vbCrLf
s = s & "                                    UPDATE  tblCustomers  " & vbCrLf
s = s & "                                    SET     Job_No = @SpecHomeJob  " & vbCrLf
s = s & "                                          , webdeleted = 0  " & vbCrLf
s = s & "                                          , Sold_To_Customer = @ClientInventoryHomeID  " & vbCrLf
s = s & "                                    WHERE   Customer_No = @Customer_No  " & vbCrLf
s = s & "     " & vbCrLf
s = s & "       -- Update Spec Home status  " & vbCrLf
s = s & "                                    UPDATE  tblCustomers  " & vbCrLf
s = s & "                                    SET     Job_No = ''  " & vbCrLf
s = s & "                                          , Sold = 1  " & vbCrLf
s = s & "                                          , InventoryStatus = 2  " & vbCrLf
s = s & "                                          , WebUpdated = 0  " & vbCrLf
s = s & "                                          , NoDataChanged = 0  " & vbCrLf
s = s & "                                          , Sale_posted = 1  " & vbCrLf
s = s & "                                          , webdeleted = 1  " & vbCrLf
s = s & "                                          , Sold_To_Customer = @Customer_No  " & vbCrLf
s = s & "                                    WHERE   Customer_No = @ClientInventoryHomeID  " & vbCrLf
s = s & "                                END  " & vbCrLf
s = s & "                        END  " & vbCrLf
s = s & "    -- Release Cancelled SC Spec Homes  " & vbCrLf
s = s & "                    ELSE -- IsCancelled = 1 / IsPurchased = false  " & vbCrLf
s = s & "                        BEGIN  " & vbCrLf
s = s & "       SELECT @SpecHomeJob = Job_No, @SpecHomeID = Sold_To_Customer  " & vbCrLf
s = s & "                            FROM   tblCustomers  " & vbCrLf
s = s & "                            WHERE  Customer_No = @Customer_No;  " & vbCrLf
s = s & "                            " & vbCrLf
s = s & "                            IF /*ISNULL(@SpecHomeJob, '') != '' AND*/ ISNULL(@SpecHomeID, '') != ''  " & vbCrLf
s = s & "                                BEGIN  " & vbCrLf
s = s & "         --Update Spec Home status  " & vbCrLf
s = s & "                                    UPDATE  tblCustomers  " & vbCrLf
s = s & "                                    SET     Sold_To_Customer = NULL  " & vbCrLf
s = s & "                                           ,Sold=0" & vbCrLf
s = s & "                                           ,Job_No=@SpecHomeJob" & vbCrLf
s = s & "                                           , Bal_Sheet_Prefix = Case WHEN ISNULL(@BalSheetPrefix,'') = '' THEN Bal_Sheet_Prefix ELSE @BalSheetPrefix END  " & vbCrLf
s = s & "                                           , Income_Prefix = Case WHEN ISNULL(@IncomePrefix,'') = '' THEN Income_Prefix ELSE @IncomePrefix END  " & vbCrLf
s = s & "                                    WHERE   Customer_No = @SpecHomeID  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "         --Update Sales Contract status  " & vbCrLf
s = s & "                                    UPDATE  tblCustomers  " & vbCrLf
s = s & "                                    SET     Sold_To_Customer = NULL  " & vbCrLf
s = s & "                                           ,Job_No=NULL,Cancelled=1" & vbCrLf
s = s & "                                    WHERE   Customer_No = @Customer_No  " & vbCrLf
s = s & "" & vbCrLf
s = s & "set @sqlStatement = 'Update HFComments etc for spec home assemblies' " & vbCrLf
s = s & "                                   " & vbCrLf
s = s & "                                   update a" & vbCrLf
s = s & "                                   set a.HFComments=b.comments" & vbCrLf
s = s & "                                   from  EstimateAssemblies a" & vbCrLf
s = s & "                                   join tblcustomers c on c.customer_no = a.customer_no" & vbCrLf
s = s & "                                   join tblScheduleB b on b.Customer_No = c.sold_to_Customer and a.OptionID=b.opt and a.HFDescription=b.Description" & vbCrLf
s = s & "                                   where isnull(c.sold_to_customer,'')<>'' and c.Customer_No = @ClientInventoryHomeID  and isnull(b.comments,'')<>'' and isnull(a.HFComments,'') <>isnull(b.comments,'')" & vbCrLf
s = s & "" & vbCrLf
s = s & "set @sqlStatement = 'Update Color etc for spec home assemblies' " & vbCrLf
s = s & "                                   update a" & vbCrLf
s = s & "                                   set Color = b.Color" & vbCrLf
s = s & "                                   from  EstimateAssemblies a" & vbCrLf
s = s & "                                   join tblcustomers c on c.customer_no = a.customer_no" & vbCrLf
s = s & "                                   join tblScheduleB b on b.Customer_No = c.sold_to_Customer and a.OptionID=b.opt and a.HFDescription=b.Description" & vbCrLf
s = s & "                                   where isnull(c.sold_to_customer,'')<>'' and c.Customer_No = @ClientInventoryHomeID and isnull(b.Color,'')<>'' and isnull(a.Color,'') <> isnull(b.Color,'')" & vbCrLf
s = s & "" & vbCrLf
s = s & "set @sqlStatement = 'Update Other for spec home assemblies' " & vbCrLf
s = s & "" & vbCrLf
s = s & "                                   update a" & vbCrLf
s = s & "                                   set Other= b.other" & vbCrLf
s = s & "                                   from  EstimateAssemblies a" & vbCrLf
s = s & "                                   join tblcustomers c on c.customer_no = a.customer_no" & vbCrLf
s = s & "                                   join tblScheduleB b on b.Customer_No = c.sold_to_Customer and a.OptionID=b.opt and a.HFDescription=b.Description" & vbCrLf
s = s & "                                   where isnull(c.sold_to_customer,'')<>'' and c.Customer_No = @ClientInventoryHomeID and isnull(b.Other,'')<>'' and isnull(a.Other,'') <> isnull(b.Other,'')" & vbCrLf
s = s & "" & vbCrLf
s = s & "set @sqlStatement = 'Update Style for spec home assemblies' " & vbCrLf
s = s & "" & vbCrLf
s = s & "                                   update a" & vbCrLf
s = s & "                                   set Style= b.Style" & vbCrLf
s = s & "                                   from  EstimateAssemblies a" & vbCrLf
s = s & "                                   join tblcustomers c on c.customer_no = a.customer_no" & vbCrLf
s = s & "                                   join tblScheduleB b on b.Customer_No = c.sold_to_Customer and a.OptionID=b.opt and a.HFDescription=b.Description" & vbCrLf
s = s & "                                   where isnull(c.sold_to_customer,'')<>'' and c.Customer_No = @ClientInventoryHomeID and isnull(b.Style,'')<>'' and isnull(a.Style,'') <> isnull(b.Style,'')" & vbCrLf
s = s & "                                   " & vbCrLf
s = s & "set @sqlStatement = 'Update Finish for spec home assemblies' " & vbCrLf
s = s & "" & vbCrLf
s = s & "                                   update a" & vbCrLf
s = s & "                                   set Finish= b.Finish" & vbCrLf
s = s & "                                   from  EstimateAssemblies a" & vbCrLf
s = s & "                                   join tblcustomers c on c.customer_no = a.customer_no" & vbCrLf
s = s & "                                   join tblScheduleB b on b.Customer_No = c.sold_to_Customer and a.OptionID=b.opt and a.HFDescription=b.Description" & vbCrLf
s = s & "                                   where isnull(c.sold_to_customer,'')<>'' and c.Customer_No = @ClientInventoryHomeID and isnull(b.Finish,'')<>'' and isnull(a.Finish,'') <> isnull(b.Finish,'')" & vbCrLf
s = s & "" & vbCrLf
s = s & "set @sqlStatement = 'Update Location for spec home assemblies'                                     " & vbCrLf
s = s & "" & vbCrLf
s = s & "                                   update a" & vbCrLf
s = s & "                                   set [Location]= b.[Location]" & vbCrLf
s = s & "                                   from  EstimateAssemblies a" & vbCrLf
s = s & "                                   join tblcustomers c on c.customer_no = a.customer_no" & vbCrLf
s = s & "                                   join tblScheduleB b on b.Customer_No = c.sold_to_Customer and a.OptionID=b.opt and a.HFDescription=b.Description" & vbCrLf
s = s & "                                   where isnull(c.sold_to_customer,'')<>'' and c.Customer_No = @ClientInventoryHomeID and isnull(b.[Location],'')<>'' and isnull(a.[Location],'') <> isnull(b.[Location],'')" & vbCrLf
s = s & "                                   " & vbCrLf
s = s & "                                   " & vbCrLf
s = s & "                                END  " & vbCrLf
s = s & "                              " & vbCrLf
s = s & "                        END  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                END TRY  " & vbCrLf
s = s & "                BEGIN CATCH  " & vbCrLf
s = s & "    -- Empty catch  " & vbCrLf
s = s & "                END CATCH;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                FETCH NEXT FROM SalesContract_Cursor INTO @IsCancelled,  " & vbCrLf
s = s & "                    @ClientInventoryHomeID, @Customer_No;  " & vbCrLf
s = s & "            END;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        CLOSE SalesContract_Cursor;  " & vbCrLf
s = s & "   " & vbCrLf
s = s & "        DEALLOCATE SalesContract_Cursor;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  set @sqlStatement = 'Update Misc1 and 2 dates'                                   " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  -- Misc 1 & Misc 2  " & vbCrLf
s = s & "        UPDATE  CD  " & vbCrLf
s = s & "        SET     CD.Description = D.Description  " & vbCrLf
s = s & "              , CD.Order_index = D.OrderIndex  " & vbCrLf
s = s & "              , CD.CustDate = D.StartDate  " & vbCrLf
s = s & "              , CD.Date_Completed = D.FinishDate  " & vbCrLf
s = s & "              , CD.PercentComplete = CASE WHEN D.IsCompleted = 1 THEN 100  " & vbCrLf
s = s & "                                          ELSE CD.PercentComplete  " & vbCrLf
s = s & "                                     END  " & vbCrLf
s = s & "              , CD.Comments = D.Comments  " & vbCrLf
s = s & "              , CD.sales_view = D.IsSalesView  " & vbCrLf
s = s & "              , CD.vendor_webview = D.IsVendorWebView  " & vbCrLf
s = s & "              , CD.Web_View = D.IsContractWebView  " & vbCrLf
s = s & "     , CD.Amount = ISNULL(D.Amount, 0)  " & vbCrLf
s = s & "        FROM    @tblMiscDates D  " & vbCrLf
s = s & "                INNER JOIN dbo.Customer_Date CD ON D.CustomerNo = CD.Customer_No  " & vbCrLf
s = s & "                                                   AND D.DateField = CD.Date_Field  " & vbCrLf
s = s & "                                                   AND D.ClientDateType = CD.DateType;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "      " & vbCrLf
s = s & "set @sqlStatement = 'Insert Misc1 and 2 dates'                                     " & vbCrLf
s = s & "" & vbCrLf
s = s & "        INSERT  INTO dbo.Customer_Date  " & vbCrLf
s = s & "                ( Customer_No  " & vbCrLf
s = s & "                , Date_Field  " & vbCrLf
s = s & "                , Comments  " & vbCrLf
s = s & "                , Description  " & vbCrLf
s = s & "                , Check_Box  " & vbCrLf
s = s & "                , notice_reqd  " & vbCrLf
s = s & "                , Amount  " & vbCrLf
s = s & "                , Order_index  " & vbCrLf
s = s & "                , Notice_days  " & vbCrLf
s = s & "                , CustDate  " & vbCrLf
s = s & "                , Notice_Sent  " & vbCrLf
s = s & "                , Date_Completed  " & vbCrLf
s = s & "                , Deposit  " & vbCrLf
s = s & "                , DateType  " & vbCrLf
s = s & "                , PercentComplete  " & vbCrLf
s = s & "                , sales_view  " & vbCrLf
s = s & "                )  " & vbCrLf
s = s & "                SELECT  D.CustomerNo  " & vbCrLf
s = s & "                      , D.DateField  " & vbCrLf
s = s & "                      , D.Comments  " & vbCrLf
s = s & "                      , D.Description  " & vbCrLf
s = s & "                      , 0  " & vbCrLf
s = s & "                      , 0  " & vbCrLf
s = s & "                      , ISNULL(D.Amount, 0)  " & vbCrLf
s = s & "                      , ISNULL(D.OrderIndex, 1)  " & vbCrLf
s = s & "                      , 0  " & vbCrLf
s = s & "                      , D.StartDate  " & vbCrLf
s = s & "                      , 0  " & vbCrLf
s = s & "                      , D.FinishDate  " & vbCrLf
s = s & "                      , 0  " & vbCrLf
s = s & "                      , D.ClientDateType  " & vbCrLf
s = s & "                      , CASE WHEN D.IsCompleted = 1 THEN 100  " & vbCrLf
s = s & "                             ELSE 0  " & vbCrLf
s = s & "                        END  " & vbCrLf
s = s & "                      , 0  " & vbCrLf
s = s & "                FROM    @tblMiscDates D  " & vbCrLf
s = s & "                LEFT JOIN dbo.Customer_Date CD ON D.CustomerNo = CD.Customer_No AND D.DateField = CD.Date_Field  " & vbCrLf
s = s & "               inner join dbo.tblCustomers c on c.customer_no = D.CustomerNo" & vbCrLf
s = s & "                AND D.ClientDateType = CD.DateType  " & vbCrLf
s = s & "                WHERE   CD.Date_Field IS NULL  " & vbCrLf
s = s & "                        AND CD.Customer_No IS NULL;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  set @sqlStatement = 'Update Multifamily units '                                  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  -- Multifamily Unit Update  " & vbCrLf
s = s & "        UPDATE  mfu  " & vbCrLf
s = s & "        SET     mfu.Community = crm.Community  " & vbCrLf
s = s & "              , mfu.ParkingNo = crm.ParkingNo  " & vbCrLf
s = s & "              , mfu.[Description] = crm.[Description]  " & vbCrLf
s = s & "              , mfu.CostAmount = ISNULL(crm.CostAmount, 0)  " & vbCrLf
s = s & "              , mfu.SellingPrice = crm.SellingPrice  " & vbCrLf
s = s & "              , mfu.UnitNo = crm.UnitNo  " & vbCrLf
s = s & "              , mfu.LegalUnitNo = crm.LegalUnitNo  " & vbCrLf
s = s & "              , mfu.Comments = crm.Comments  " & vbCrLf
s = s & "              , mfu.Sold = crm.Sold  " & vbCrLf
s = s & "              , mfu.SoldToCustomer = crm.SoldToCustomer  " & vbCrLf
s = s & "              , mfu.IsParking = crm.IsParking  " & vbCrLf
s = s & "              , mfu.CommunityPhase = crm.CommunityPhase  " & vbCrLf
s = s & "              , mfu.InventoryHomeID = crm.InventoryHomeID  " & vbCrLf
s = s & "              , mfu.ItemType = crm.ItemType  " & vbCrLf
s = s & "              , mfu.CRMID = crm.CRMID  " & vbCrLf
s = s & "        FROM    dbo.tblParking mfu  " & vbCrLf
s = s & "                INNER JOIN @tblMultifamilyUnit crm ON crm.Community = mfu.Community  " & vbCrLf
s = s & "                                                      AND crm.ParkingNo = mfu.ParkingNo;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  set @sqlStatement = 'Insert Lending Companies'" & vbCrLf
s = s & "  -- Lending and Title Company  " & vbCrLf
s = s & "  INSERT  INTO dbo.Lending_Companies  " & vbCrLf
s = s & "                ( Lending_Company  " & vbCrLf
s = s & "                , Company_Name  " & vbCrLf
s = s & "                )  " & vbCrLf
s = s & "                SELECT DISTINCT  tc.TitleCompany  " & vbCrLf
s = s & "                      , tc.CompanyName  " & vbCrLf
s = s & "                FROM    @tblTitleCompany tc  " & vbCrLf
s = s & "                        LEFT JOIN dbo.Lending_Companies lc ON lc.Lending_Company = tc.TitleCompany  " & vbCrLf
s = s & "                                                                " & vbCrLf
s = s & "                WHERE   lc.Lending_Company IS NULL AND tc.IsLenderCompany = 1;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  set @sqlStatement = 'Insert Lawfirms'" & vbCrLf
s = s & "" & vbCrLf
s = s & "        INSERT  INTO dbo.Law_Firms  " & vbCrLf
s = s & "                ( Law_Firm  " & vbCrLf
s = s & "                , Firm_Name  " & vbCrLf
s = s & "          )  " & vbCrLf
s = s & "                SELECT DISTINCT tc.TitleCompany  " & vbCrLf
s = s & "                      , tc.CompanyName  " & vbCrLf
s = s & "                FROM    @tblTitleCompany tc  " & vbCrLf
s = s & "                        LEFT JOIN dbo.Law_Firms lf ON lf.Law_Firm = tc.TitleCompany  " & vbCrLf
s = s & "                                                        " & vbCrLf
s = s & "                WHERE   lf.Law_Firm IS NULL AND tc.IsLenderCompany = 0;  " & vbCrLf
s = s & " " & vbCrLf
s = s & "  set @sqlStatement = 'Insert Lending Company Contacts'" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  -- Lender Company Contact  " & vbCrLf
s = s & "  INSERT  INTO dbo.BANKCONTACTS  " & vbCrLf
s = s & "                ( Lending_Company  " & vbCrLf
s = s & "                , ContactName  " & vbCrLf
s = s & "                , WorkPhone  " & vbCrLf
s = s & "                , HomePhone  " & vbCrLf
s = s & "                , Cell  " & vbCrLf
s = s & "                , Email  " & vbCrLf
s = s & "                , BRANCHADDRESS1  " & vbCrLf
s = s & "                , BRANCHADDRESS2  " & vbCrLf
s = s & "                , BRANCHCITY  " & vbCrLf
s = s & "                , BRANCHPROVINCE  " & vbCrLf
s = s & "                , BRANCHZIP  " & vbCrLf
s = s & "                , Comments  " & vbCrLf
s = s & "          )  " & vbCrLf
s = s & "                SELECT  tcc.TitleCompany  " & vbCrLf
s = s & "                      , tcc.ContactName  " & vbCrLf
s = s & "                      , tcc.WorkPhone  " & vbCrLf
s = s & "                      , tcc.HomePhone  " & vbCrLf
s = s & "                      , tcc.CellPhone  " & vbCrLf
s = s & "                      , tcc.Email  " & vbCrLf
s = s & "                      , tcc.Address1  " & vbCrLf
s = s & "                      , tcc.Address2  " & vbCrLf
s = s & "                      , tcc.City  " & vbCrLf
s = s & "                      , tcc.Province  " & vbCrLf
s = s & "                      , tcc.Zip  " & vbCrLf
s = s & "                      , tcc.Comments  " & vbCrLf
s = s & "                FROM    @tblTitleCompanyContact tcc  " & vbCrLf
s = s & "                        LEFT JOIN dbo.BANKCONTACTS bc ON bc.ContactName = tcc.ContactName  " & vbCrLf
s = s & "                WHERE   bc.ContactName IS NULL AND tcc.IsLenderCompany = 1;  " & vbCrLf
s = s & " " & vbCrLf
s = s & " " & vbCrLf
s = s & "  set @sqlStatement = 'Insert Reators' " & vbCrLf
s = s & "" & vbCrLf
s = s & "  -- Realtors  " & vbCrLf
s = s & "        INSERT  INTO dbo.Realtors  " & vbCrLf
s = s & "                ( RealtorID  " & vbCrLf
s = s & "                , RealtorName  " & vbCrLf
s = s & "    , [Description]  " & vbCrLf
s = s & "                , Address1  " & vbCrLf
s = s & "                , Address2  " & vbCrLf
s = s & "                , City  " & vbCrLf
s = s & "                , Prov  " & vbCrLf
s = s & "                , Zip  " & vbCrLf
s = s & "                , Phone  " & vbCrLf
s = s & "                , Fax  " & vbCrLf
s = s & "                , Email  " & vbCrLf
s = s & "                , ModifiedDate  " & vbCrLf
s = s & "          )  " & vbCrLf
s = s & "                SELECT distinct crm.RealtorID  " & vbCrLf
s = s & "                      , crm.RealtorName  " & vbCrLf
s = s & "       , crm.RealtorName  " & vbCrLf
s = s & "                      , crm.Address1  " & vbCrLf
s = s & "                      , crm.Address2  " & vbCrLf
s = s & "                      , crm.City  " & vbCrLf
s = s & "                      , crm.Prov  " & vbCrLf
s = s & "                      , crm.Zip  " & vbCrLf
s = s & "                      , crm.Phone  " & vbCrLf
s = s & "                      , crm.Fax  " & vbCrLf
s = s & "                      , crm.Email  " & vbCrLf
s = s & "                      , GETDATE()  " & vbCrLf
s = s & "                FROM    @tblRealtors crm  " & vbCrLf
s = s & "                        LEFT JOIN dbo.Realtors r ON r.RealtorID = crm.RealtorID  " & vbCrLf
s = s & "                WHERE   r.RealtorID IS NULL;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  set @sqlStatement = 'Insert Models' " & vbCrLf
s = s & "" & vbCrLf
s = s & "  -- Models  " & vbCrLf
s = s & "  INSERT  INTO dbo.tblModels  " & vbCrLf
s = s & "    ( [Area]  " & vbCrLf
s = s & "    , [CommunityPhase]  " & vbCrLf
s = s & "    , [Model]  " & vbCrLf
s = s & "    , [Style]  " & vbCrLf
s = s & "    , [Series]  " & vbCrLf
s = s & "    , [Elevation]  " & vbCrLf
s = s & "    , [Description]  " & vbCrLf
s = s & "    , [ModelSize]  " & vbCrLf
s = s & "    , [NoOfBedRooms]  " & vbCrLf
s = s & "    , [NoOfBathRooms]  " & vbCrLf
s = s & "    , [Base_House]  " & vbCrLf
s = s & "    , [Cost_Amount]  " & vbCrLf
s = s & "    , [Inactive]  " & vbCrLf
s = s & "    , [max_width]  " & vbCrLf
s = s & "    , [max_length]  " & vbCrLf
s = s & "    , [UseTax]  " & vbCrLf
s = s & "    , [NetTax]  " & vbCrLf
s = s & "    , [TotalAmount]  " & vbCrLf
s = s & "    , [Comments]  " & vbCrLf
s = s & "    , [SalesWorksheet]  " & vbCrLf
s = s & "    , [Spec_Document]  " & vbCrLf
s = s & "    , [assembly]  " & vbCrLf
s = s & "    , [Brochure]  " & vbCrLf
s = s & "    , [IncentiveRetail]  " & vbCrLf
s = s & "    , [IncentiveCost]  " & vbCrLf
s = s & "    , [LastSalesWorksheet]  " & vbCrLf
s = s & "    , [Margin]  " & vbCrLf
s = s & "    , [Markup]  " & vbCrLf
s = s & "    , [DivisionID]  " & vbCrLf
s = s & "    , [CreatedBy]  " & vbCrLf
s = s & "    , [CreatedDate]  " & vbCrLf
s = s & "    , [ModifiedBy]  " & vbCrLf
s = s & "    , [ModifiedDate]  " & vbCrLf
s = s & "    )  " & vbCrLf
s = s & "    SELECT  crm.Area  " & vbCrLf
s = s & "       , crm.CommunityPhase  " & vbCrLf
s = s & "       , crm.Model  " & vbCrLf
s = s & "       , crm.Style  " & vbCrLf
s = s & "       , crm.Series  " & vbCrLf
s = s & "       , crm.Elevation  " & vbCrLf
s = s & "       , crm.Description  " & vbCrLf
s = s & "       , crm.ModelSize  " & vbCrLf
s = s & "       , crm.NoOfBedRooms  " & vbCrLf
s = s & "       , crm.NoOfBathRooms  " & vbCrLf
s = s & "       , crm.BaseHouse  " & vbCrLf
s = s & "       , crm.CostAmount  " & vbCrLf
s = s & "       , crm.Inactive  " & vbCrLf
s = s & "       , crm.MaxWidth  " & vbCrLf
s = s & "       , crm.MaxLength  " & vbCrLf
s = s & "       , crm.UseTax  " & vbCrLf
s = s & "       , crm.NetTax  " & vbCrLf
s = s & "       , isnull(crm.TotalAmount,0.00)   " & vbCrLf
s = s & "       , crm.Comments  " & vbCrLf
s = s & "       , crm.SalesWorksheet  " & vbCrLf
s = s & "       , crm.SpecDocument  " & vbCrLf
s = s & "       , crm.Assembly  " & vbCrLf
s = s & "       , crm.Brochure  " & vbCrLf
s = s & "       , crm.IncentiveRetail  " & vbCrLf
s = s & "       , crm.IncentiveCost  " & vbCrLf
s = s & "       , crm.LastSalesWorksheet  " & vbCrLf
s = s & "       , crm.Margin  " & vbCrLf
s = s & "       , crm.Markup  " & vbCrLf
s = s & "       , crm.DivisionID  " & vbCrLf
s = s & "       , crm.CreatedBy  " & vbCrLf
s = s & "       , crm.CreateDate  " & vbCrLf
s = s & "       , crm.ModifiedBy  " & vbCrLf
s = s & "       , crm.ModifiedDate  " & vbCrLf
s = s & "    FROM    @tblModels crm  " & vbCrLf
s = s & "      LEFT JOIN dbo.tblModels m ON ( crm.Model = m.Model COLLATE DATABASE_DEFAULT  " & vbCrLf
s = s & "                AND crm.Area = m.Area COLLATE DATABASE_DEFAULT  " & vbCrLf
s = s & "                AND crm.Series = m.Series COLLATE DATABASE_DEFAULT  " & vbCrLf
s = s & "                AND crm.Elevation = m.Elevation COLLATE DATABASE_DEFAULT  " & vbCrLf
s = s & "                AND crm.DivisionID = m.DivisionID  " & vbCrLf
s = s & "              )  " & vbCrLf
s = s & "    WHERE   m.seq IS NULL;  " & vbCrLf
s = s & "    " & vbCrLf
s = s & "  -- Update the entries added by the triggers as processed in CRM  " & vbCrLf
s = s & "  UPDATE  hf  " & vbCrLf
s = s & "  SET     hf.ProcessedInCRM = 1  " & vbCrLf
s = s & "  FROM    DeletedOptionsforCRM hf  " & vbCrLf
s = s & "    INNER JOIN @tblDeletedOptionsForHomefront del ON hf.CRMID = del.CRMID;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  -- Update Lot Status for Sold Lots  " & vbCrLf
s = s & "  UPDATE  crm  " & vbCrLf
s = s & "  SET     crm.Status = 'Sold'  " & vbCrLf
s = s & "  FROM    tblLotInventory crm  " & vbCrLf
s = s & "  INNER JOIN @tblSalesContract s ON crm.Lot = s.Lot AND crm.Lot_No = s.lotNo" & vbCrLf
s = s & "  Where isnull(s.Closed,0) =0  and isnull(s.IsCancelled,0) = 0;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  -- Update Lot Status that have changed in CRM  " & vbCrLf
s = s & "        UPDATE  crm  " & vbCrLf
s = s & "        SET     crm.Status = upd.LotStatusDesc  " & vbCrLf
s = s & "        FROM    tblLotInventory crm  " & vbCrLf
s = s & "                INNER JOIN @tblLotStatusUpdates upd ON crm.Lot_No = upd.LotNo  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "--  insert Customer Design Selections from CRM  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  set @sqlStatement = 'Insert CustomerDesignSelections' " & vbCrLf
s = s & "" & vbCrLf
s = s & " INSERT INTO dbo.CustomerDesignSelections  " & vbCrLf
s = s & " (  " & vbCrLf
s = s & "     CustomerSelectionID,  " & vbCrLf
s = s & "     Customer,  " & vbCrLf
s = s & "     WizardType,  " & vbCrLf
s = s & "     [Page],  " & vbCrLf
s = s & "     PageSort,  " & vbCrLf
s = s & "     Section,  " & vbCrLf
s = s & "     SectionSort,  " & vbCrLf
s = s & "     Question1,  " & vbCrLf
s = s & "     Question1Sort,  " & vbCrLf
s = s & "     Question1Response,  " & vbCrLf
s = s & "     Question2,  " & vbCrLf
s = s & "     Question2Sort,  " & vbCrLf
s = s & "     Question2Response,  " & vbCrLf
s = s & "     Question3,  " & vbCrLf
s = s & "     Question3Sort,  " & vbCrLf
s = s & "     Question3Response,  " & vbCrLf
s = s & "  ModifiedDate,  " & vbCrLf
s = s & "  Deleted  " & vbCrLf
s = s & " )  " & vbCrLf
s = s & " SELECT s.CustomerSelectionID,  " & vbCrLf
s = s & "     s.Customer,  " & vbCrLf
s = s & "     s.WizardType,  " & vbCrLf
s = s & "     s.Page,  " & vbCrLf
s = s & "     s.PageSort,  " & vbCrLf
s = s & "     s.Section,  " & vbCrLf
s = s & "     s.SectionSort,  " & vbCrLf
s = s & "     s.Question1,  " & vbCrLf
s = s & "     s.Question1Sort,  " & vbCrLf
s = s & "     s.Question1Response,  " & vbCrLf
s = s & "     s.Question2,  " & vbCrLf
s = s & "     s.Question2Sort,  " & vbCrLf
s = s & "     s.Question2Response,  " & vbCrLf
s = s & "     s.Question3,  " & vbCrLf
s = s & "     s.Question3Sort,  " & vbCrLf
s = s & "     s.Question3Response,  " & vbCrLf
s = s & "  s.ModifiedDate,  " & vbCrLf
s = s & "  s.Deleted  " & vbCrLf
s = s & "  FROM @tblCustomerDesignSelections s  " & vbCrLf
s = s & "  JOIN dbo.tblCustomers cust ON cust.Customer_No = s.Customer  " & vbCrLf
s = s & "  LEFT OUTER JOIN dbo.CustomerDesignSelections c ON c.CustomerSelectionID=s.CustomerSelectionID  " & vbCrLf
s = s & "  WHERE c.CustomerSelectionID IS NULL  " & vbCrLf
s = s & "" & vbCrLf
s = s & "   set @sqlStatement = 'Update CustomerDesignSelections'  " & vbCrLf
s = s & "  " & vbCrLf
s = s & " UPDATE  CRM  " & vbCrLf
s = s & "        SET [Page]=q.[Page],  " & vbCrLf
s = s & "     PageSort = q.PageSort,  " & vbCrLf
s = s & "     Section = q.Section,  " & vbCrLf
s = s & "     SectionSort = q.SectionSort,  " & vbCrLf
s = s & "     Question1 = q.Question1,  " & vbCrLf
s = s & "     Question1Sort = q.Question1Sort,  " & vbCrLf
s = s & "     Question1Response = q.Question1Response,  " & vbCrLf
s = s & "     Question2 = q.Question2,  " & vbCrLf
s = s & "     Question2Sort = q.Question2Sort,  " & vbCrLf
s = s & "     Question2Response = q.Question2Response,  " & vbCrLf
s = s & "     Question3 = q.Question3,  " & vbCrLf
s = s & "     Question3Sort = q.Question3Sort,  " & vbCrLf
s = s & "     Question3Response = q.Question3Response,  " & vbCrLf
s = s & "  ModifiedDate = q.ModifiedDate,  " & vbCrLf
s = s & "  Deleted = q.Deleted  " & vbCrLf
s = s & "        FROM    CustomerDesignSelections CRM  " & vbCrLf
s = s & "        INNER JOIN @tblCustomerDesignSelections q ON CRM.CustomerSelectionID = q.CustomerSelectionID  " & vbCrLf
s = s & "  WHERE CRM.ModifiedDate <=  DATEADD( HOUR, @TimeZoneUTCOffSet,q.ModifiedDate )  " & vbCrLf
s = s & "  " & vbCrLf
s = s & " " & vbCrLf
s = s & " " & vbCrLf
s = s & " " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  -- Lot Premium Inserts and Update  " & vbCrLf
s = s & "       /* DECLARE @LotPremiumInfo TABLE  " & vbCrLf
s = s & "            (  " & vbCrLf
s = s & "              [LotNo] [VARCHAR](10) NOT NULL  " & vbCrLf
s = s & "            , [LotPremiumValue] [FLOAT] NOT NULL  " & vbCrLf
s = s & "            );  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        INSERT  INTO @LotPremiumInfo  " & vbCrLf
s = s & "                ( LotNo  " & vbCrLf
s = s & "                , LotPremiumValue  " & vbCrLf
s = s & "          )  " & vbCrLf
s = s & "                SELECT  DISTINCT  " & vbCrLf
s = s & "                        sc.LotNo  -- Lot_No - varchar(10)  " & vbCrLf
s = s & "                      , sc.LotPremiumValue  -- PremiumValue - float  " & vbCrLf
s = s & "                FROM    @tblSalesContract sc  " & vbCrLf
s = s & "    WHERE ISNULL(sc.LotNo, '') != ''  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  DISTINCT  " & vbCrLf
s = s & "                        q.LotNo  -- Lot_No - varchar(10)  " & vbCrLf
s = s & "                      , q.LotPremiumValue  -- PremiumValue - float  " & vbCrLf
s = s & "                FROM    @tblQuotes q  " & vbCrLf
s = s & "    WHERE ISNULL(q.LotNo, '') != ''  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT DISTINCT  " & vbCrLf
s = s & "                        ih.LotNo  -- Lot_No - varchar(10)  " & vbCrLf
s = s & "                      , ih.PremiumAmount  -- PremiumValue - float  " & vbCrLf
s = s & "                FROM    @tblInventoryHome ih  " & vbCrLf
s = s & "    WHERE ISNULL(ih.LotNo, '') != '';  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        INSERT  INTO dbo.LotPremium  " & vbCrLf
s = s & "                ( Lot_No  " & vbCrLf
s = s & "                , PremiumID  " & vbCrLf
s = s & "                , Description  " & vbCrLf
s = s & "                , PremiumValue  " & vbCrLf
s = s & "                , PremiumOverride  " & vbCrLf
s = s & "                , UseOverride  " & vbCrLf
s = s & "                , Cost  " & vbCrLf
s = s & "                )  " & vbCrLf
s = s & "                SELECT  lotP.LotNo  -- Lot_No - varchar(10)  " & vbCrLf
s = s & "                      , 'New'  -- PremiumID - varchar(10) Hardcoded as New?  " & vbCrLf
s = s & "                      , NULL  -- Description - varchar(75)  " & vbCrLf
s = s & "                      , lotP.LotPremiumValue  -- PremiumValue - float  " & vbCrLf
s = s & "                      , 0  -- PremiumOverride - float  " & vbCrLf
s = s & "                      , 0  -- UseOverride - bit  " & vbCrLf
s = s & "                      , 0  -- Cost - float  " & vbCrLf
s = s & "                FROM    @LotPremiumInfo lotP  " & vbCrLf
s = s & "                        LEFT JOIN dbo.LotPremium lp ON lotP.LotNo = lp.Lot_No  " & vbCrLf
s = s & "                WHERE   lp.Lot_No IS NULL AND ISNULL(lotP.LotPremiumValue, 0) != 0;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        UPDATE  lp  " & vbCrLf
s = s & "        SET     PremiumValue = lotP.LotPremiumValue  " & vbCrLf
s = s & "        FROM    @LotPremiumInfo lotP  " & vbCrLf
s = s & "                INNER JOIN dbo.LotPremium lp ON lotP.LotNo = lp.Lot_No;*/  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "" & vbCrLf
s = s & "        DECLARE @tblCustomerNo AS TABLE  " & vbCrLf
s = s & "            (  " & vbCrLf
s = s & "             CustomerNo NVARCHAR(12)  " & vbCrLf
s = s & "            )  " & vbCrLf
s = s & "        DECLARE @CustomerNo NVARCHAR(12)  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        INSERT  INTO @tblCustomerNo  " & vbCrLf
s = s & "                SELECT  Customer_No  " & vbCrLf
s = s & "                FROM    @tblSalesContract  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  Customer_no  " & vbCrLf
s = s & "                FROM    @tblSalesContractAddendum  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  Customer_No  " & vbCrLf
s = s & "                FROM    @tblQuotes  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  Customer_no  " & vbCrLf
s = s & "                FROM    @tblQuoteDetails  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  Customer_No  " & vbCrLf
s = s & "                FROM    @tblChangeOrderMaster  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  Customer_No  " & vbCrLf
s = s & "                FROM    @tblChangeOrderDetails  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  CustomerNo  " & vbCrLf
s = s & "                FROM    @tblDepositDates  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  CustomerNo  " & vbCrLf
s = s & "                FROM    @tblSchedulingDates  " & vbCrLf
s = s & "                UNION  " & vbCrLf
s = s & "                SELECT  CustomerNo  " & vbCrLf
s = s & "                FROM    @tblConditionDates  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "" & vbCrLf
s = s & "        DECLARE Customer_Cursor CURSOR  " & vbCrLf
s = s & "        FOR  " & vbCrLf
s = s & "            SELECT  CustomerNo  " & vbCrLf
s = s & "            FROM    @tblCustomerNo;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        OPEN Customer_Cursor;  " & vbCrLf
s = s & "        FETCH NEXT FROM Customer_Cursor INTO @CustomerNo;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "    " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  declare @JobStatus INT,@Job varchar(12),@DivisionID int  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        WHILE @@FETCH_STATUS = 0  " & vbCrLf
s = s & "            BEGIN  " & vbCrLf
s = s & "    -- Use a try/catch to always close and deallocate the cursor  " & vbCrLf
s = s & "                BEGIN TRY  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "      " & vbCrLf
s = s & "    -- Execute the Client SPs to recalculate totals for the Customer No.  " & vbCrLf
s = s & "                                          " & vbCrLf
s = s & "     " & vbCrLf
s = s & "                                          " & vbCrLf
s = s & "" & vbCrLf
s = s & "     set @sqlStatement = 'Exec CalcEddenumTotals' " & vbCrLf
s = s & "     EXEC CalcAddendumTotals @CustomerNo  " & vbCrLf
s = s & "   set @sqlStatement = 'Exec CalcCOMasterTotals' " & vbCrLf
s = s & "     EXEC CalcCOMasterTotals @CustomerNo  " & vbCrLf
s = s & "" & vbCrLf
s = s & "  set @sqlStatement = 'Exec CustomerCalcTotals' " & vbCrLf
s = s & "     -- Added on 9-Aug-2016:  " & vbCrLf
s = s & "     EXEC CustomerCalcTotals @CustomerNo  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  ----set the customers construction_status----------------------------  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "     set @JobStatus = 0  " & vbCrLf
s = s & "     set @Job = (select job_no from tblCustomers where Customer_No = @CustomerNo)  " & vbCrLf
s = s & "     set @DivisionID = (select DivisionID from tblCustomers where Customer_No = @CustomerNo)  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "     Set @JobStatus = (select top 1 isnull(js.Job_Status,0)  " & vbCrLf
s = s & "                      from scheduletasks t  " & vbCrLf
s = s & "                      join schedule s on (s.scheduleid = t.scheduleid)  " & vbCrLf
s = s & "                      join librarytasks l on (t.LibraryTaskID = l.TaskID)  " & vbCrLf
s = s & "                      join tblJobStatus js on (js.Description = l.Description)  " & vbCrLf
s = s & "                      join tblJobs j on  (s.job_no= j.Job_no and s.DivisionID = j.DivisionID)  " & vbCrLf
s = s & "                      where  t.librarytaskid <> 0 and t.ActualEndDate is not null and s.job_no = @Job and s.DivisionID = @DivisionID  " & vbCrLf
s = s & "                      ORDER BY T.ActualEndDate DESC  " & vbCrLf
s = s & "                      )  " & vbCrLf
s = s & "     IF isnull(@JobStatus,0) <> 0   " & vbCrLf
s = s & "       update tblCustomers   " & vbCrLf
s = s & "       set Construction_Status=@JobStatus  " & vbCrLf
s = s & "       where Job_no=@Job  " & vbCrLf
s = s & "       and DivisionID=@DivisionID  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "                END TRY  " & vbCrLf
s = s & "                BEGIN CATCH  " & vbCrLf
s = s & "    -- Empty catch  " & vbCrLf
s = s & "                END CATCH;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "    FETCH NEXT FROM Customer_Cursor INTO @CustomerNo;  " & vbCrLf
s = s & "            END;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        CLOSE Customer_Cursor;  " & vbCrLf
s = s & "        DEALLOCATE Customer_Cursor;  " & vbCrLf
s = s & "       " & vbCrLf
s = s & "  COMMIT TRANSACTION" & vbCrLf
s = s & "  END TRY" & vbCrLf
s = s & "  BEGIN CATCH" & vbCrLf
s = s & "    " & vbCrLf
s = s & "      DECLARE @ErrorLine INT = ERROR_LINE();" & vbCrLf
s = s & "     DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();" & vbCrLf
s = s & "      DECLARE @ErrorSeverity INT = ERROR_SEVERITY();" & vbCrLf
s = s & "     DECLARE @ErrorState INT = ERROR_STATE();" & vbCrLf
s = s & "     DECLARE @ERRORPROCNAME NVARCHAR(4000) = ERROR_PROCEDURE();" & vbCrLf
s = s & "" & vbCrLf
s = s & "     Set @ErrorMessage= @ErrorMessage+' Proc or trigger name '+isnull( @ERRORPROCNAME,'')+ ' Line number: ' + CAST(@ErrorLine AS VARCHAR(10))" & vbCrLf
s = s & "     " & vbCrLf
s = s & "     ROLLBACK TRANSACTION" & vbCrLf
s = s & "     " & vbCrLf
s = s & "     insert into DataTransferErrorLog(DataTransferErrorID,Entity,ErrorDate,ErrorMessage) " & vbCrLf
s = s & "     select NEWID(),@sqlStatement,getdate(),@ErrorMessage" & vbCrLf
s = s & "" & vbCrLf
s = s & "     RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);" & vbCrLf
s = s & "     RETURN;" & vbCrLf
s = s & "  END CATCH" & vbCrLf
s = s & "  " & vbCrLf
s = s & "BEGIN TRY" & vbCrLf
s = s & "BEGIN TRANSACTION " & vbCrLf
s = s & "  " & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join EstimateAssemblies b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join tblScheduleB b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "and b.EstimateIndex<>0" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join ChangeOrderMaster b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join ChangeOrderDetails b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join Customer_Date b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "left outer join customer_date d on d.customer_no=c.customer_no and d.date_field = b.Date_Field" & vbCrLf
s = s & "where d.Customer_No is null and" & vbCrLf
s = s & " c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join NEW_DEPOSIT_TRANS b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join POST_DEPOSIT_TRANS b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join NEW_MORTGAGE_TRANS b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "update b" & vbCrLf
s = s & "set Customer_No = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join POST_MORTGAGE_TRANS b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "update d" & vbCrLf
s = s & "set Sold_To_Customer = c.Customer_No" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join tblCustomers b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "join tblcustomers d on d.Customer_No = b.Sold_To_Customer" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and b.Sold_To_Customer is not null" & vbCrLf
s = s & "" & vbCrLf
s = s & "Update c" & vbCrLf
s = s & "set Job_No = b.Job_No,Sold_To_Customer = b.Sold_To_Customer" & vbCrLf
s = s & ",AR_Customer_Deposit=b.AR_Customer_Deposit" & vbCrLf
s = s & ",EstimateIndex = b.EstimateIndex" & vbCrLf
s = s & "from tblcustomers c" & vbCrLf
s = s & "join tblCustomers b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "and len(c.customer_no)=11" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "   DECLARE Customer_Cursor CURSOR  " & vbCrLf
s = s & "        FOR  " & vbCrLf
s = s & "            SELECT  c.Customer_No as CustomerNo  " & vbCrLf
s = s & "            FROM    tblcustomers c" & vbCrLf
s = s & "           where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "           and len(c.customer_no)=11;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        OPEN Customer_Cursor;  " & vbCrLf
s = s & "        FETCH NEXT FROM Customer_Cursor INTO @CustomerNo;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "    " & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "    WHILE @@FETCH_STATUS = 0  " & vbCrLf
s = s & "    BEGIN  " & vbCrLf
s = s & "    -- Use a try/catch to always close and deallocate the cursor  " & vbCrLf
s = s & "   BEGIN TRY  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "      " & vbCrLf
s = s & "    -- Execute the Client SPs to recalculate totals for the Customer No.  " & vbCrLf
s = s & "" & vbCrLf
s = s & "    " & vbCrLf
s = s & "     EXEC CalcAddendumTotals @CustomerNo  " & vbCrLf
s = s & "   " & vbCrLf
s = s & "     EXEC CalcCOMasterTotals @CustomerNo  " & vbCrLf
s = s & "" & vbCrLf
s = s & "  " & vbCrLf
s = s & "     -- Added on 9-Aug-2016:  " & vbCrLf
s = s & "     EXEC CustomerCalcTotals @CustomerNo  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  ----set the customers construction_status----------------------------  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "     " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "    END TRY  " & vbCrLf
s = s & "    BEGIN CATCH  " & vbCrLf
s = s & "-- Empty catch  " & vbCrLf
s = s & "    END CATCH;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "    FETCH NEXT FROM Customer_Cursor INTO @CustomerNo;  " & vbCrLf
s = s & "   END;  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "        CLOSE Customer_Cursor;  " & vbCrLf
s = s & "        DEALLOCATE Customer_Cursor;  " & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "" & vbCrLf
s = s & "   Update b" & vbCrLf
s = s & "   set " & vbCrLf
s = s & "   Job_No = ''" & vbCrLf
s = s & "   ,crmid   = null" & vbCrLf
s = s & "   , RefNotes = cast(c.crmid   as varchar(40))" & vbCrLf
s = s & "   ,Inactive = 1" & vbCrLf
s = s & "   from tblcustomers c" & vbCrLf
s = s & "   join tblCustomers b on b.Customer_No = '19'+substring(c.customer_no,4,10)" & vbCrLf
s = s & "   where c.CRMID in (select crmid from tblCustomers where crmid is not null group by crmid having count(customer_no)>1)" & vbCrLf
s = s & "   and len(c.customer_no)=11" & vbCrLf
s = s & "  " & vbCrLf
s = s & "COMMIT TRANSACTION" & vbCrLf
s = s & "END TRY" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  begin catch" & vbCrLf
s = s & "     set @ErrorLine = ERROR_LINE();" & vbCrLf
s = s & "     set @ErrorMessage = ERROR_MESSAGE();" & vbCrLf
s = s & "      set @ErrorSeverity = ERROR_SEVERITY();" & vbCrLf
s = s & "     set @ErrorState  = ERROR_STATE();" & vbCrLf
s = s & "     set @ERRORPROCNAME = ERROR_PROCEDURE();" & vbCrLf
s = s & "" & vbCrLf
s = s & "     Set @ErrorMessage= @ErrorMessage+' Proc or trigger name '+isnull( @ERRORPROCNAME,'')+ ' Line number: ' + CAST(@ErrorLine AS VARCHAR(10))" & vbCrLf
s = s & "     " & vbCrLf
s = s & "     ROLLBACK TRANSACTION" & vbCrLf
s = s & "     " & vbCrLf
s = s & "     insert into DataTransferErrorLog(DataTransferErrorID,Entity,ErrorDate,ErrorMessage) " & vbCrLf
s = s & "     select NEWID(),@sqlStatement,getdate(),@ErrorMessage" & vbCrLf
s = s & "" & vbCrLf
s = s & "     RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);" & vbCrLf
s = s & "     RETURN;" & vbCrLf
s = s & "  end catch" & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  " & vbCrLf
s = s & "  END  " & vbCrLf
    
    DT_Set_DataFromCRM_b = s
End Function

Public Function DT_Set_DataFromCRM() As String

    DT_Set_DataFromCRM = DT_Set_DataFromCRM_a() & vbCrLf & _
                         DT_Set_DataFromCRM_b()

End Function




Public Function DT_Get_QuoteAndSCDataForCRM() As String
    Dim s As String

    s = s
    s = s & "CREATE PROCEDURE [dbo].[DT_Get_QuoteAndSCDataForCRM]" & vbCrLf
    s = s & "AS" & vbCrLf
    s = s & "    BEGIN" & vbCrLf
    s = s & "-- Get LastUpdate Dates" & vbCrLf
    s = s & "        DECLARE @tblLastUpdate TABLE" & vbCrLf
    s = s & "            (" & vbCrLf
    s = s & "              TableName NVARCHAR(50)" & vbCrLf
    s = s & "            , LastUpdateDate DATETIME" & vbCrLf
    s = s & "            );" & vbCrLf
    s = s & "       " & vbCrLf
    s = s & "        INSERT  INTO @tblLastUpdate" & vbCrLf
    s = s & "                ( TableName" & vbCrLf
    s = s & "                , LastUpdateDate" & vbCrLf
    s = s & "                )" & vbCrLf
    s = s & "                SELECT DISTINCT" & vbCrLf
    s = s & "                        REPLACE(REPLACE(OptionName, 'LastUpdate<', ''), '>', '')" & vbCrLf
    s = s & "                      , CAST (OptionValue AS DATETIME)" & vbCrLf
    s = s & "                FROM    AppOptions" & vbCrLf
    s = s & "                WHERE   OptionName IN ( 'LastUpdate<tblQuotesAndSalesContract>', 'LastUpdate<tblInventoryHomes>'," & vbCrLf
    s = s & "                                        'LastUpdate<tblInventoryHomesAddendum>', 'LastUpdate<tblChangeOrder>','LastUpdate<tblCustomerDesignSelections>'" & vbCrLf
    s = s & "                                     )" & vbCrLf
    s = s & "                        AND divisionid = 1;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       -- Table Name " & vbCrLf
    s = s & "        SELECT  'DataTransferItems' AS Name;" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       -- Quote Details and Sales Contract Addendum" & vbCrLf
    s = s & "        SELECT  s.CRMID" & vbCrLf
    s = s & "              , b.CRMID AS ParentCRMID" & vbCrLf
    s = s & "              , s.CRMType" & vbCrLf
    s = s & "              , s.Customer_No CustomerNo" & vbCrLf
    s = s & "              , 'tblScheduleB' AS EntityName" & vbCrLf
    s = s & "              , NULL AS ChangeOrderNo" & vbCrLf
    s = s & "              , s.seq AS ClientSeq" & vbCrLf
    s = s & "        FROM    tblScheduleB s" & vbCrLf
    s = s & "                INNER JOIN tblCustomers b ON s.Customer_No = b.Customer_No" & vbCrLf
    s = s & "        WHERE   s.ModifiedDate >= ( SELECT  lu.LastUpdateDate" & vbCrLf
    s = s & "                                    FROM    @tblLastUpdate lu" & vbCrLf
    s = s & "                                    WHERE   lu.TableName = 'tblQuotesAndSalesContract'" & vbCrLf
    s = s & "                                  )" & vbCrLf
    s = s & "                AND s.CRMID IS NOT NULL" & vbCrLf
    s = s & "                AND s.CRMType IS NOT NULL" & vbCrLf
    s = s & "        UNION" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       -- Inventory Homes Addendum" & vbCrLf
    s = s & "        SELECT  sb.CRMID" & vbCrLf
    s = s & "              , c.CRMID AS ParentCRMID" & vbCrLf
    s = s & "              , sb.CRMType" & vbCrLf
    s = s & "              , sb.Customer_No AS CustomerNo" & vbCrLf
    s = s & "              , 'tblScheduleB' AS EntityName" & vbCrLf
    s = s & "              , NULL AS ChangeOrderNo" & vbCrLf
    s = s & "              , sb.seq AS ClientSeq" & vbCrLf
    s = s & "        FROM    [dbo].[tblScheduleB] sb" & vbCrLf
    s = s & "                INNER JOIN tblCustomers c ON sb.Customer_No = c.Customer_No" & vbCrLf
    s = s & "        WHERE   ( c.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "                  AND c.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "                  AND c.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "                  AND ( c.Home_Selection = 'Spec Home'" & vbCrLf
    s = s & "                        OR c.Home_Selection = 'Show Home'" & vbCrLf
    s = s & "                      )" & vbCrLf
    s = s & "                )" & vbCrLf
    s = s & "                AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "                AND c.Sold = 0" & vbCrLf
    s = s & "                AND Inactive = 0" & vbCrLf
    s = s & "                AND NotAvailableforSale = 0" & vbCrLf
    s = s & "                AND ISNULL(c.Model, '') != ''" & vbCrLf
    s = s & "                AND ISNULL(c.Job_No, '') != ''" & vbCrLf
    s = s & "                AND sb.CRMID IS NOT NULL" & vbCrLf
    s = s & "                AND sb.CRMType IS NOT NULL" & vbCrLf
    s = s & "                AND sb.ModifiedDate >= ( SELECT lu.LastUpdateDate" & vbCrLf
    s = s & "                                         FROM   @tblLastUpdate lu" & vbCrLf
    s = s & "                                         WHERE  lu.TableName = 'tblInventoryHomesAddendum'" & vbCrLf
    s = s & "                                       )" & vbCrLf
    s = s & "        UNION" & vbCrLf
    s = s & "                                   " & vbCrLf
    s = s & "        -- Inventory Homes CO Details" & vbCrLf
    s = s & "        SELECT  cd.CRMID" & vbCrLf
    s = s & "              , i.CRMID AS ParentCRMID" & vbCrLf
    s = s & "              , cd.CRMType" & vbCrLf
    s = s & "              , cd.Customer_No AS CustomerNo" & vbCrLf
    s = s & "              , 'ChangeOrderDetails' AS EntityName" & vbCrLf
    s = s & "              , cd.Change_Order_No AS ChangeOrderNo" & vbCrLf
    s = s & "              , cd.seq AS ClientSeq" & vbCrLf
    s = s & "        FROM    tblCustomers i" & vbCrLf
    s = s & "                INNER JOIN ChangeOrderMaster cm ON i.Customer_No = cm.Customer_No" & vbCrLf
    s = s & "                INNER JOIN ChangeOrderDetails cd ON i.Customer_No = cd.Customer_No" & vbCrLf
    s = s & "                                                    AND cm.Change_Order_No = cd.Change_Order_No" & vbCrLf
    s = s & "        WHERE   ( i.Home_Selection != 'PreSale'" & vbCrLf
    s = s & "                  AND i.PreSale_Selection = 'PreSale'" & vbCrLf
    s = s & "                  AND i.PreSale_Selection != 'Lot Only'" & vbCrLf
    s = s & "                  AND ( i.Home_Selection = 'Spec Home'" & vbCrLf
    s = s & "                        OR i.Home_Selection = 'Show Home'" & vbCrLf
    s = s & "                      )" & vbCrLf
    s = s & "                )" & vbCrLf
    s = s & "                AND ISNULL(Sold_To_Customer, '') = ''" & vbCrLf
    s = s & "                AND i.Sold = 0" & vbCrLf
    s = s & "                AND Inactive = 0" & vbCrLf
    s = s & "                AND NotAvailableforSale = 0" & vbCrLf
    s = s & "                AND ISNULL(i.Model, '') != ''" & vbCrLf
    s = s & "                AND ISNULL(i.Job_No, '') != ''" & vbCrLf
    s = s & "                AND cd.CRMID IS NOT NULL" & vbCrLf
    s = s & "                AND cd.CRMType IS NOT NULL" & vbCrLf
    s = s & "                AND cd.TSTMP >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblInventoryHomes'" & vbCrLf
    s = s & "                                )" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        UNION" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       -- Sales Contract CO Details" & vbCrLf
    s = s & "        SELECT  COD.CRMID" & vbCrLf
    s = s & "              , CO.CRMID AS ParentCRMID" & vbCrLf
    s = s & "              , COD.CRMType" & vbCrLf
    s = s & "              , COD.Customer_No AS CustomerNo" & vbCrLf
    s = s & "              , 'ChangeOrderDetails' AS EntityName" & vbCrLf
    s = s & "              , CO.Change_Order_No AS ChangeOrderNo" & vbCrLf
    s = s & "              , COD.seq AS ClientSeq" & vbCrLf
    s = s & "        FROM    ChangeOrderDetails COD" & vbCrLf
    s = s & "                INNER JOIN ChangeOrderMaster CO ON COD.Customer_No = CO.Customer_No" & vbCrLf
    s = s & "                                                   AND COD.Change_Order_No = CO.Change_Order_No" & vbCrLf
    s = s & "        WHERE   CO.CRMID IS NOT NULL AND CO.CRMType != 4" & vbCrLf
    s = s & "                AND CO.TSTMP >= ( SELECT    lu.LastUpdateDate" & vbCrLf
    s = s & "                                  FROM      @tblLastUpdate lu" & vbCrLf
    s = s & "                                  WHERE     lu.TableName = 'tblChangeOrder'" & vbCrLf
    s = s & "                                )" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "       UNION" & vbCrLf
    s = s & "       -- Customer Design Selections" & vbCrLf
    s = s & "               SELECT  s.CustomerSelectionID CRMID" & vbCrLf
    s = s & "              , s.CustomerSelectionID AS ParentCRMID" & vbCrLf
    s = s & "              , 0 CRMType              " & vbCrLf
    s = s & "             , s.Customer CustomerNo" & vbCrLf
    s = s & "              , 'tblCustomerDesignSelections' AS EntityName" & vbCrLf
    s = s & "              , NULL AS ChangeOrderNo" & vbCrLf
    s = s & "              ,0 AS ClientSeq" & vbCrLf
    s = s & "        FROM    dbo.CustomerDesignSelections s" & vbCrLf
    s = s & "                INNER JOIN tblCustomers b ON s.Customer = b.Customer_No" & vbCrLf
    s = s & "        WHERE   s.ModifiedDate >= ( SELECT  lu.LastUpdateDate" & vbCrLf
    s = s & "                                    FROM    @tblLastUpdate lu" & vbCrLf
    s = s & "                                    WHERE   lu.TableName = 'tblCustomerDesignSelections'" & vbCrLf
    s = s & "                                  )" & vbCrLf
    s = s & "                AND s.CustomerSelectionID IS NOT NULL" & vbCrLf
    s = s & "               ;" & vbCrLf
    s = s & "                " & vbCrLf
    s = s & "    END;" & vbCrLf

    DT_Get_QuoteAndSCDataForCRM = s

End Function

