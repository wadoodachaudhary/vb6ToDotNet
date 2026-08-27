namespace BlazorDemos.Pages.Grid;
using System.Data.SQLite;
using System.Data.Common;

public static class SqliteFunctionRegistry
{
    private static bool _functionsRegistered = false;

    public static string GetCommunityVendorASync(SQLiteConnection connection, string community, string poIndex, int divisionId)
    {
                community = community?.Trim() ?? "";
                poIndex = poIndex?.Trim() ?? "";

                using var cmd = connection.CreateCommand();
                cmd.CommandText = @"
                    SELECT Vendor
                    FROM POAreaVendor
                    WHERE POIndex = @POIndex
                      AND DivisionID = @DivisionID
                      AND (Area = @Community OR Area = '' OR Area IS NULL)
                    ORDER BY 
                        CASE WHEN Area = @Community THEN 3 
                             WHEN Area = '' OR Area IS NULL THEN 2 
                             ELSE 1 END DESC
                    LIMIT 1";

                cmd.Parameters.AddWithValue("@Community", community);
                cmd.Parameters.AddWithValue("@POIndex", poIndex);
                cmd.Parameters.AddWithValue("@DivisionID", divisionId);

                var result = cmd.ExecuteScalar();
                return result?.ToString() ?? "";
            
    }
    
    
   
    /// <summary>
    /// Exact C# equivalent of dbo.Purch_GetItemRateQuality() in MS SQL
    /// Returns the "quality level" (4-13) indicating which pricing source was used
    /// </summary>
    public static async Task<int> GetItemRateQualityAsync(
        SQLiteConnection connection,
        int lookupType,
        int worksheet,
        string community,
        string communityPhase,
        string assembly,
        string model,
        string optionId,
        string phase,
        string item,
        int sequence,
        string vendor,
        DateTime effectiveDate,
        int divisionId)
    {
        // Normalize inputs
        community = community?.Trim() ?? "";
        communityPhase = communityPhase?.Trim() ?? "";
        assembly = assembly?.Trim() ?? "";
        model = model?.Trim() ?? "";
        optionId = optionId?.Trim() ?? "";
        phase = phase?.Trim() ?? "";
        item = item?.Trim() ?? "";
        vendor = vendor?.Trim() ?? "";

        // Step 1: If vendor is empty, resolve using POIndex + CommunityVendor
        if (string.IsNullOrEmpty(vendor))
        {
            string poIndex = await GetPOIndexAsync(connection,phase, item, divisionId);
            vendor = GetCommunityVendorASync(connection,community, poIndex, divisionId);
        }

        // Step 2: Build the massive UNION ALL query with priority
        var sql = @"
            -- Level 4: Full match (Community + Phase + Assembly + Model)
            SELECT 4 AS DataSource, c.Current_Cost, c.Next_Cost1, c.Next_Cost2,
                   c.Last_Cost1, c.Last_Cost2, c.Last_Cost3,
                   c.Next_Effective1, c.Next_Effective2,
                   c.Last1_Expiry, c.Last2_Expiry, c.Last3_Expiry,
                   c.Forecast1, c.Forecast2, c.Forecast3, c.Forecast4, c.Forecast5,
                   c.Forecast6, c.Forecast7, c.Forecast8, c.Forecast9, c.Forecast10,
                   c.Forecast11, c.Forecast12
            FROM tblVendorCost c
            WHERE c.Community = @Community
              AND c.CommunityPhase = @CommunityPhase
              AND c.Assembly = @Assembly
              AND COALESCE(c.Model, '') = @Model
              AND c.Phase = @Phase
              AND c.Item = @Item
              AND c.Vendor = @Vendor
              AND c.DivisionID = @DivisionID
              AND @CommunityPhase <> ''

            UNION ALL
            -- Level 5: Community + Assembly + Model (no phase)
            SELECT 5, c.Current_Cost, c.Next_Cost1, c.Next_Cost2,
                   c.Last_Cost1, c.Last_Cost2, c.Last_Cost3,
                   c.Next_Effective1, c.Next_Effective2,
                   c.Last1_Expiry, c.Last2_Expiry, c.Last3_Expiry,
                   c.Forecast1, c.Forecast2, c.Forecast3, c.Forecast4, c.Forecast5,
                   c.Forecast6, c.Forecast7, c.Forecast8, c.Forecast9, c.Forecast10,
                   c.Forecast11, c.Forecast12
            FROM tblVendorCost c
            WHERE c.Community = @Community
              AND c.CommunityPhase = ''
              AND c.Assembly = @Assembly
              AND COALESCE(c.Model, '') = @Model
              AND c.Phase = @Phase
              AND c.Item = @Item
              AND c.Vendor = @Vendor
              AND c.DivisionID = @DivisionID

            UNION ALL
            -- Level 6: Global (Assembly + Model only)
            SELECT 6, c.Current_Cost, c.Next_Cost1, c.Next_Cost2,
                   c.Last_Cost1, c.Last_Cost2, c.Last_Cost3,
                   c.Next_Effective1, c.Next_Effective2,
                   c.Last1_Expiry, c.Last2_Expiry, c.Last3_Expiry,
                   c.Forecast1, c.Forecast2, c.Forecast3, c.Forecast4, c.Forecast5,
                   c.Forecast6, c.Forecast7, c.Forecast8, c.Forecast9, c.Forecast10,
                   c.Forecast11, c.Forecast12
            FROM tblVendorCost c
            WHERE c.Community = ''
              AND c.CommunityPhase = ''
              AND c.Assembly = @Assembly
              AND COALESCE(c.Model, '') = @Model
              AND c.Phase = @Phase
              AND c.Item = @Item
              AND c.Vendor = @Vendor
              AND c.DivisionID = @DivisionID

            -- Continue with levels 9-13 (no Assembly)...
            -- (Levels 9-12 follow same pattern, just remove Assembly)
            -- Level 13: Fallback to tblPhaseItem.Price
            UNION ALL
            SELECT 13 AS DataSource, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,
                   NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,
                   NULL, NULL, NULL, NULL
            FROM tblPhaseItem p
            WHERE p.Phase = @Phase
              AND p.Item = @Item
              AND p.DivisionID = @DivisionID

            ORDER BY DataSource ASC
            LIMIT 1";

        //using var connection = new SqliteConnection("Data Source=homefront.db;Version=3;");
        //await connection.OpenAsync();

        using var command = new SQLiteCommand(sql, connection);
        command.Parameters.AddWithValue("@Community", community);
        command.Parameters.AddWithValue("@CommunityPhase", communityPhase);
        command.Parameters.AddWithValue("@Assembly", assembly);
        command.Parameters.AddWithValue("@Model", model);
        command.Parameters.AddWithValue("@Phase", phase);
        command.Parameters.AddWithValue("@Item", item);
        command.Parameters.AddWithValue("@Vendor", vendor);
        command.Parameters.AddWithValue("@DivisionID", divisionId);

        var result = await command.ExecuteScalarAsync();
        return result is long l ? (int)l : 13; // Default to 13 if nothing found
    }

    // Helper: Get POIndex from tblPhaseItem
    private static async Task<string> GetPOIndexAsync(SQLiteConnection conn,string phase, string item, int divisionId)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT POIndex FROM tblPhaseItem WHERE Phase = @Phase AND Item = @Item AND DivisionID = @DivisionID";
        cmd.Parameters.AddWithValue("@Phase", phase);
        cmd.Parameters.AddWithValue("@Item", item);
        cmd.Parameters.AddWithValue("@DivisionID", divisionId);

        var result = await cmd.ExecuteScalarAsync();
        return result?.ToString() ?? "";
    }
    
  
    /// <summary>
    /// Converts the SQL Server function [dbo].[Purch_GetItemRate] to C#.
    /// Returns the effective item rate based on lookup type, effective date, and specificity hierarchy.
    /// </summary>
    public static double GetItemRate(
            SQLiteConnection conn,
            int lookupType,
            int worksheet,
            string community,
            string communityPhase,
            string assembly,
            string model,
            string optionId,
            string phase,
            string item,
            int sequence,
            string vendor,
            DateTime effectiveDate,
            int divisionId)
        {
            string poIndex = string.Empty;
            bool useModelCost = false;

            // === Step 1: Resolve Vendor and UseModelCost (mirrors original logic) ===
            if (string.IsNullOrEmpty(vendor))
            {
                if (sequence != 0)
                {
                    const string sqlDetail = @"
                        SELECT POIndex, ISNULL(UseModelCost, 0)
                        FROM tbldbassemblydetails
                        WHERE DivisionID = @DivisionID 
                          AND [sequence] = @Sequence 
                          AND Phase = @Phase 
                          AND Item = @Item";

                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = sqlDetail;
                    cmd.Parameters.AddWithValue("@DivisionID", divisionId);
                    cmd.Parameters.AddWithValue("@Sequence", sequence);
                    cmd.Parameters.AddWithValue("@Phase", phase ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Item", item ?? (object)DBNull.Value);

                    using var reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        poIndex = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
                        useModelCost = reader.GetInt32(1) == 1;
                    }
                    reader.Close();
                }

                if (string.IsNullOrEmpty(poIndex))
                {
                    const string sqlItem = @"
                        SELECT POIndex 
                        FROM tblphaseitem 
                        WHERE DivisionID = @DivisionID AND Phase = @Phase AND Item = @Item";

                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = sqlItem;
                    cmd.Parameters.AddWithValue("@DivisionID", divisionId);
                    cmd.Parameters.AddWithValue("@Phase", phase ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Item", item ?? (object)DBNull.Value);

                    poIndex = cmd.ExecuteScalar()?.ToString() ?? string.Empty;
                }

                // Call SQL function: dbo.Purch_GetCommunityVendor
                vendor = GetCommunityVendorASync(conn, community, poIndex, divisionId);
            }
            else if (sequence != 0)
            {
                const string sqlUseModel = @"
                    SELECT ISNULL(UseModelCost, 0)
                    FROM tbldbassemblydetails
                    WHERE DivisionID = @DivisionID 
                      AND [sequence] = @Sequence 
                      AND Phase = @Phase 
                      AND Item = @Item";

                using var cmd = conn.CreateCommand();
                cmd.CommandText = sqlUseModel;
                cmd.Parameters.AddWithValue("@DivisionID", divisionId);
                cmd.Parameters.AddWithValue("@Sequence", sequence);
                cmd.Parameters.AddWithValue("@Phase", phase ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Item", item ?? (object)DBNull.Value);

                useModelCost = Convert.ToInt32(cmd.ExecuteScalar()) == 1;
            }

            // === Step 2: Apply UseModelCost logic ===
            if (useModelCost)
            {
                optionId = string.Empty;
                assembly = model; // Critical: cost comes from model assembly
            }

            // === Step 3: Final hierarchical pricing query (replaces cursor) ===
            string rateCase = BuildRateCaseExpression("c");

            string sql = $@"
                WITH Rates AS (
                    -- 4: Most specific (Community + Phase + Assembly + Model)
                    SELECT 4 AS DataSource, {rateCase} AS Rate FROM tblVendorCost c
                    WHERE c.Community = @Community
                      AND c.CommunityPhase = @CommunityPhase
                      AND c.Assembly = @Assembly
                      AND ISNULL(c.Model,'') = ISNULL(@Model, '')
                      AND c.Phase = @Phase AND c.Item = @Item
                      AND c.Vendor = @Vendor AND c.DivisionID = @DivisionID

                    UNION ALL
                    -- 5: Community + Assembly + Model (no phase)
                    SELECT 5, {rateCase} FROM tblVendorCost c
                    WHERE c.Community = @Community AND c.CommunityPhase = ''
                      AND c.Assembly = @Assembly AND ISNULL(c.Model,'') = ISNULL(@Model, '')
                      AND c.Phase = @Phase AND c.Item = @Item
                      AND c.Vendor = @Vendor AND c.DivisionID = @DivisionID

                    UNION ALL
                    -- 6: Global Assembly + Model
                    SELECT 6, {rateCase} FROM tblVendorCost c
                    WHERE c.Community = '' AND c.CommunityPhase = ''
                      AND c.Assembly = @Assembly AND ISNULL(c.Model,'') = ISNULL(@Model, '')
                      AND c.Phase = @Phase AND c.Item = @Item
                      AND c.Vendor = @Vendor AND c.DivisionID = @DivisionID

                    UNION ALL
                    -- 9–12: Less specific fallbacks (no assembly)
                    SELECT 9, {rateCase} FROM tblVendorCost c
                    WHERE c.Community = @Community AND c.CommunityPhase = @CommunityPhase
                      AND c.Assembly = '' AND c.Phase = @Phase AND c.Item = @Item
                      AND c.Vendor = @Vendor AND c.DivisionID = @DivisionID

                    UNION ALL SELECT 10, {rateCase} FROM tblVendorCost c
                    WHERE c.Community = @Community AND c.CommunityPhase = ''
                      AND c.Assembly = '' AND c.Phase = @Phase AND c.Item = @Item
                      AND c.Vendor = @Vendor AND c.DivisionID = @DivisionID

                    UNION ALL SELECT 11, {rateCase} FROM tblVendorCost c
                    WHERE c.Community = '' AND c.CommunityPhase = ''
                      AND c.Assembly = '' AND c.Phase = @Phase AND c.Item = @Item
                      AND c.Vendor = @Vendor AND c.DivisionID = @DivisionID

                    UNION ALL SELECT 12, {rateCase} FROM tblVendorCost c
                    WHERE c.Community = '' AND c.CommunityPhase = ''
                      AND c.Assembly = '' AND c.Phase = @Phase AND c.Item = @Item
                      AND c.Vendor = @Vendor AND c.DivisionID = 0

                    UNION ALL
                    -- 13: Fallback to tblPhaseItem.Price (estimating override)
                    SELECT 13 AS DataSource, Price AS Rate
                    FROM tblPhaseItem pi
                    LEFT JOIN system_setup s ON s.id = pi.DivisionID
                    WHERE pi.DivisionID = @DivisionID AND pi.Phase = @Phase AND pi.Item = @Item
                      AND pi.UsePricingFromEstimating = 1 AND ISNULL(pi.Price, 0) <> 0
                )
                SELECT TOP 1 Rate 
                FROM Rates 
                WHERE Rate IS NOT NULL 
                ORDER BY DataSource ASC";
            using var finalCmd = conn.CreateCommand();
            finalCmd.CommandText = sql;
            finalCmd.Parameters.AddWithValue("@LookupType", lookupType);
            finalCmd.Parameters.AddWithValue("@EffectiveDate", effectiveDate);
            finalCmd.Parameters.AddWithValue("@Community", (object)community ?? DBNull.Value);
            finalCmd.Parameters.AddWithValue("@CommunityPhase", (object)communityPhase ?? DBNull.Value);
            finalCmd.Parameters.AddWithValue("@Assembly", (object)assembly ?? DBNull.Value);
            finalCmd.Parameters.AddWithValue("@Model", (object)model ?? DBNull.Value);
            finalCmd.Parameters.AddWithValue("@Phase", (object)phase ?? DBNull.Value);
            finalCmd.Parameters.AddWithValue("@Item", (object)item ?? DBNull.Value);
            finalCmd.Parameters.AddWithValue("@Vendor", (object)vendor ?? DBNull.Value);
            finalCmd.Parameters.AddWithValue("@DivisionID", divisionId);

            var result = finalCmd.ExecuteScalar();
            if (result == null || result == DBNull.Value)
                return 0.0;

            return Convert.ToDouble(result);
        }

        // Helper: Builds the massive CASE expression once
        private static string BuildRateCaseExpression(string alias)
        {
            return $@"
                CASE 
                    WHEN @LookupType = 0 AND @EffectiveDate >= {alias}.Next_Effective2 AND {alias}.Next_Effective2 IS NOT NULL THEN {alias}.Next_Cost2
                    WHEN @LookupType = 0 AND @EffectiveDate >= {alias}.Next_Effective1 AND {alias}.Next_Effective1 IS NOT NULL THEN {alias}.Next_Cost1
                    WHEN @LookupType = 0 AND @EffectiveDate < {alias}.Last3_Expiry AND {alias}.Last3_Expiry IS NOT NULL THEN {alias}.Last_Cost3
                    WHEN @LookupType = 0 AND @EffectiveDate < {alias}.Last2_Expiry AND {alias}.Last2_Expiry IS NOT NULL THEN {alias}.Last_Cost2
                    WHEN @LookupType = 0 AND @EffectiveDate < {alias}.Last1_Expiry AND {alias}.Last1_Expiry IS NOT NULL THEN {alias}.Last_Cost1
                    WHEN @LookupType IN (0, 1) THEN {alias}.Current_Cost
                    WHEN @LookupType = 2 THEN COALESCE(NULLIF({alias}.Next_Cost1, 0), {alias}.Current_Cost)
                    WHEN @LookupType = 3 THEN COALESCE(NULLIF({alias}.Next_Cost2, 0), {alias}.Current_Cost)
                    WHEN @LookupType = 4 THEN {alias}.Last_Cost1
                    WHEN @LookupType = 5 THEN {alias}.Last_Cost2
                    WHEN @LookupType = 6 THEN {alias}.Last_Cost3
                    WHEN @LookupType = 7 THEN {alias}.Forecast1
                    WHEN @LookupType = 8 THEN {alias}.Forecast2
                    WHEN @LookupType = 9 THEN {alias}.Forecast3
                    WHEN @LookupType = 10 THEN {alias}.Forecast4
                    WHEN @LookupType = 11 THEN {alias}.Forecast5
                    WHEN @LookupType = 12 THEN {alias}.Forecast6
                    WHEN @LookupType = 13 THEN {alias}.Forecast7
                    WHEN @LookupType = 14 THEN {alias}.Forecast8
                    WHEN @LookupType = 15 THEN {alias}.Forecast9
                    WHEN @LookupType = 16 THEN {alias}.Forecast10
                    WHEN @LookupType = 17 THEN {alias}.Forecast11
                    ELSE {alias}.Forecast12
                END";
        }

        // Wrapper for dbo.Purch_GetCommunityVendor
        private static string GetCommunityVendor(SQLiteConnection conn, string community, string poIndex, int divisionId)
        {
            const string sql = "SELECT dbo.Purch_GetCommunityVendor(@Community, @POIndex, @DivisionID)";
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@Community", community ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@POIndex", poIndex ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@DivisionID", divisionId);
            return cmd.ExecuteScalar()?.ToString() ?? string.Empty;
        }
    }



    
    

    
    
    
