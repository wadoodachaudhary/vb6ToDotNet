namespace BlazorDemos.Pages.Grid;
public static class DictionaryExtensions
{
    public static int GetInt(this Dictionary<string, object> row, string key, int defaultValue = 0)
    {
        return row.TryGetValue(key, out var val) && 
               val != null && 
               int.TryParse(val.ToString(), out var i) 
            ? i : defaultValue;
    }

    public static int? GetNullableInt(this Dictionary<string, object> row, string key)
    {
        if (!row.TryGetValue(key, out var val) || val == null || val is DBNull)
            return null;
        
        return int.TryParse(val.ToString(), out var i) ? i : null;
    }
}