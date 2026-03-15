// MRegionalSettings.cs - Converted from VB6 MRegionalSettings.bas
// This module provides regional/locale settings functionality using .NET CultureInfo

using System.Globalization;

namespace HomeFront.Components.Pages;

public static class MRegionalSettings
{
    // ── Locale Constants (preserved from VB6 for reference) ──────────────

    public const int LOCALE_ICENTURY = 0x24;
    public const int LOCALE_ICOUNTRY = 0x5;
    public const int LOCALE_ICURRDIGITS = 0x19;
    public const int LOCALE_ICURRENCY = 0x1B;
    public const int LOCALE_IDATE = 0x21;
    public const int LOCALE_IDAYLZERO = 0x26;
    public const int LOCALE_IDEFAULTCODEPAGE = 0xB;
    public const int LOCALE_IDEFAULTCOUNTRY = 0xA;
    public const int LOCALE_IDEFAULTLANGUAGE = 0x9;
    public const int LOCALE_IDIGITS = 0x11;
    public const int LOCALE_IINTLCURRDIGITS = 0x1A;
    public const int LOCALE_ILANGUAGE = 0x1;
    public const int LOCALE_ILDATE = 0x22;
    public const int LOCALE_ILZERO = 0x12;
    public const int LOCALE_IMEASURE = 0xD;
    public const int LOCALE_IMONLZERO = 0x27;
    public const int LOCALE_INEGCURR = 0x1C;
    public const int LOCALE_INEGSEPBYSPACE = 0x57;
    public const int LOCALE_INEGSIGNPOSN = 0x53;
    public const int LOCALE_INEGSYMPRECEDES = 0x56;
    public const int LOCALE_IPOSSEPBYSPACE = 0x55;
    public const int LOCALE_IPOSSIGNPOSN = 0x52;
    public const int LOCALE_IPOSSYMPRECEDES = 0x54;
    public const int LOCALE_ITIME = 0x23;
    public const int LOCALE_ITLZERO = 0x25;
    public const long LOCALE_NOUSEROVERRIDE = 0x80000000;
    public const int LOCALE_S1159 = 0x28;
    public const int LOCALE_S2359 = 0x29;
    public const int LOCALE_SABBREVCTRYNAME = 0x7;
    public const int LOCALE_SABBREVDAYNAME1 = 0x31;
    public const int LOCALE_SABBREVDAYNAME2 = 0x32;
    public const int LOCALE_SABBREVDAYNAME3 = 0x33;
    public const int LOCALE_SABBREVDAYNAME4 = 0x34;
    public const int LOCALE_SABBREVDAYNAME5 = 0x35;
    public const int LOCALE_SABBREVDAYNAME6 = 0x36;
    public const int LOCALE_SABBREVDAYNAME7 = 0x37;
    public const int LOCALE_SABBREVLANGNAME = 0x3;
    public const int LOCALE_SABBREVMONTHNAME1 = 0x44;
    public const int LOCALE_SCOUNTRY = 0x6;
    public const int LOCALE_SCURRENCY = 0x14;
    public const int LOCALE_SDATE = 0x1D;
    public const int LOCALE_SDAYNAME1 = 0x2A;
    public const int LOCALE_SDAYNAME2 = 0x2B;
    public const int LOCALE_SDAYNAME3 = 0x2C;
    public const int LOCALE_SDAYNAME4 = 0x2D;
    public const int LOCALE_SDAYNAME5 = 0x2E;
    public const int LOCALE_SDAYNAME6 = 0x2F;
    public const int LOCALE_SDAYNAME7 = 0x30;
    public const int LOCALE_SDECIMAL = 0xE;
    public const int LOCALE_SENGCOUNTRY = 0x1002;
    public const int LOCALE_SENGLANGUAGE = 0x1001;
    public const int LOCALE_SGROUPING = 0x10;
    public const int LOCALE_SINTLSYMBOL = 0x15;
    public const int LOCALE_SLANGUAGE = 0x2;
    public const int LOCALE_SLIST = 0xC;
    public const int LOCALE_SLONGDATE = 0x20;
    public const int LOCALE_SMONDECIMALSEP = 0x16;
    public const int LOCALE_SMONGROUPING = 0x18;
    public const int LOCALE_SMONTHNAME1 = 0x38;
    public const int LOCALE_SMONTHNAME10 = 0x41;
    public const int LOCALE_SMONTHNAME11 = 0x42;
    public const int LOCALE_SMONTHNAME12 = 0x43;
    public const int LOCALE_SMONTHNAME2 = 0x39;
    public const int LOCALE_SMONTHNAME3 = 0x3A;
    public const int LOCALE_SMONTHNAME4 = 0x3B;
    public const int LOCALE_SMONTHNAME5 = 0x3C;
    public const int LOCALE_SMONTHNAME6 = 0x3D;
    public const int LOCALE_SMONTHNAME7 = 0x3E;
    public const int LOCALE_SMONTHNAME8 = 0x3F;
    public const int LOCALE_SMONTHNAME9 = 0x40;
    public const int LOCALE_SMONTHOUSANDSEP = 0x17;
    public const int LOCALE_SNATIVECTRYNAME = 0x8;
    public const int LOCALE_SNATIVEDIGITS = 0x13;
    public const int LOCALE_SNATIVELANGNAME = 0x4;
    public const int LOCALE_SNEGATIVESIGN = 0x51;
    public const int LOCALE_SPOSITIVESIGN = 0x50;
    public const int LOCALE_SSHORTDATE = 0x1F;
    public const int LOCALE_STHOUSAND = 0xF;
    public const int LOCALE_STIME = 0x1E;
    public const int LOCALE_STIMEFORMAT = 0x1003;

    /// <summary>
    /// Retrieves the regional short date format pattern for the current user's locale.
    /// Replaces the VB6 GetLocaleInfo kernel32 API call with .NET CultureInfo.
    /// </summary>
    /// <returns>The short date format pattern (e.g., "M/d/yyyy"), or empty string on failure.</returns>
    public static string GetLocale()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.DateTimeFormat.ShortDatePattern;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Retrieves the date separator for the current user's locale.
    /// </summary>
    /// <returns>The date separator string (e.g., "/").</returns>
    public static string GetDateSeparator()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.DateTimeFormat.DateSeparator;
        }
        catch
        {
            return "/";
        }
    }

    /// <summary>
    /// Retrieves the time separator for the current user's locale.
    /// </summary>
    /// <returns>The time separator string (e.g., ":").</returns>
    public static string GetTimeSeparator()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.DateTimeFormat.TimeSeparator;
        }
        catch
        {
            return ":";
        }
    }

    /// <summary>
    /// Retrieves the decimal separator for the current user's locale.
    /// </summary>
    /// <returns>The decimal separator string (e.g., ".").</returns>
    public static string GetDecimalSeparator()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.NumberFormat.NumberDecimalSeparator;
        }
        catch
        {
            return ".";
        }
    }

    /// <summary>
    /// Retrieves the thousands separator for the current user's locale.
    /// </summary>
    /// <returns>The thousands separator string (e.g., ",").</returns>
    public static string GetThousandsSeparator()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.NumberFormat.NumberGroupSeparator;
        }
        catch
        {
            return ",";
        }
    }

    /// <summary>
    /// Retrieves the currency symbol for the current user's locale.
    /// </summary>
    /// <returns>The currency symbol string (e.g., "$").</returns>
    public static string GetCurrencySymbol()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.NumberFormat.CurrencySymbol;
        }
        catch
        {
            return "$";
        }
    }

    /// <summary>
    /// Retrieves the list separator for the current user's locale.
    /// </summary>
    /// <returns>The list separator string (e.g., ",").</returns>
    public static string GetListSeparator()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.TextInfo.ListSeparator;
        }
        catch
        {
            return ",";
        }
    }

    /// <summary>
    /// Retrieves the long date format pattern for the current user's locale.
    /// </summary>
    /// <returns>The long date format pattern.</returns>
    public static string GetLongDateFormat()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.DateTimeFormat.LongDatePattern;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Retrieves the time format pattern for the current user's locale.
    /// </summary>
    /// <returns>The time format pattern.</returns>
    public static string GetTimeFormat()
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return culture.DateTimeFormat.LongTimePattern;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Retrieves the current user's locale LCID (equivalent to VB6 GetUserDefaultLCID).
    /// </summary>
    /// <returns>The locale identifier (LCID).</returns>
    public static int GetUserDefaultLCID()
    {
        try
        {
            return CultureInfo.CurrentCulture.LCID;
        }
        catch
        {
            return 1033; // Default to en-US
        }
    }

    /// <summary>
    /// Retrieves the country/region name for the current user's locale.
    /// </summary>
    /// <returns>The country/region display name.</returns>
    public static string GetCountryName()
    {
        try
        {
            var region = new RegionInfo(CultureInfo.CurrentCulture.Name);
            return region.DisplayName;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Retrieves the language name for the current user's locale.
    /// </summary>
    /// <returns>The language display name.</returns>
    public static string GetLanguageName()
    {
        try
        {
            return CultureInfo.CurrentCulture.DisplayName;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Generic locale info retrieval by LCType constant.
    /// Maps VB6 LOCALE_ constants to .NET CultureInfo properties.
    /// </summary>
    /// <param name="lcType">The locale constant (e.g., LOCALE_SSHORTDATE).</param>
    /// <returns>The locale information string.</returns>
    public static string GetLocaleInfo(int lcType)
    {
        try
        {
            var culture = CultureInfo.CurrentCulture;
            return lcType switch
            {
                LOCALE_SSHORTDATE => culture.DateTimeFormat.ShortDatePattern,
                LOCALE_SLONGDATE => culture.DateTimeFormat.LongDatePattern,
                LOCALE_SDATE => culture.DateTimeFormat.DateSeparator,
                LOCALE_STIME => culture.DateTimeFormat.TimeSeparator,
                LOCALE_STIMEFORMAT => culture.DateTimeFormat.LongTimePattern,
                LOCALE_S1159 => culture.DateTimeFormat.AMDesignator,
                LOCALE_S2359 => culture.DateTimeFormat.PMDesignator,
                LOCALE_SDECIMAL => culture.NumberFormat.NumberDecimalSeparator,
                LOCALE_STHOUSAND => culture.NumberFormat.NumberGroupSeparator,
                LOCALE_SCURRENCY => culture.NumberFormat.CurrencySymbol,
                LOCALE_SMONDECIMALSEP => culture.NumberFormat.CurrencyDecimalSeparator,
                LOCALE_SMONTHOUSANDSEP => culture.NumberFormat.CurrencyGroupSeparator,
                LOCALE_SLIST => culture.TextInfo.ListSeparator,
                LOCALE_SNATIVELANGNAME => culture.NativeName,
                LOCALE_SENGLANGUAGE => culture.EnglishName,
                LOCALE_SLANGUAGE => culture.DisplayName,
                LOCALE_SPOSITIVESIGN => culture.NumberFormat.PositiveSign,
                LOCALE_SNEGATIVESIGN => culture.NumberFormat.NegativeSign,
                LOCALE_IDIGITS => culture.NumberFormat.NumberDecimalDigits.ToString(),
                LOCALE_ICURRDIGITS => culture.NumberFormat.CurrencyDecimalDigits.ToString(),
                LOCALE_SDAYNAME1 => culture.DateTimeFormat.GetDayName(DayOfWeek.Monday),
                LOCALE_SDAYNAME2 => culture.DateTimeFormat.GetDayName(DayOfWeek.Tuesday),
                LOCALE_SDAYNAME3 => culture.DateTimeFormat.GetDayName(DayOfWeek.Wednesday),
                LOCALE_SDAYNAME4 => culture.DateTimeFormat.GetDayName(DayOfWeek.Thursday),
                LOCALE_SDAYNAME5 => culture.DateTimeFormat.GetDayName(DayOfWeek.Friday),
                LOCALE_SDAYNAME6 => culture.DateTimeFormat.GetDayName(DayOfWeek.Saturday),
                LOCALE_SDAYNAME7 => culture.DateTimeFormat.GetDayName(DayOfWeek.Sunday),
                LOCALE_SABBREVDAYNAME1 => culture.DateTimeFormat.GetAbbreviatedDayName(DayOfWeek.Monday),
                LOCALE_SABBREVDAYNAME2 => culture.DateTimeFormat.GetAbbreviatedDayName(DayOfWeek.Tuesday),
                LOCALE_SABBREVDAYNAME3 => culture.DateTimeFormat.GetAbbreviatedDayName(DayOfWeek.Wednesday),
                LOCALE_SABBREVDAYNAME4 => culture.DateTimeFormat.GetAbbreviatedDayName(DayOfWeek.Thursday),
                LOCALE_SABBREVDAYNAME5 => culture.DateTimeFormat.GetAbbreviatedDayName(DayOfWeek.Friday),
                LOCALE_SABBREVDAYNAME6 => culture.DateTimeFormat.GetAbbreviatedDayName(DayOfWeek.Saturday),
                LOCALE_SABBREVDAYNAME7 => culture.DateTimeFormat.GetAbbreviatedDayName(DayOfWeek.Sunday),
                LOCALE_SMONTHNAME1 => culture.DateTimeFormat.GetMonthName(1),
                LOCALE_SMONTHNAME2 => culture.DateTimeFormat.GetMonthName(2),
                LOCALE_SMONTHNAME3 => culture.DateTimeFormat.GetMonthName(3),
                LOCALE_SMONTHNAME4 => culture.DateTimeFormat.GetMonthName(4),
                LOCALE_SMONTHNAME5 => culture.DateTimeFormat.GetMonthName(5),
                LOCALE_SMONTHNAME6 => culture.DateTimeFormat.GetMonthName(6),
                LOCALE_SMONTHNAME7 => culture.DateTimeFormat.GetMonthName(7),
                LOCALE_SMONTHNAME8 => culture.DateTimeFormat.GetMonthName(8),
                LOCALE_SMONTHNAME9 => culture.DateTimeFormat.GetMonthName(9),
                LOCALE_SMONTHNAME10 => culture.DateTimeFormat.GetMonthName(10),
                LOCALE_SMONTHNAME11 => culture.DateTimeFormat.GetMonthName(11),
                LOCALE_SMONTHNAME12 => culture.DateTimeFormat.GetMonthName(12),
                LOCALE_SCOUNTRY => new RegionInfo(culture.Name).DisplayName,
                LOCALE_SENGCOUNTRY => new RegionInfo(culture.Name).EnglishName,
                LOCALE_SNATIVECTRYNAME => new RegionInfo(culture.Name).NativeName,
                LOCALE_SABBREVCTRYNAME => new RegionInfo(culture.Name).TwoLetterISORegionName,
                LOCALE_SINTLSYMBOL => new RegionInfo(culture.Name).ISOCurrencySymbol,
                LOCALE_ILANGUAGE => culture.LCID.ToString("X4"),
                _ => string.Empty
            };
        }
        catch
        {
            return string.Empty;
        }
    }
}
