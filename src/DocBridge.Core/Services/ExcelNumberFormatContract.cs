using System.Globalization;
using System.Runtime.InteropServices;

namespace DocBridge.Core.Services;

/// <summary>
/// format_range numberFormat "General" must stay the public fixture value.
/// Korean Excel often throws 0x800A03EC on Range.NumberFormat = "General".
/// Fall back to NumberFormatLocal using Application.International
/// xlGeneralFormatName (26). Do not drop or rewrite the requested format.
/// </summary>
public static class ExcelNumberFormatContract
{
    public const string General = "General";
    public const int XlGeneralFormatName = 26;
    public const int CannotSetProperty = unchecked((int)0x800A03EC);

    public interface INumberFormatSurface
    {
        void SetNumberFormat(string format);
        void SetNumberFormatLocal(string format);
        string? LocalGeneralName { get; }
    }

    public static bool IsGeneral(string? format)
    {
        if (string.IsNullOrWhiteSpace(format))
            return false;
        if (string.Equals(format, General, StringComparison.OrdinalIgnoreCase))
            return true;
        if (string.Equals(format, "G/표준", StringComparison.OrdinalIgnoreCase))
            return true;
        if (string.Equals(format, "G/일반", StringComparison.OrdinalIgnoreCase))
            return true;
        return false;
    }

    public static bool ReadbackMatches(string? requested, string? actual) =>
        string.Equals(requested, actual, StringComparison.Ordinal) ||
        (IsGeneral(requested) && IsGeneral(actual));

    public static bool IsCannotSetNumberFormat(Exception ex) =>
        ex is COMException com && com.HResult == CannotSetProperty;

    public static void Assign(INumberFormatSurface surface, string format)
    {
        try
        {
            surface.SetNumberFormat(format);
        }
        catch (Exception ex) when (IsCannotSetNumberFormat(ex) && IsGeneral(format))
        {
            var local = surface.LocalGeneralName;
            if (string.IsNullOrWhiteSpace(local))
                throw;
            surface.SetNumberFormatLocal(local);
        }
    }

    public static string DescribeCannotSet(string range, string format) =>
        $"{range}: NumberFormat '{format}' could not be set ({CannotSetProperty.ToString("X8", CultureInfo.InvariantCulture)})";
}
