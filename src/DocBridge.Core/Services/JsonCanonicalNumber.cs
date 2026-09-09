using System.Globalization;
using System.Numerics;
using System.Text;

namespace DocBridge.Core.Services;

/// <summary>
/// Exact JSON-number identity for token binding. Coefficient and exponent are
/// parsed as digits only; values are never converted through <see cref="decimal"/>
/// or IEEE-754. Equivalent forms such as 1, 1.0, and 1e0 collapse. Distinct
/// nonzero magnitudes, including 1e-100 versus 0 and >29-significant-digit
/// neighbors, stay distinct. Huge exponents are never expanded into zeros.
/// This is not RFC 8785.
/// </summary>
internal static class JsonCanonicalNumber
{
    public static string Normalize(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return raw ?? "null";
        return Normalize(raw.AsSpan().Trim());
    }

    public static string Normalize(ReadOnlySpan<char> text)
    {
        if (text.IsEmpty) return "null";
        if (IsNonFiniteToken(text))
            throw new ArgumentOutOfRangeException(nameof(text), text.ToString(), "JSON canonical numbers must be finite");

        var negative = false;
        if (text[0] == '+') text = text[1..];
        else if (text[0] == '-')
        {
            negative = true;
            text = text[1..];
        }

        if (text.IsEmpty) return "null";

        var expIndex = IndexOfExponentSeparator(text);
        var mantissa = text;
        var exponent = BigInteger.Zero;
        if (expIndex >= 0)
        {
            mantissa = text[..expIndex];
            if (!TryParseExponent(text[(expIndex + 1)..], out exponent))
                return PreserveUnparsed(negative, text);
        }

        if (!TryParseMantissa(mantissa, out var digits, out var fractionLength))
            return PreserveUnparsed(negative, text);

        exponent -= fractionLength;
        StripLeadingZeros(digits);
        StripTrailingZeros(digits, ref exponent);
        if (digits.Length == 0) return "0";

        var result = new StringBuilder(digits.Length + 8);
        if (negative) result.Append('-');
        result.Append(digits);
        if (!exponent.IsZero)
        {
            result.Append('e');
            result.Append(exponent.ToString(CultureInfo.InvariantCulture));
        }
        return result.ToString();
    }

    private static bool IsNonFiniteToken(ReadOnlySpan<char> text)
    {
        if (text.Equals("NaN", StringComparison.OrdinalIgnoreCase)) return true;
        if (text.Equals("Infinity", StringComparison.OrdinalIgnoreCase)) return true;
        if (text.Equals("+Infinity", StringComparison.OrdinalIgnoreCase)) return true;
        if (text.Equals("-Infinity", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static int IndexOfExponentSeparator(ReadOnlySpan<char> text)
    {
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] is 'e' or 'E') return i;
        }
        return -1;
    }

    private static bool TryParseExponent(ReadOnlySpan<char> text, out BigInteger exponent)
    {
        exponent = BigInteger.Zero;
        if (text.IsEmpty) return false;

        var negative = false;
        if (text[0] == '+') text = text[1..];
        else if (text[0] == '-')
        {
            negative = true;
            text = text[1..];
        }

        if (text.IsEmpty || !AreAsciiDigits(text)) return false;
        exponent = BigInteger.Parse(text, CultureInfo.InvariantCulture);
        if (negative) exponent = -exponent;
        return true;
    }

    private static bool TryParseMantissa(ReadOnlySpan<char> mantissa, out StringBuilder digits, out int fractionLength)
    {
        digits = new StringBuilder(mantissa.Length);
        fractionLength = 0;
        if (mantissa.IsEmpty) return false;

        var dot = mantissa.IndexOf('.');
        if (dot < 0)
        {
            if (!AreAsciiDigits(mantissa)) return false;
            digits.Append(mantissa);
            return true;
        }

        var integer = mantissa[..dot];
        var fraction = mantissa[(dot + 1)..];
        if (integer.Length == 0 && fraction.Length == 0) return false;
        if (integer.Length > 0 && !AreAsciiDigits(integer)) return false;
        if (fraction.Length > 0 && !AreAsciiDigits(fraction)) return false;
        digits.Append(integer);
        digits.Append(fraction);
        fractionLength = fraction.Length;
        return true;
    }

    private static bool AreAsciiDigits(ReadOnlySpan<char> text)
    {
        for (var i = 0; i < text.Length; i++)
        {
            if (!char.IsAsciiDigit(text[i])) return false;
        }
        return true;
    }

    private static void StripLeadingZeros(StringBuilder digits)
    {
        var start = 0;
        while (start < digits.Length && digits[start] == '0') start++;
        if (start > 0) digits.Remove(0, start);
    }

    private static void StripTrailingZeros(StringBuilder digits, ref BigInteger exponent)
    {
        var end = digits.Length;
        while (end > 0 && digits[end - 1] == '0')
        {
            end--;
            exponent++;
        }
        if (end < digits.Length) digits.Length = end;
    }

    private static string PreserveUnparsed(bool negative, ReadOnlySpan<char> text)
    {
        var raw = text.ToString();
        return negative ? "-" + raw : raw;
    }
}
