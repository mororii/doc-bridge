using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class JsonCanonicalTests
{
    [Theory]
    [InlineData("1", "1.0", "1e0")]
    [InlineData("1", "1E0", "1.00e+0")]
    [InlineData("10", "1e1", "1.0e1")]
    [InlineData("100", "1e2", "10e1")]
    [InlineData("123", "123.0", "1.23e2")]
    [InlineData("9007199254740993", "9007199254740993.0", "9.007199254740993e15")]
    public void Equivalent_numeric_forms_share_one_canonical_identity(string a, string b, string c)
    {
        Assert.Equal(Normalize(a), Normalize(b));
        Assert.Equal(Normalize(a), Normalize(c));
        Assert.Equal(CanonicalNumber(a), CanonicalNumber(b));
        Assert.Equal(CanonicalNumber(a), CanonicalNumber(c));
        Assert.Equal(Hash(NumberOp(a)), Hash(NumberOp(b)));
    }

    [Fact]
    public void Tiny_exponent_does_not_underflow_to_zero()
    {
        Assert.NotEqual(Normalize("0"), Normalize("1e-100"));
        Assert.NotEqual(CanonicalNumber("0"), CanonicalNumber("1e-100"));
        Assert.Equal("1e-100", Normalize("1e-100"));
        Assert.Equal("1e-100", Normalize("10e-101"));
        Assert.Equal("1e-100", Normalize("0.1e-99"));
        Assert.NotEqual(Hash(NumberOp("0")), Hash(NumberOp("1e-100")));
    }

    [Fact]
    public void More_than_29_significant_digits_stay_exact_and_adjacent_values_differ()
    {
        const string almostOne = "1.00000000000000000000000000001";
        const string one = "1";
        var neighborA = "1." + new string('0', 40) + "1";
        var neighborB = "1." + new string('0', 40) + "2";

        Assert.NotEqual(Normalize(one), Normalize(almostOne));
        Assert.NotEqual(CanonicalNumber(one), CanonicalNumber(almostOne));
        Assert.NotEqual(Normalize(neighborA), Normalize(neighborB));
        Assert.NotEqual(CanonicalNumber(neighborA), CanonicalNumber(neighborB));
        Assert.Contains("1", Normalize(almostOne));
        Assert.NotEqual("1", Normalize(almostOne));
        Assert.NotEqual(Hash(NumberOp(one)), Hash(NumberOp(almostOne)));
        Assert.NotEqual(Hash(NumberOp(neighborA)), Hash(NumberOp(neighborB)));
    }

    [Fact]
    public void Huge_exponents_are_not_expanded_and_stay_distinct()
    {
        Assert.Equal("1e9999", Normalize("1e9999"));
        Assert.Equal("1e9999", Normalize("1E+09999"));
        Assert.Equal("1e-9999", Normalize("1e-9999"));
        Assert.NotEqual(Normalize("1e9999"), Normalize("1e9998"));
        Assert.NotEqual(Normalize("1e-9999"), Normalize("1e-9998"));
        Assert.NotEqual(Normalize("1e-9999"), Normalize("0"));
        Assert.DoesNotContain("00000", Normalize("1e9999"));
        Assert.True(Normalize("1e9999").Length < 20);

        AssertParsedHugeExponentIdentity("1e100", "1E+100", "10e99", "1e99");
        AssertParsedHugeExponentIdentity("1e9999", "1E+09999", "10e9998", "1e9998");
    }

    [Theory]
    [InlineData("-0")]
    [InlineData("-0.0")]
    [InlineData("-0e10")]
    [InlineData("0e-10")]
    [InlineData("0.000")]
    public void Signed_and_exponent_zero_forms_collapse_to_zero(string raw)
        => Assert.Equal("0", Normalize(raw));

    [Fact]
    public void Negative_zero_is_not_a_nonzero_magnitude()
    {
        Assert.Equal(CanonicalNumber("0"), CanonicalNumber("-0"));
        Assert.Equal(Hash(NumberOp("0")), Hash(NumberOp("-0.0")));
        Assert.NotEqual(CanonicalNumber("-1e-100"), CanonicalNumber("0"));
    }

    [Fact]
    public void Large_integers_keep_identity_without_double_rounding()
    {
        const string exact = "9007199254740993";
        const string lost = "9007199254740992";
        Assert.NotEqual(Normalize(exact), Normalize(lost));
        Assert.Equal(exact, Normalize(exact));
        Assert.Equal(exact, Normalize("9007199254740993.0"));
        Assert.NotEqual(CanonicalNumber(exact), CanonicalNumber(lost));
        Assert.NotEqual(Hash(NumberOp(exact)), Hash(NumberOp(lost)));
    }

    [Fact]
    public void Object_key_order_is_sorted_but_arrays_nulls_and_strings_are_preserved()
    {
        var left = JsonNode.Parse("""{"b":1,"a":{"z":true,"m":null}}""")!;
        var right = JsonNode.Parse("""{"a":{"m":null,"z":true},"b":1}""")!;
        Assert.Equal(Json.Canonical(left), Json.Canonical(right));
        Assert.Equal(Hash(left.AsObject()), Hash(right.AsObject()));

        Assert.NotEqual(Json.Canonical(JsonNode.Parse("""{"a":null}""")), Json.Canonical(JsonNode.Parse("""{"a":""}""")));
        Assert.NotEqual(Json.Canonical(JsonNode.Parse("""{"a":null}""")), Json.Canonical(JsonNode.Parse("""{}""")));
        Assert.NotEqual(Json.Canonical(JsonNode.Parse("[1,2]")), Json.Canonical(JsonNode.Parse("[2,1]")));
        Assert.NotEqual(Json.Canonical(JsonNode.Parse("""{"a":"가"}""")), Json.Canonical(JsonNode.Parse("""{"a":"나"}""")));
        Assert.Contains("−", Json.Canonical(JsonValue.Create("−")));
        Assert.Contains("㎜", Json.Canonical(JsonValue.Create("㎜")));
    }

    [Fact]
    public void DateTime_json_value_uses_json_string_representation_when_TryGetValue_string_fails()
    {
        var when = new DateTime(2026, 9, 9, 4, 5, 6, DateTimeKind.Utc);
        var value = JsonValue.Create(when);
        Assert.False(value.TryGetValue<string>(out _));
        var canonical = Json.Canonical(value);
        Assert.StartsWith("\"", canonical);
        Assert.Equal(value.ToJsonString(Json.Compact), canonical);

        var offset = JsonValue.Create(new DateTimeOffset(when, TimeSpan.Zero));
        Assert.False(offset.TryGetValue<string>(out _));
        Assert.Equal(offset.ToJsonString(Json.Compact), Json.Canonical(offset));

        var guid = Guid.Parse("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
        var guidValue = JsonValue.Create(guid);
        Assert.Equal(guidValue.ToJsonString(Json.Compact), Json.Canonical(guidValue));
    }

    [Fact]
    public void Managed_numbers_match_their_json_serialized_roundtrip()
    {
        AssertManagedMatchesSerialized(JsonValue.Create(0.1d));
        AssertManagedMatchesSerialized(JsonValue.Create(0.1f));
        AssertManagedMatchesSerialized(JsonValue.Create(1.2345d));
        AssertManagedMatchesSerialized(JsonValue.Create(1.2345f));
        AssertManagedMatchesSerialized(JsonValue.Create(double.Epsilon));
        AssertManagedMatchesSerialized(JsonValue.Create(float.Epsilon));
        AssertManagedMatchesSerialized(JsonValue.Create(1));
        AssertManagedMatchesSerialized(JsonValue.Create(9007199254740993L));
    }

    [Fact]
    public void Nonfinite_numbers_are_rejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => Json.Canonical(JsonValue.Create(double.NaN)));
        Assert.Throws<ArgumentOutOfRangeException>(() => Json.Canonical(JsonValue.Create(double.PositiveInfinity)));
        Assert.Throws<ArgumentOutOfRangeException>(() => Json.Canonical(JsonValue.Create(double.NegativeInfinity)));
        Assert.Throws<ArgumentOutOfRangeException>(() => Json.Canonical(JsonValue.Create(float.NaN)));
        Assert.Throws<ArgumentOutOfRangeException>(() => Json.Canonical(JsonValue.Create(float.PositiveInfinity)));
        Assert.Throws<ArgumentOutOfRangeException>(() => JsonCanonicalNumber.Normalize("NaN"));
        Assert.Throws<ArgumentOutOfRangeException>(() => JsonCanonicalNumber.Normalize("-Infinity"));
    }

    [Fact]
    public void Token_hash_accepts_reordered_keys_and_rejects_meaningful_changes()
    {
        var first = Json.ParseObject("""{"op":"set_values","range":"A1","values":[[1]],"target":{"sheet":"매출"}}""")!;
        var second = Json.ParseObject("""{"target":{"sheet":"매출"},"values":[[1.0]],"range":"A1","op":"set_values"}""")!;
        var changed = Json.ParseObject("""{"op":"set_values","range":"A1","values":[[2]],"target":{"sheet":"매출"}}""")!;
        Assert.Equal(ConfirmTokenService.HashOps(new[] { first }), ConfirmTokenService.HashOps(new[] { second }));
        Assert.NotEqual(ConfirmTokenService.HashOps(new[] { first }), ConfirmTokenService.HashOps(new[] { changed }));
    }

    private static void AssertParsedHugeExponentIdentity(
        string value, string equivalent, string scaledEquivalent, string adjacent)
    {
        Assert.Equal(CanonicalNumber(value), CanonicalNumber(equivalent));
        Assert.Equal(CanonicalNumber(value), CanonicalNumber(scaledEquivalent));
        Assert.Equal(Hash(NumberOp(value)), Hash(NumberOp(equivalent)));
        Assert.Equal(Hash(NumberOp(value)), Hash(NumberOp(scaledEquivalent)));
        Assert.NotEqual(CanonicalNumber(value), CanonicalNumber(adjacent));
        Assert.NotEqual(Hash(NumberOp(value)), Hash(NumberOp(adjacent)));
        Assert.DoesNotContain("00000", CanonicalNumber(value));
    }

    private static void AssertManagedMatchesSerialized(JsonValue created)
    {
        var serialized = created.ToJsonString(Json.Compact);
        var parsed = JsonNode.Parse(serialized);
        Assert.Equal(Json.Canonical(created), Json.Canonical(parsed));
        Assert.Equal(Json.Canonical(created), Json.Canonical(JsonNode.Parse(Json.Canonical(created))));
    }

    private static string Normalize(string raw) => JsonCanonicalNumber.Normalize(raw);

    private static string CanonicalNumber(string raw) =>
        Json.Canonical(JsonNode.Parse(raw));

    private static JsonObject NumberOp(string number) =>
        Json.ParseObject($"{{\"op\":\"probe\",\"n\":{number}}}")!;

    private static string Hash(JsonObject op) => ConfirmTokenService.HashOps(new[] { op });
}
