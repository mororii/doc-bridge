using System.Reflection;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using Xunit;

namespace DocBridge.Core.Tests;

public sealed class ExcelPictureReplacementFakeTests
{
    private static JsonObject Replace(PictureSheetFake sheet, PictureShapeFake old, JsonObject op)
    {
        var leaseType = typeof(ExcelAdapter).GetNestedType("DataPictureLease", BindingFlags.NonPublic)!;
        var lease = Activator.CreateInstance(leaseType, sheet, "S", old, "P", new JsonObject())!;
        var method = typeof(ExcelAdapter).GetMethod("ApplyPictureReplacement", BindingFlags.Static | BindingFlags.NonPublic)!;
        return (JsonObject)method.Invoke(null, new[] { lease, op })!;
    }

    [Fact]
    public void Replacement_preserves_crop_rotation_and_honors_requested_geometry()
    {
        var file = Path.GetTempFileName();
        try
        {
            var old = new PictureShapeFake { Name = "P", Width = 200, Height = 100, Rotation = 45, Placement = 2, LockAspectRatio = -1 };
            old.PictureFormat.CropLeft = 10;
            var sheet = new PictureSheetFake();
            var op = new JsonObject { ["path"] = file, ["position"] = new JsonObject { ["left"] = 30, ["width"] = 300 }, ["crop"] = new JsonObject { ["right"] = 15 }, ["lockAspectRatio"] = false };
            var state = Replace(sheet, old, op);
            Assert.True(old.Deleted);
            Assert.False(sheet.Shapes.Added.Deleted);
            Assert.Equal("P", sheet.Shapes.Added.Name);
            Assert.Equal(300, sheet.Shapes.Added.Width);
            Assert.Equal(100, sheet.Shapes.Added.Height);
            Assert.Equal(45, sheet.Shapes.Added.Rotation);
            Assert.Equal(2, sheet.Shapes.Added.Placement);
            Assert.Equal(0, sheet.Shapes.Added.LockAspectRatio);
            Assert.Equal(10, sheet.Shapes.Added.PictureFormat.CropLeft);
            Assert.Equal(15, sheet.Shapes.Added.PictureFormat.CropRight);
            Assert.Equal(-1, sheet.Shapes.RequestedWidth);
            Assert.Equal(30, state["left"]!.GetValue<double>());
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Invalid_crop_is_checked_against_original_replacement_dimensions()
    {
        var file = Path.GetTempFileName();
        try
        {
            var old = new PictureShapeFake { Name = "P", Width = 1000 };
            var sheet = new PictureSheetFake();
            var op = new JsonObject { ["path"] = file, ["crop"] = new JsonObject { ["left"] = 90, ["right"] = 20 } };
            Assert.Throws<TargetInvocationException>(() => Replace(sheet, old, op));
            Assert.False(old.Deleted);
            Assert.True(sheet.Shapes.Added.Deleted);
            Assert.Equal("P", old.Name);
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Failed_name_handoff_keeps_original_and_ambiguous_delete_keeps_replacement()
    {
        var file = Path.GetTempFileName();
        try
        {
            var original = new PictureShapeFake { Name = "P" };
            var sheet = new PictureSheetFake();
            sheet.Shapes.Added.RejectFinalName = true;
            var op = new JsonObject { ["path"] = file };
            Assert.Throws<TargetInvocationException>(() => Replace(sheet, original, op));
            Assert.False(original.Deleted);
            Assert.Equal("P", original.Name);
            Assert.True(sheet.Shapes.Added.Deleted);

            original = new PictureShapeFake { Name = "P", ThrowAfterDelete = true };
            sheet = new PictureSheetFake();
            Assert.Throws<TargetInvocationException>(() => Replace(sheet, original, op));
            Assert.True(original.Deleted);
            Assert.False(sheet.Shapes.Added.Deleted);
            Assert.Equal("P", sheet.Shapes.Added.Name);
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Native_property_readers_support_new_members_and_preserve_unknown_booleans()
    {
        var font = new PictureFontFake();
        var flags = BindingFlags.Static | BindingFlags.NonPublic;
        var readDouble = typeof(ExcelAdapter).GetMethod("TryComDouble", flags)!;
        var readInt = typeof(ExcelAdapter).GetMethod("TryComInt", flags)!;
        var readBool = typeof(ExcelAdapter).GetMethod("TryComBool", flags)!;
        Assert.Equal(11.5, readDouble.Invoke(null, new object[] { font, "Size" }));
        Assert.Equal(-4119, readInt.Invoke(null, new object[] { font, "Underline" }));
        Assert.Null(readBool.Invoke(null, new object[] { font, "Bold" }));
        Assert.Equal(false, readBool.Invoke(null, new object[] { font, "Italic" }));
    }
}

public sealed class PictureFontFake
{
    public double Size => 11.5;
    public int Underline => -4119;
    public int Bold => -2;
    public bool Italic => false;
}

public sealed class PictureSheetFake { public PictureShapesFake Shapes { get; } = new(); }
public sealed class PictureShapesFake
{
    public PictureShapeFake Added { get; } = new();
    public double RequestedWidth { get; private set; }
    public PictureShapeFake AddPicture(string path, int linked, int embedded, float left, float top, float width, float height)
    { RequestedWidth = width; return Added; }
}
public sealed class PictureShapeFake
{
    private string _name = "created";
    public string Name { get => _name; set { if (RejectFinalName && value == "P") throw new InvalidOperationException("name failure"); _name = value; } }
    public bool RejectFinalName { get; set; }
    public bool ThrowAfterDelete { get; set; }
    public bool Deleted { get; private set; }
    public int Type => 13;
    public double Left { get; set; } = 10;
    public double Top { get; set; } = 20;
    public double Width { get; set; } = 100;
    public double Height { get; set; } = 80;
    public double Rotation { get; set; }
    public int LockAspectRatio { get; set; }
    public int Placement { get; set; } = 1;
    public PictureFormatFake PictureFormat { get; } = new();
    public void Delete() { Deleted = true; if (ThrowAfterDelete) throw new InvalidOperationException("ambiguous deletion"); }
}
public sealed class PictureFormatFake
{
    public double CropLeft { get; set; }
    public double CropTop { get; set; }
    public double CropRight { get; set; }
    public double CropBottom { get; set; }
}
