using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static JsonObject ApplyPictureReplacement(DataPictureLease picture, JsonObject op)
    {
        var path = Path.GetFullPath(Json.GetString(op, "path")!);
        if (!File.Exists(path)) throw new InvalidOperationException("[EXCEL_PICTURE_NOT_FOUND] replacement file does not exist");
        object? shapes = null; object? replacement = null; object? oldFormat = null; object? newFormat = null;
        var originalRenamed = false;
        var deletionAttempted = false;
        try
        {
            var old = picture.Shape;
            var position = new JsonObject();
            foreach (var field in new[] { "left", "top", "width", "height" })
            {
                var property = char.ToUpperInvariant(field[0]) + field[1..];
                var actual = TryComDouble(old, property) ?? throw new InvalidOperationException("[EXCEL_PICTURE_UNREADABLE] " + field);
                position[field] = Json.GetObj(op, "position")?[field]?.DeepClone() ?? JsonValue.Create(actual);
            }
            var rotation = TryComDouble(old, "Rotation") ?? throw new InvalidOperationException("[EXCEL_PICTURE_UNREADABLE] rotation");
            var aspect = TryComInt(old, "LockAspectRatio") ?? throw new InvalidOperationException("[EXCEL_PICTURE_UNREADABLE] aspect setting");
            if (aspect is not (0 or -1)) throw new InvalidOperationException("[EXCEL_PICTURE_UNREADABLE] mixed aspect setting");
            var expectedAspect = op.ContainsKey("lockAspectRatio") ? Json.GetBool(op, "lockAspectRatio") : aspect == -1;
            var placement = TryComInt(old, "Placement");
            oldFormat = (object)((dynamic)old).PictureFormat;
            var crop = new JsonObject();
            foreach (var side in new[] { "left", "top", "right", "bottom" })
            {
                var property = "Crop" + char.ToUpperInvariant(side[0]) + side[1..];
                var actual = TryComDouble(oldFormat, property) ?? throw new InvalidOperationException("[EXCEL_PICTURE_UNREADABLE] " + property);
                crop[side] = Json.GetObj(op, "crop")?[side]?.DeepClone() ?? JsonValue.Create(actual);
            }
            shapes = (object)((dynamic)picture.Sheet).Shapes;
            replacement = (object)((dynamic)shapes).AddPicture(path, 0, -1, 0f, 0f, -1f, -1f);
            var tempName = "__docbridge_replace_" + Guid.NewGuid().ToString("N");
            ((dynamic)replacement).Name = tempName;
            var nativeWidth = TryComDouble(replacement, "Width") ?? throw new InvalidOperationException("[EXCEL_PICTURE_UNREADABLE] native width");
            var nativeHeight = TryComDouble(replacement, "Height") ?? throw new InvalidOperationException("[EXCEL_PICTURE_UNREADABLE] native height");
            var horizontalCrop = ExcelShapeFormatContract.ReadFiniteNumber(crop["left"]!) + ExcelShapeFormatContract.ReadFiniteNumber(crop["right"]!);
            var verticalCrop = ExcelShapeFormatContract.ReadFiniteNumber(crop["top"]!) + ExcelShapeFormatContract.ReadFiniteNumber(crop["bottom"]!);
            if (horizontalCrop >= nativeWidth || verticalCrop >= nativeHeight)
                throw new InvalidOperationException("[EXCEL_PICTURE_CROP] crop consumes the replacement image; supply compatible crop values");
            newFormat = (object)((dynamic)replacement).PictureFormat;
            ApplyCrop(newFormat, crop);
            ((dynamic)replacement).LockAspectRatio = 0;
            ApplyPosition(replacement, ReadPosition(new JsonObject { ["position"] = position.DeepClone() }, 0, 0, 0, 0));
            ((dynamic)replacement).Rotation = rotation;
            if (placement is not null) ((dynamic)replacement).Placement = placement.Value;
            ((dynamic)replacement).LockAspectRatio = expectedAspect ? -1 : 0;
            var expected = new JsonObject { ["name"] = tempName, ["position"] = position.DeepClone(), ["rotation"] = rotation, ["crop"] = crop.DeepClone(), ["lockAspectRatio"] = expectedAspect };
            var checks = new List<string>();
            VerifyPicture(ReadPictureState(replacement, picture.SheetName), expected, picture.SheetName, checks);
            if (checks.Count != 0) throw new InvalidOperationException("[EXCEL_PICTURE_READBACK] " + string.Join("; ", checks));
            ((dynamic)old).Name = "__docbridge_previous_" + Guid.NewGuid().ToString("N");
            originalRenamed = true;
            ((dynamic)replacement).Name = picture.Name;
            deletionAttempted = true;
            ((dynamic)old).Delete();
            return ReadPictureState(replacement, picture.SheetName);
        }
        catch
        {
            if (!deletionAttempted)
            {
                try { if (replacement is not null) ((dynamic)replacement).Delete(); } catch { }
                try { if (originalRenamed) ((dynamic)picture.Shape).Name = picture.Name; } catch { }
            }
            throw;
        }
        finally
        {
            RotHelper.ReleaseComReference(newFormat); RotHelper.ReleaseComReference(oldFormat);
            RotHelper.ReleaseComReference(replacement); RotHelper.ReleaseComReference(shapes);
        }
    }

    private static void ApplyCrop(object format, JsonObject crop)
    {
        if (crop.ContainsKey("left")) ((dynamic)format).CropLeft = ExcelShapeFormatContract.ReadFiniteNumber(crop["left"]!);
        if (crop.ContainsKey("top")) ((dynamic)format).CropTop = ExcelShapeFormatContract.ReadFiniteNumber(crop["top"]!);
        if (crop.ContainsKey("right")) ((dynamic)format).CropRight = ExcelShapeFormatContract.ReadFiniteNumber(crop["right"]!);
        if (crop.ContainsKey("bottom")) ((dynamic)format).CropBottom = ExcelShapeFormatContract.ReadFiniteNumber(crop["bottom"]!);
    }
}
