using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Real JSON Schema input branches for the Excel data/reporting ops.
/// <see cref="ExcelDataOperationsContract.DescribeSchema"/> is metadata only
/// and is not tools/list inputSchema.
/// </summary>
public static class ExcelDataOperationSchema
{
    public const string SchemaId = "docbridge.excel.data.ops.input/v1";

    /// <summary>
    /// Exact ToolRegistry hook. After ExcelApplyOpsSchema builds the flattened
    /// <c>properties.ops.items</c> node:
    /// <code>
    /// ExcelDataOperationSchema.AttachToApplyItems(
    ///     (JsonObject)((JsonObject)schema["properties"]!)["ops"]!["items"]!);
    /// </code>
    /// Data op names leave the shared enum. Their fields live only on per-op
    /// anyOf branches so <c>position</c>/<c>rows</c>/<c>columns</c>/<c>values</c>
    /// do not collide with move_sheet, set_row_heights, set_column_widths, or
    /// set_values.
    /// </summary>
    public static void AttachToApplyItems(JsonObject items)
    {
        ArgumentNullException.ThrowIfNull(items);
        var leftover = (JsonObject)items.DeepClone();
        leftover.Remove("anyOf");
        if (leftover["properties"] is JsonObject leftoverProps &&
            leftoverProps["op"] is JsonObject leftoverOp &&
            leftoverOp["enum"] is JsonArray leftoverEnum)
        {
            var kept = new JsonArray();
            foreach (var node in leftoverEnum)
            {
                var name = node is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;
                if (!string.IsNullOrWhiteSpace(name) && !ExcelDataOperationsContract.IsDataOperation(name))
                    kept.Add(name);
            }
            leftoverOp["enum"] = kept;
        }

        var anyOf = new JsonArray { leftover };
        foreach (var branch in OpAnyOf().OfType<JsonObject>())
            anyOf.Add(branch.DeepClone());

        items.Remove("properties");
        items.Remove("required");
        items.Remove("additionalProperties");
        items["anyOf"] = anyOf;
    }

    public static JsonArray OpAnyOf()
    {
        var branches = new JsonArray();
        foreach (var name in ExcelDataOperationsContract.WriteOpNames)
            branches.Add(OpSchema(name));
        return branches;
    }

    public static JsonObject ForToolsListItems() => new()
    {
        ["title"] = "Excel data/reporting operation",
        ["$id"] = SchemaId,
        ["anyOf"] = OpAnyOf(),
    };

    public static JsonObject OpSchema(string opName)
    {
        if (!ExcelDataOperationsContract.IsDataOperation(opName))
            throw new ArgumentOutOfRangeException(nameof(opName), opName, "not a data operation");
        return opName switch
        {
            ExcelDataOperationsContract.CreateTable => Branch(opName,
                ["range"],
                new JsonObject
                {
                    ["range"] = A1("table body including header row"),
                    ["name"] = ObjectName("ListObject name; Excel defined-name rules"),
                    ["hasHeaders"] = Bool("first row is headers"),
                    ["styleName"] = NonEmptyString("Excel table style, e.g. TableStyleMedium2"),
                    ["showTotals"] = Bool(),
                    ["totals"] = TotalsColumns(),
                },
                "Create a ListObject. position/rows/columns/values are not this op."),
            ExcelDataOperationsContract.ResizeTable => Branch(opName,
                ["name", "range"],
                new JsonObject { ["name"] = DefinedName(), ["range"] = A1("new table range") },
                "Resize an existing ListObject."),
            ExcelDataOperationsContract.StyleTable => Branch(opName,
                ["name"],
                new JsonObject
                {
                    ["name"] = DefinedName(),
                    ["styleName"] = NonEmptyString("Excel table style"),
                    ["showHeaders"] = Bool(),
                    ["showTotals"] = Bool(),
                    ["showAutoFilter"] = Bool(),
                    ["showRowStripes"] = Bool(),
                    ["showColumnStripes"] = Bool(),
                },
                "Style flags or styleName; at least one is required by the validator."),
            ExcelDataOperationsContract.SetTableTotals => Branch(opName,
                ["name", "showTotals"],
                new JsonObject
                {
                    ["name"] = DefinedName(),
                    ["showTotals"] = Bool(),
                    ["columns"] = TotalsColumns(),
                },
                "columns is an array of {column,function} — not column-width objects."),
            ExcelDataOperationsContract.AddTableColumn => Branch(opName,
                ["name", "columnName"],
                new JsonObject
                {
                    ["name"] = DefinedName(),
                    ["columnName"] = NonEmptyString("new ListColumn header"),
                    ["formula"] = Formula(),
                    ["insertAfter"] = Integer(1, 16_384, "1-based ListColumns index"),
                },
                "Add a ListColumn. insertAfter is a finite integer ≥ 1."),
            ExcelDataOperationsContract.DeleteTable => Named(opName, DefinedName(), "Delete a ListObject (cells remain)."),
            ExcelDataOperationsContract.SortRange => Branch(opName,
                ["range", "keys"],
                new JsonObject { ["range"] = A1(), ["keys"] = SortKeys(), ["hasHeaders"] = Bool() },
                "Sort a rectangular range. keys is 1..3 {column,order} objects."),
            ExcelDataOperationsContract.SortTable => Branch(opName,
                ["name", "keys"],
                new JsonObject { ["name"] = DefinedName(), ["keys"] = SortKeys() },
                "Sort a ListObject."),
            ExcelDataOperationsContract.SetAutoFilter => Branch(opName,
                ["criteria"],
                new JsonObject
                {
                    ["range"] = A1("filter range; XOR with name"),
                    ["name"] = DefinedName(),
                    ["criteria"] = FilterCriteria(),
                },
                "Exactly one of range or name. criteria is not a values matrix."),
            ExcelDataOperationsContract.ClearAutoFilter => Branch(opName,
                [],
                new JsonObject { ["range"] = A1(), ["name"] = DefinedName() },
                "Exactly one of range or name."),
            ExcelDataOperationsContract.RemoveDuplicates => Branch(opName,
                ["range"],
                new JsonObject
                {
                    ["range"] = A1(),
                    ["hasHeaders"] = Bool(),
                    ["columns"] = DuplicateColumns(),
                },
                "columns is a 1-based integer list — not set_column_widths objects."),
            ExcelDataOperationsContract.TextToColumns => Branch(opName,
                ["range"],
                new JsonObject
                {
                    ["range"] = A1("source column"),
                    ["destination"] = A1("optional destination origin"),
                    ["dataType"] = Enumerated(["delimited"], "fixed is rejected until FieldInfo exists"),
                    ["comma"] = Bool(),
                    ["tab"] = Bool(),
                    ["semicolon"] = Bool(),
                    ["space"] = Bool(),
                    ["other"] = Bool(),
                    ["otherChar"] = BoundedString(1, "required when other=true"),
                    ["consecutiveDelimiter"] = Bool(),
                    ["textQualifier"] = NonEmptyString(),
                },
                "Delimited split only. other=true requires otherChar."),
            ExcelDataOperationsContract.DefineName => NameWrite(opName, includeReplace: true),
            ExcelDataOperationsContract.UpdateName => NameWrite(opName, includeReplace: false),
            ExcelDataOperationsContract.DeleteName => Branch(opName,
                ["name", "scope"],
                new JsonObject { ["name"] = DefinedName(), ["scope"] = NameScope() },
                "Delete a defined name. workbook scope does not need target.sheet."),
            ExcelDataOperationsContract.SetDataValidation => Branch(opName,
                ["range", "type"],
                ValidationProperties(),
                "Cell validation. type is a validation token, not a chart type."),
            ExcelDataOperationsContract.ClearDataValidation => Ranged(opName, "Clear validation on a rectangle."),
            ExcelDataOperationsContract.AddConditionalFormat => Branch(opName,
                ["range", "rule"],
                new JsonObject
                {
                    ["range"] = A1(),
                    ["rule"] = ConditionalRule(),
                    ["style"] = ConditionalStyle(),
                },
                "Typed FormatCondition. expression uses rule.formula1."),
            ExcelDataOperationsContract.ClearConditionalFormats => Ranged(opName, "Delete FormatConditions on a rectangle."),
            ExcelDataOperationsContract.SetRichText => Branch(opName, ["cell", "baseFont", "runs"], new JsonObject
            {
                ["cell"] = A1("single cell"), ["baseFont"] = new JsonObject { ["type"] = "object" },
                ["runs"] = new JsonObject { ["type"] = "array", ["minItems"] = 1, ["maxItems"] = 256 },
            }, "Replace all cell text and character runs; text is literal, including leading '='."),
            ExcelDataOperationsContract.UpdateConditionalFormat => Branch(opName, ["range", "index", "expectedFingerprint"], new JsonObject
            {
                ["range"] = A1(), ["index"] = Integer(1, 64, "1-based inspected rule index"), ["expectedFingerprint"] = NonEmptyString(),
                ["rule"] = ConditionalRule(), ["style"] = ConditionalStyle(),
            }, "Update one expression/cellValue rule after matching its inspected fingerprint."),
            ExcelDataOperationsContract.DeleteConditionalFormat => Branch(opName, ["range", "index", "expectedFingerprint"], new JsonObject
            { ["range"] = A1(), ["index"] = Integer(1, 64), ["expectedFingerprint"] = NonEmptyString() }, "Delete one inspected rule only."),
            ExcelDataOperationsContract.AppendTableRows => TableRows(opName, false),
            ExcelDataOperationsContract.InsertTableRows => TableRows(opName, true),
            ExcelDataOperationsContract.DeleteTableRows => Branch(opName, ["name", "index", "count"], new JsonObject
            { ["name"] = DefinedName(), ["index"] = Integer(1, 1_048_576), ["count"] = Integer(1, 1000) }, "Delete ListRows only; headers and totals are not data rows."),
            ExcelDataOperationsContract.CreateChart => Branch(opName,
                ["sourceRange", "chartType"],
                ChartProperties(requirePositionSize: true),
                "Native ChartObject. position is {left,top,width,height} points, not a move_sheet string."),
            ExcelDataOperationsContract.UpdateChart => Branch(opName,
                ["name"],
                ChartProperties(requirePositionSize: false, allowPartialPosition: true),
                "Update chart chrome / source / type. At least one updatable field is required."),
            ExcelDataOperationsContract.DeleteChart => Named(opName, ObjectName(), "Delete a ChartObject."),
            ExcelDataOperationsContract.InsertSheetPicture => Branch(opName,
                ["path"],
                new JsonObject
                {
                    ["path"] = NonEmptyString("absolute image path"),
                    ["name"] = ObjectName(),
                    ["position"] = Position(requireSize: false),
                    ["lockAspectRatio"] = Bool(),
                },
                "Sheet picture. position is a point box."),
            ExcelDataOperationsContract.UpdatePicture => Branch(opName,
                ["name"],
                new JsonObject
                {
                    ["name"] = ObjectName(),
                    ["position"] = Position(requireSize: false, allowPartial: true),
                    ["lockAspectRatio"] = Bool(),
                    ["path"] = NonEmptyString("absolute replacement image path"),
                    ["crop"] = PictureCrop(),
                },
                "Move/resize/crop or replace a picture."),
            ExcelDataOperationsContract.DeletePicture => Branch(opName,
                ["name"],
                new JsonObject
                {
                    ["name"] = ObjectName(),
                    ["path"] = NonEmptyString("absolute image path used only to recreate pixels on restore"),
                },
                "Delete a picture Shape."),
            ExcelDataOperationsContract.SetCellNote => Branch(opName,
                ["text"],
                new JsonObject
                {
                    ["text"] = BoundedString(32_767),
                    ["range"] = A1("single cell"),
                    ["cell"] = A1("single cell"),
                    ["visible"] = Bool(),
                },
                "Comment on one cell."),
            ExcelDataOperationsContract.ClearCellNote => CellTarget(opName, "Clear a cell comment."),
            ExcelDataOperationsContract.SetHyperlink => Branch(opName,
                [],
                new JsonObject
                {
                    ["range"] = A1("single cell"),
                    ["cell"] = A1("single cell"),
                    ["address"] = BoundedString(2048),
                    ["subAddress"] = BoundedString(255),
                    ["textToDisplay"] = BoundedString(255),
                    ["screenTip"] = BoundedString(255),
                },
                "Requires address and/or subAddress."),
            ExcelDataOperationsContract.ClearHyperlink => CellTarget(opName, "Clear a hyperlink."),
            ExcelDataOperationsContract.CreatePivot => Branch(opName,
                ["sourceRange", "destination", "name"],
                PivotProperties(includeSourceAndDestination: true),
                "rows/columns/filters are field-name strings; values is [{field,function}] — not a cell matrix."),
            ExcelDataOperationsContract.UpdatePivot => Branch(opName,
                ["name"],
                PivotProperties(includeSourceAndDestination: false),
                "Requires rows, columns, values, filters, or sourceRange. Omitted dimensions are preserved."),
            ExcelDataOperationsContract.RefreshPivot => Named(opName, ObjectName(), "Refresh a PivotTable by name."),
            ExcelDataOperationsContract.DeletePivot => Named(opName, ObjectName(), "Delete a PivotTable."),
            ExcelDataOperationsContract.InsertShape => Branch(opName,
                ["shapeType"],
                ShapeProperties(includeShapeType: true, requirePositionSize: true),
                "AutoShape. position is a point box. shapeType is required."),
            ExcelDataOperationsContract.UpdateShape => Branch(opName,
                ["name"],
                ShapeProperties(includeShapeType: false, requirePositionSize: false, allowPartialPosition: true),
                "Requires position, text, or fillColor."),
            ExcelDataOperationsContract.DeleteShape => Named(opName, ObjectName(), "Delete a Shape."),
            ExcelDataOperationsContract.InsertTextbox => Branch(opName,
                [],
                ShapeProperties(includeShapeType: false, requirePositionSize: true),
                "Requires position and/or text. Not a shapeType AutoShape."),
            ExcelDataOperationsContract.UpdateTextbox => Branch(opName,
                ["name"],
                ShapeProperties(includeShapeType: false, requirePositionSize: false, allowPartialPosition: true),
                "Requires position, text, or fillColor."),
            ExcelDataOperationsContract.DeleteTextbox => Named(opName, ObjectName(), "Delete a textbox Shape."),
            ExcelDataOperationsContract.CreateConnector => Branch(opName, ["connectorType", "begin", "end"], ConnectorProperties(), "Native straight, elbow, or curve connector."),
            ExcelDataOperationsContract.UpdateConnector => Branch(opName, ["name"], ConnectorProperties(), "Attach, detach, move endpoints, or format a connector."),
            ExcelDataOperationsContract.UpdateExternalLinks => Branch(opName, [],
                new JsonObject
                {
                    ["source"] = NonEmptyString("exact link source; omit with sourceContains to update all"),
                    ["sourceContains"] = NonEmptyString("substring filter over link sources"),
                },
                "Refresh cached values from linked workbooks. Bindings are unchanged; unreachable links are refused, not retried."),
            ExcelDataOperationsContract.ChangeLinkSource => Branch(opName, ["source", "newSource"],
                new JsonObject
                {
                    ["source"] = NonEmptyString("existing link source (full path or unique leaf)"),
                    ["newSource"] = NonEmptyString("absolute existing workbook path"),
                },
                "Point an Excel link at another workbook file. Dependent formulas are captured for restore."),
            ExcelDataOperationsContract.BreakExternalLink => Branch(opName, ["source"],
                new JsonObject { ["source"] = NonEmptyString("existing link source (full path or unique leaf)") },
                "Convert external references to cached values (high-risk). Dependent formulas are captured so restore can relink."),
            ExcelDataOperationsContract.SetCalculationMode => Branch(opName, ["mode"],
                new JsonObject
                {
                    ["mode"] = Enumerated(["manual", "automatic", "semiautomatic"]),
                    ["recalculate"] = Enumerated(["none", "full", "fullRebuild"]),
                },
                "Application calculation mode. recalculate=fullRebuild matches quantity-engine refresh semantics."),
            ExcelDataOperationsContract.FreezeValues => Branch(opName, ["range"],
                new JsonObject { ["range"] = A1("at most 20000 cells") },
                "Replace formulas with cached values in place; formats are preserved. Restore rewrites the formulas."),
            ExcelDataOperationsContract.PasteSpecial => Branch(opName, ["sourceRange", "destination"],
                new JsonObject
                {
                    ["sourceRange"] = A1("may be sheet-qualified to another sheet of the same workbook"),
                    ["destination"] = A1("anchor cell or fitting range on target.sheet"),
                    ["paste"] = Enumerated(["values", "formulas", "formats", "all"], "all/formats copy natively (borders included); values/formulas keep destination formats"),
                    ["transpose"] = Bool(),
                    ["skipBlanks"] = Bool("skip blank source cells"),
                    ["operation"] = Enumerated(["none", "add", "subtract", "multiply", "divide"]),
                },
                "Clipboard-free transfer. Destination must fit the (transposed) source shape."),
            ExcelDataOperationsContract.GoalSeek => Branch(opName, ["cell", "goalCell", "goal"],
                new JsonObject
                {
                    ["cell"] = A1("single changing cell"),
                    ["goalCell"] = A1("single formula cell on the same sheet"),
                    ["goal"] = Number(null, null, "finite target number"),
                    ["tolerance"] = Number(0, 1, "relative tolerance, default scales with Excel MaxChange"),
                },
                "Range.GoalSeek. goalCell must contain a formula in preview."),
            ExcelDataOperationsContract.ProtectWorkbook => Branch(opName, [],
                new JsonObject
                {
                    ["structure"] = Bool("protect workbook structure (default true)"),
                    ["windows"] = Bool("protect workbook windows (default false)"),
                },
                "Passwordless workbook protection only; password workbooks are refused."),
            ExcelDataOperationsContract.UnprotectWorkbook => Branch(opName, [],
                new JsonObject(),
                "Remove workbook structure/windows protection. Password workbooks are refused."),
            ExcelDataOperationsContract.SetSplitPanes => Branch(opName, [],
                new JsonObject
                {
                    ["cell"] = A1("anchor cell; rows above / columns left stay fixed"),
                    ["splitRows"] = Integer(0, 1_048_575),
                    ["splitColumns"] = Integer(0, 16_383),
                    ["freeze"] = Bool("FreezePanes instead of Split"),
                    ["remove"] = Bool("clear split and freeze"),
                },
                "Window split or freeze anchored at a cell or explicit offsets."),
            ExcelDataOperationsContract.CreateSparkline => Branch(opName, ["location", "sourceData", "type"],
                SparklineProperties(),
                "Sparkline group on one cell, row, or column."),
            ExcelDataOperationsContract.UpdateSparkline => Branch(opName, ["location"],
                SparklineProperties(),
                "Requires sourceData, type, or a display option."),
            ExcelDataOperationsContract.DeleteSparkline => Branch(opName, ["location"],
                new JsonObject { ["location"] = A1("sparkline group location") },
                "Delete the sparkline group at the location."),
            ExcelDataOperationsContract.CreateSlicer => Branch(opName, ["source", "field", "name"],
                new JsonObject
                {
                    ["source"] = ObjectName("ListObject or PivotTable name"),
                    ["field"] = NonEmptyString("table column header or pivot field caption"),
                    ["name"] = ObjectName("required: stable identity for delete/verify/restore"),
                    ["caption"] = BoundedString(255),
                    ["position"] = Position(requireSize: false, allowPartial: true),
                },
                "Native slicer on the target sheet. Requested names must be unique workbook-wide."),
            ExcelDataOperationsContract.DeleteSlicer => Branch(opName, ["name"],
                new JsonObject
                {
                    ["name"] = ObjectName(),
                    ["source"] = ObjectName("optional with field: exact source for restore"),
                    ["field"] = NonEmptyString("optional with source: exact field for restore"),
                },
                "Delete a slicer (cache is removed when its last slicer goes). Without source+field, restore falls back to the slicer caption as field."),
            ExcelDataOperationsContract.ApplyCellStyle => Branch(opName, ["range", "styleName"],
                new JsonObject
                {
                    ["range"] = A1("at most 5000 cells"),
                    ["styleName"] = NonEmptyString("existing workbook style; see cellStyles inspect scope"),
                },
                "Apply a named cell style. Previous per-cell style names are captured for restore."),
            _ => throw new ArgumentOutOfRangeException(nameof(opName), opName, "schema branch missing"),
        };
    }

    public static bool TryValidate(JsonObject instance, ICollection<string> errors, string path = "$") =>
        TryValidateOp(instance, errors, path);

    public static bool TryValidateOp(JsonObject instance, ICollection<string> errors, string path = "$")
    {
        ArgumentNullException.ThrowIfNull(instance);
        ArgumentNullException.ThrowIfNull(errors);
        var op = Json.GetString(instance, "op");
        if (!ExcelDataOperationsContract.IsDataOperation(op))
        {
            errors.Add($"{path}.op is not a published data operation");
            return false;
        }
        return ValidateNode(instance, OpSchema(op!), errors, path);
    }

    /// <summary>
    /// Evaluate an instance against the published tools/list <c>ops.items</c> node
    /// after <see cref="AttachToApplyItems"/> (leftover + one typed anyOf branch per data op).
    /// </summary>
    public static bool TryValidatePublishedItems(JsonObject items, JsonObject instance, ICollection<string> errors)
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(instance);
        ArgumentNullException.ThrowIfNull(errors);
        if (items["anyOf"] is not JsonArray)
        {
            errors.Add("published items is missing anyOf; call AttachToApplyItems first");
            return false;
        }

        // Evaluate the published items node as-is (outer constraints AND anyOf).
        // Do not strip leftover properties — a re-added position:string bag
        // must reject a chart position object so the conjunction stays visible.
        return ValidateNode(instance, items, errors, "$");
    }

    public static int CountPublishedMatches(JsonObject items, JsonObject instance)
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(instance);
        if (items["anyOf"] is not JsonArray anyOf) return 0;
        var matches = 0;
        foreach (var node in anyOf.OfType<JsonObject>())
        {
            var nested = new List<string>();
            if (ValidateNode(instance, node, nested, "$") && nested.Count == 0)
                matches++;
        }
        return matches;
    }

    public static List<string> Validate(JsonObject instance)
    {
        var errors = new List<string>();
        TryValidateOp(instance, errors);
        return errors;
    }

    private static JsonObject Branch(string opName, string[] required, JsonObject properties, string description)
    {
        properties["op"] = new JsonObject
        {
            ["type"] = "string",
            ["const"] = opName,
        };
        properties["target"] = Target();
        properties["targetWorkbook"] = NonEmptyString("open workbook name or absolute path");
        var requiredList = new JsonArray { "op" };
        foreach (var field in required)
            requiredList.Add(field);
        return new JsonObject
        {
            ["type"] = "object",
            ["additionalProperties"] = false,
            ["description"] = description,
            ["required"] = requiredList,
            ["properties"] = properties,
        };
    }

    private static JsonObject Named(string opName, JsonObject nameSchema, string description) =>
        Branch(opName, ["name"], new JsonObject { ["name"] = nameSchema }, description);

    private static JsonObject Ranged(string opName, string description) =>
        Branch(opName, ["range"], new JsonObject { ["range"] = A1() }, description);

    private static JsonObject TableRows(string opName, bool insert) =>
        Branch(opName, insert ? ["name", "index", "rows"] : ["name", "rows"], new JsonObject
        {
            ["name"] = DefinedName(),
            ["index"] = Integer(1, 1_048_576, "1-based ListRows index"),
            ["rows"] = new JsonObject { ["type"] = "array", ["minItems"] = 1, ["maxItems"] = ExcelTableRowsContract.MaxRows },
        }, "Insert ListRows; rows must match the inspected table column count and omit formulas only by supplying their calculated values.");

    private static JsonObject CellTarget(string opName, string description) =>
        Branch(opName, [], new JsonObject { ["range"] = A1("single cell"), ["cell"] = A1("single cell") }, description);

    private static JsonObject NameWrite(string opName, bool includeReplace)
    {
        var props = new JsonObject
        {
            ["name"] = DefinedName(),
            ["refersTo"] = Formula(),
            ["scope"] = NameScope(),
            ["comment"] = BoundedString(255),
        };
        if (includeReplace) props["replace"] = Bool();
        return Branch(opName, ["name", "refersTo", "scope"], props, "Defined name. workbook scope may omit target.sheet.");
    }

    private static JsonObject ChartProperties(bool requirePositionSize, bool allowPartialPosition = false)
    {
        return new JsonObject
        {
            ["name"] = ObjectName(),
            ["title"] = BoundedString(255, "empty string clears the title"),
            ["hasLegend"] = Bool(),
            ["legendPosition"] = Enumerated(ExcelDataObjectCatalog.LegendPositions.Keys, "legend position"),
            ["plotBy"] = Enumerated(["rows", "columns"]),
            ["position"] = Position(requirePositionSize, allowPartialPosition),
            ["series"] = ChartSeries(),
            ["chartFill"] = ColorOrNone("ChartArea fill"),
            ["plotFill"] = ColorOrNone("PlotArea fill"),
            ["chartBorder"] = ColorOrNone("ChartArea line"),
            ["plotBorder"] = ColorOrNone("PlotArea line"),
            ["axes"] = ChartAxes(),
            ["plotArea"] = PlotArea(),
            ["dataLabels"] = DataLabels(),
            ["sourceRange"] = A1("rectangular or union A1; may be another sheet"),
            ["chartType"] = ChartType(),
        };
    }

    private static JsonObject PivotProperties(bool includeSourceAndDestination)
    {
        var props = new JsonObject
        {
            ["name"] = ObjectName(),
            ["rows"] = FieldNameArray("row fields — strings, not row-height objects"),
            ["columns"] = FieldNameArray("column fields — strings, not width objects"),
            ["filters"] = FieldNameArray("page fields"),
            ["values"] = PivotValues(),
            ["sourceRange"] = PivotSource(),
        };
        if (includeSourceAndDestination)
        {
            props["destination"] = A1("top-left of the pivot, may be sheet-qualified; must match target.sheet");
        }
        return props;
    }

    private static JsonObject ShapeProperties(bool includeShapeType, bool requirePositionSize,
        bool allowPartialPosition = false)
    {
        var props = new JsonObject
        {
            ["name"] = ObjectName(),
            ["position"] = Position(requirePositionSize, allowPartialPosition),
            ["text"] = BoundedString(32_767),
            ["fillColor"] = Color("OLE/hex color"),
            ["lineColor"] = Color("OLE/hex color"),
            ["lineWeight"] = PositiveNumber(ExcelShapeFormatContract.MaxLineWeight, "points; > 0"),
            ["lineVisible"] = Bool(),
            ["rotation"] = Integer(0, 360, "integer degrees; 360 normalizes to 0"),
            ["font"] = ShapeFont(),
        };
        if (includeShapeType)
            props["shapeType"] = Enumerated(ExcelDataObjectCatalog.ShapeTypes.Keys);
        return props;
    }

    private static JsonObject ShapeFont() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["minProperties"] = 1,
        ["description"] = "Whole shape/textbox text font. Omitted properties are preserved.",
        ["properties"] = new JsonObject
        {
            ["name"] = NonEmptyString("font name"),
            ["size"] = PositiveNumber(ExcelShapeFormatContract.MaxFontSize, "points; > 0"),
            ["bold"] = Bool(),
            ["italic"] = Bool(),
            ["color"] = Color("OLE/hex color"),
        },
    };

    private static JsonObject ValidationProperties() => new()
    {
        ["range"] = A1(),
        ["type"] = Enumerated(ExcelDataObjectCatalog.ValidationTypes.Keys),
        ["operator"] = Enumerated(ExcelDataObjectCatalog.ComparisonOperators.Keys),
        ["formula1"] = NonEmptyString(),
        ["formula2"] = NonEmptyString(),
        ["source"] = NonEmptyString("list source or A1"),
        ["inCellDropdown"] = Bool(),
        ["ignoreBlank"] = Bool(),
        ["showInput"] = Bool(),
        ["inputTitle"] = BoundedString(32),
        ["inputMessage"] = BoundedString(255),
        ["showError"] = Bool(),
        ["errorTitle"] = BoundedString(32),
        ["errorMessage"] = BoundedString(255),
        ["errorStyle"] = Enumerated(ExcelDataObjectCatalog.ErrorStyles.Keys),
    };

    private static JsonObject Target() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["description"] = "Write target. Sheet-scoped data writes need sheet or a sheet-qualified range.",
        ["properties"] = new JsonObject
        {
            ["sheet"] = NonEmptyString("exact worksheet name"),
            ["workbook"] = NonEmptyString(),
            ["scope"] = Enumerated(["sheet", "workbook"]),
        },
    };

    private static JsonObject Position(bool requireSize, bool allowPartial = false)
    {
        var node = new JsonObject
        {
            ["type"] = "object",
            ["additionalProperties"] = false,
            ["description"] = allowPartial
                ? "Partial point box. Updates may send width/height only; omitted coordinates are not rewritten as 0 by schema."
                : "Chart/shape/picture box in points. Not a move_sheet before/after string.",
            ["properties"] = new JsonObject
            {
                ["left"] = Number(0, 20_000),
                ["top"] = Number(0, 20_000),
                ["width"] = Number(0, 20_000),
                ["height"] = Number(0, 20_000),
            },
        };
        if (allowPartial)
        {
            node["minProperties"] = 1;
            return node;
        }
        var required = new JsonArray { "left", "top" };
        if (requireSize)
        {
            required.Add("width");
            required.Add("height");
        }
        node["required"] = required;
        return node;
    }

    private static JsonObject SortKeys() => new()
    {
        ["type"] = "array",
        ["minItems"] = 1,
        ["maxItems"] = 3,
        ["items"] = new JsonObject
        {
            ["type"] = "object",
            ["additionalProperties"] = false,
            ["required"] = new JsonArray("column"),
            ["properties"] = new JsonObject
            {
                ["column"] = new JsonObject
                {
                    ["oneOf"] = new JsonArray(
                        Integer(1, 16_384, "1-based index"),
                        NonEmptyString("header or column letter")),
                },
                ["order"] = Enumerated(["asc", "desc", "ascending", "descending"]),
            },
        },
    };

    private static JsonObject FilterCriteria() => new()
    {
        ["type"] = "array",
        ["minItems"] = 1,
        ["maxItems"] = 256,
        ["items"] = new JsonObject
        {
            ["type"] = "object",
            ["additionalProperties"] = false,
            ["required"] = new JsonArray("column", "operator"),
            ["properties"] = new JsonObject
            {
                ["column"] = new JsonObject
                {
                    ["oneOf"] = new JsonArray(Integer(1, 16_384), NonEmptyString()),
                },
                ["operator"] = Enumerated(ExcelDataObjectCatalog.FilterOperators),
                ["value"] = TrueSchema("scalar criterion"),
                ["value2"] = TrueSchema("between upper bound"),
                ["values"] = new JsonObject
                {
                    ["type"] = "array",
                    ["minItems"] = 1,
                    ["maxItems"] = 100,
                    ["items"] = NonEmptyString(),
                },
            },
        },
    };

    private static JsonObject TotalsColumns() => new()
    {
        ["type"] = "array",
        ["minItems"] = 1,
        ["maxItems"] = 256,
        ["description"] = "Table totals — not set_column_widths objects.",
        ["items"] = new JsonObject
        {
            ["type"] = "object",
            ["additionalProperties"] = false,
            ["required"] = new JsonArray("column", "function"),
            ["properties"] = new JsonObject
            {
                ["column"] = new JsonObject
                {
                    ["oneOf"] = new JsonArray(Integer(1, 16_384), NonEmptyString()),
                },
                ["function"] = Enumerated(ExcelDataObjectCatalog.TotalsFunctions.Keys),
                ["formula"] = Formula(),
            },
        },
    };

    private static JsonObject DuplicateColumns() => new()
    {
        ["type"] = "array",
        ["minItems"] = 1,
        ["description"] = "1-based integer indexes only.",
        ["items"] = Integer(1, 16_384),
    };

    private static JsonObject FieldNameArray(string description) => new()
    {
        ["type"] = "array",
        ["description"] = description,
        ["items"] = NonEmptyString(),
    };

    private static JsonObject PivotValues() => new()
    {
        ["type"] = "array",
        ["minItems"] = 1,
        ["maxItems"] = 64,
        ["description"] = "Pivot data fields. Not a set_values 2D matrix.",
        ["items"] = new JsonObject
        {
            ["type"] = "object",
            ["additionalProperties"] = false,
            ["required"] = new JsonArray("field"),
            ["properties"] = new JsonObject
            {
                ["field"] = NonEmptyString("source column caption"),
                ["function"] = Enumerated(ExcelDataObjectCatalog.PivotFunctions.Keys,
                    "omit to default to sum; JSON null is invalid"),
                ["caption"] = BoundedString(255, "PivotField.Caption"),
                ["numberFormat"] = BoundedString(255, "PivotField.NumberFormat"),
            },
        },
    };

    private static JsonObject PivotSource() => new()
    {
        ["type"] = "string",
        ["minLength"] = 1,
        ["description"] = "Rectangular A1 (optionally sheet-qualified) or a ListObject name. Not a union.",
    };

    private static JsonObject ChartSeries() => new()
    {
        ["type"] = "array",
        ["minItems"] = 1,
        ["maxItems"] = 64,
        ["items"] = new JsonObject
        {
            ["type"] = "object",
            ["additionalProperties"] = false,
            ["properties"] = new JsonObject
            {
                ["values"] = A1("Series.Values A1 or union"),
                ["categories"] = A1("Series.XValues"),
                ["name"] = new JsonObject
                {
                    ["type"] = "string",
                    ["minLength"] = 1,
                    ["description"] = "literal series name or A1 name range",
                },
                ["range"] = A1("optional whole-series source"),
                ["index"] = Number(1, 64, "1-based SeriesCollection index; create_chart must be 1..n contiguous"),
                ["chartType"] = ChartType("per-series XlChartType; omit to inherit chart-level chartType"),
                ["axisGroup"] = Enumerated(
                    ExcelDataObjectCatalog.AxisGroups.Keys,
                    "primary|secondary; omit to keep Series.AxisGroup"),
                ["lineColor"] = Color(),
                ["lineWeight"] = Number(0.25, 10, "points"),
                ["marker"] = Enumerated(ExcelDataObjectCatalog.MarkerStyles.Keys),
                ["dataLabels"] = DataLabels(),
                ["trendline"] = Trendline(),
                ["points"] = ChartPoints(),
            },
        },
    };

    private static JsonObject ConnectorProperties() => new()
    {
        ["name"] = ObjectName(), ["connectorType"] = Enumerated(["straight", "elbow", "curve"]),
        ["begin"] = ConnectorEndpoint(), ["end"] = ConnectorEndpoint(), ["lineColor"] = Color(),
        ["lineWeight"] = PositiveNumber(ExcelShapeFormatContract.MaxLineWeight, "points"), ["lineVisible"] = Bool(),
    };
    private static JsonObject ConnectorEndpoint() => new()
    {
        ["type"] = "object", ["additionalProperties"] = false,
        ["properties"] = new JsonObject { ["x"] = Number(0, 20_000), ["y"] = Number(0, 20_000), ["shape"] = ObjectName(), ["site"] = Integer(1, 1000) },
    };

    private static JsonObject SparklineProperties()
    {
        var props = new JsonObject
        {
            ["location"] = A1("sparkline group location"),
            ["sourceData"] = A1("rectangular data range"),
            ["type"] = Enumerated(["line", "column"], "winLoss is rejected until a native mapping is proven"),
            ["markers"] = Bool(),
            ["lineColor"] = Color(),
            ["showHigh"] = Bool(),
            ["showLow"] = Bool(),
            ["showNegative"] = Bool(),
            ["showFirst"] = Bool(),
            ["showLast"] = Bool(),
        };
        return props;
    }

    private static JsonObject ChartAxes() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["properties"] = new JsonObject
        {
            ["category"] = Axis("primary category"),
            ["value"] = Axis("primary value; group=secondary is a shorthand for valueSecondary"),
            ["valueSecondary"] = Axis("secondary value axis (xlValue, xlSecondary)"),
        },
    };

    private static JsonObject Axis(string? description = null) => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["description"] = description ?? "chart axis",
        ["properties"] = new JsonObject
        {
            ["group"] = Enumerated(ExcelDataObjectCatalog.AxisGroups.Keys, "xlPrimary|xlSecondary"),
            ["visible"] = Bool(),
            ["minimum"] = Number(null, null),
            ["maximum"] = Number(null, null),
            ["numberFormat"] = new JsonObject
            {
                ["type"] = "string",
                ["minLength"] = 1,
                ["maxLength"] = 255,
                ["description"] = "TickLabels.NumberFormat",
            },
            ["title"] = BoundedString(255, "empty string clears AxisTitle"),
        },
    };

    private static JsonObject PlotArea() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["properties"] = new JsonObject
        {
            ["spanChart"] = Bool("stretch PlotArea to ChartArea"),
            ["left"] = Number(0, 20_000),
            ["top"] = Number(0, 20_000),
            ["width"] = Number(0, 20_000),
            ["height"] = Number(0, 20_000),
        },
    };

    private static JsonObject DataLabels() => new()
    {
        ["type"] = "object", ["additionalProperties"] = false,
        ["properties"] = new JsonObject { ["show"] = Bool(), ["value"] = Bool(), ["category"] = Bool(), ["series"] = Bool(), ["percentage"] = Bool() },
    };

    private static JsonObject PictureCrop() => new()
    {
        ["type"] = "object", ["additionalProperties"] = false,
        ["properties"] = new JsonObject { ["left"] = Number(0, 20_000), ["top"] = Number(0, 20_000), ["right"] = Number(0, 20_000), ["bottom"] = Number(0, 20_000) },
    };
    private static JsonObject Trendline() => new()
    {
        ["type"] = "object", ["additionalProperties"] = false, ["required"] = new JsonArray("action"),
        ["properties"] = new JsonObject { ["action"] = Enumerated(["add", "update", "delete"]), ["index"] = Integer(1, 64), ["type"] = Enumerated(["linear", "exponential", "logarithmic", "polynomial", "power", "movingAverage"]), ["order"] = Integer(2, 6), ["period"] = Integer(2, 255), ["name"] = BoundedString(255), ["displayEquation"] = Bool(), ["displayRSquared"] = Bool(), ["expectedRemainingCount"] = Integer(0, 64) },
    };
    private static JsonObject ChartPoints() => new()
    {
        ["type"] = "array", ["maxItems"] = 10_000,
        ["items"] = new JsonObject { ["type"] = "object", ["additionalProperties"] = false, ["required"] = new JsonArray("index"), ["properties"] = new JsonObject { ["index"] = Integer(1, 1_000_000), ["fillColor"] = Color(), ["lineColor"] = Color(), ["dataLabels"] = DataLabels() } },
    };

    private static JsonObject ConditionalRule() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["required"] = new JsonArray("type"),
        ["properties"] = new JsonObject
        {
            ["type"] = Enumerated(ExcelDataObjectCatalog.ConditionalRuleTypes),
            ["operator"] = Enumerated(ExcelDataObjectCatalog.ComparisonOperators.Keys),
            ["formula1"] = NonEmptyString("expression and cellValue use formula1, not formula"),
            ["formula2"] = NonEmptyString(),
            ["points"] = Integer(2, 3),
            ["iconSet"] = Enumerated(ExcelDataObjectCatalog.IconSets.Keys),
        },
    };

    private static JsonObject ConditionalStyle() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["properties"] = new JsonObject
        {
            ["bold"] = Bool(),
            ["fontBold"] = Bool(),
            ["italic"] = Bool(),
            ["fontItalic"] = Bool(),
            ["fontColor"] = Color(),
            ["fillColor"] = Color(),
            ["interiorColor"] = Color(),
            ["fill"] = Color(),
        },
    };

    private static JsonObject ChartType(string? description = null) =>
        Enumerated(ExcelDataObjectCatalog.ChartTypes.Keys, description ?? "native XlChartType token");

    private static JsonObject A1(string? description = null) => new()
    {
        ["type"] = "string",
        ["minLength"] = 1,
        ["description"] = description ?? "rectangular A1, optional sheet qualifier",
    };

    private static JsonObject DefinedName() => new()
    {
        ["type"] = "string",
        ["minLength"] = 1,
        ["maxLength"] = 255,
        ["description"] = "Excel defined-name rules (not A1, not C/R).",
    };

    private static JsonObject ObjectName(string? description = null) => new()
    {
        ["type"] = "string",
        ["minLength"] = 1,
        ["maxLength"] = 255,
        ["description"] = description ?? "Excel object name without '!' or control chars",
    };

    private static JsonObject NameScope() => Enumerated(["workbook", "sheet"]);

    private static JsonObject Formula() => new()
    {
        ["type"] = "string",
        ["minLength"] = 2,
        ["maxLength"] = 255,
        ["pattern"] = "^=.*",
        ["description"] = "Excel formula starting with '='",
    };

    private static JsonObject Color(string? description = null) => new()
    {
        ["type"] = "string",
        ["minLength"] = 1,
        ["description"] = description ?? "hex (#RRGGBB), rgb(), or named color",
    };

    private static JsonObject ColorOrNone(string description) => new()
    {
        ["type"] = "string",
        ["minLength"] = 1,
        ["description"] = description + "; 'none' means Format.Fill/Line.Visible=false",
    };

    private static JsonObject Bool(string? description = null) => new()
    {
        ["type"] = "boolean",
        ["description"] = description ?? "boolean; JSON null is invalid",
    };

    private static JsonObject NonEmptyString(string? description = null) => new()
    {
        ["type"] = "string",
        ["minLength"] = 1,
        ["description"] = description ?? "non-empty string",
    };

    private static JsonObject BoundedString(int max, string? description = null) => new()
    {
        ["type"] = "string",
        ["maxLength"] = max,
        ["description"] = description ?? $"string up to {max} chars",
    };

    private static JsonObject Integer(int min, int max, string? description = null) => new()
    {
        ["type"] = "integer",
        ["minimum"] = min,
        ["maximum"] = max,
        ["description"] = description ?? $"integer {min}..{max}",
    };

    private static JsonObject Number(double? min, double? max, string? description = null)
    {
        var node = new JsonObject
        {
            ["type"] = "number",
            ["description"] = description ?? "finite number",
        };
        if (min is double minimum) node["minimum"] = minimum;
        if (max is double maximum) node["maximum"] = maximum;
        return node;
    }

    private static JsonObject PositiveNumber(double max, string description) => new()
    {
        ["type"] = "number",
        ["exclusiveMinimum"] = 0,
        ["maximum"] = max,
        ["description"] = description,
    };

    private static JsonObject Enumerated(IEnumerable<string> tokens, string? description = null)
    {
        var values = new JsonArray();
        foreach (var token in tokens)
            values.Add(token);
        return new JsonObject
        {
            ["type"] = "string",
            ["enum"] = values,
            ["description"] = description ?? "closed token set",
        };
    }

    private static JsonObject TrueSchema(string description) => new()
    {
        ["description"] = description,
    };

    internal static bool ValidateNode(JsonNode? instance, JsonObject schema, ICollection<string> errors, string path)
    {
        var before = errors.Count;
        if (schema["const"] is JsonValue constant)
        {
            if (!JsonEquals(instance, constant))
                errors.Add($"{path} must equal {constant.ToJsonString()}");
        }
        if (schema["enum"] is JsonArray enumerated)
        {
            if (instance is null || !enumerated.Any(item => JsonEquals(instance, item)))
                errors.Add($"{path} must be one of the published enum tokens");
        }
        if (schema["type"] is JsonValue typeNode && typeNode.TryGetValue<string>(out var type))
        {
            if (!InstanceIsType(instance, type))
            {
                errors.Add($"{path} must be {type} (JSON null is not a default)");
                return errors.Count == before;
            }
        }
        if (instance is JsonValue scalar)
        {
            if (schema["minLength"] is not null && scalar.TryGetValue<string>(out var text))
            {
                var minLength = schema["minLength"]!.GetValue<int>();
                if (text.Length < minLength)
                    errors.Add($"{path} must be at least {minLength} characters");
            }
            if (schema["maxLength"] is not null && scalar.TryGetValue<string>(out var bounded))
            {
                var maxLength = schema["maxLength"]!.GetValue<int>();
                if (bounded.Length > maxLength)
                    errors.Add($"{path} exceeds {maxLength} characters");
            }
            if (schema["pattern"] is JsonValue pattern && pattern.TryGetValue<string>(out var regex) &&
                scalar.TryGetValue<string>(out var patterned) &&
                !System.Text.RegularExpressions.Regex.IsMatch(patterned, regex))
                errors.Add($"{path} does not match {regex}");
            if ((schema["minimum"] is not null || schema["exclusiveMinimum"] is not null || schema["maximum"] is not null) &&
                ExcelDataOperationsContract.TryGetFiniteNumber(scalar, out var number))
            {
                if (schema["minimum"] is not null &&
                    ExcelDataOperationsContract.TryGetFiniteNumber(schema["minimum"], out var min) &&
                    number < min)
                    errors.Add($"{path} must be >= {min}");
                if (schema["exclusiveMinimum"] is not null &&
                    ExcelDataOperationsContract.TryGetFiniteNumber(schema["exclusiveMinimum"], out var exclusiveMin) &&
                    number <= exclusiveMin)
                    errors.Add($"{path} must be > {exclusiveMin}");
                if (schema["maximum"] is not null &&
                    ExcelDataOperationsContract.TryGetFiniteNumber(schema["maximum"], out var max) &&
                    number > max)
                    errors.Add($"{path} must be <= {max}");
            }
            if (string.Equals(Json.GetString(schema, "type"), "integer", StringComparison.Ordinal) &&
                ExcelDataOperationsContract.TryGetFiniteNumber(scalar, out var integral) &&
                Math.Abs(integral - Math.Truncate(integral)) > double.Epsilon)
                errors.Add($"{path} must be an integer");
        }
        if (instance is JsonArray array && schema["items"] is JsonObject itemSchema)
        {
            var minItems = Json.GetInt(schema, "minItems");
            var maxItems = Json.GetInt(schema, "maxItems");
            if (minItems is int minCount && array.Count < minCount)
                errors.Add($"{path} must contain at least {minCount} items");
            if (maxItems is int maxCount && array.Count > maxCount)
                errors.Add($"{path} must contain at most {maxCount} items");
            for (var i = 0; i < array.Count; i++)
                ValidateNode(array[i], itemSchema, errors, $"{path}[{i}]");
        }
        if (instance is JsonObject obj)
        {
            if (schema["minProperties"] is not null)
            {
                var minProperties = schema["minProperties"]!.GetValue<int>();
                if (obj.Count < minProperties)
                    errors.Add($"{path} must have at least {minProperties} properties");
            }
            if (schema["properties"] is JsonObject properties)
            {
                if (schema["required"] is JsonArray required)
                {
                    foreach (var node in required)
                    {
                        var field = node?.ToString();
                        if (string.IsNullOrWhiteSpace(field)) continue;
                        if (!obj.ContainsKey(field) || obj[field] is null)
                            errors.Add($"{path}.{field} is required and must not be null");
                    }
                }
                var allowAdditional = schema["additionalProperties"] is not JsonValue flag ||
                                      !flag.TryGetValue<bool>(out var allowed) || allowed;
                foreach (var (key, value) in obj)
                {
                    if (properties[key] is JsonObject child)
                        ValidateNode(value, child, errors, $"{path}.{key}");
                    else if (!allowAdditional)
                        errors.Add($"{path}.{key} is not allowed");
                }
            }
        }
        if (schema["oneOf"] is JsonArray oneOf)
        {
            var matches = 0;
            foreach (var node in oneOf.OfType<JsonObject>())
            {
                var nested = new List<string>();
                if (ValidateNode(instance, node, nested, path) && nested.Count == 0)
                    matches++;
            }
            if (matches != 1)
                errors.Add($"{path} must match exactly one alternative");
        }
        if (schema["anyOf"] is JsonArray anyOf)
        {
            var matches = 0;
            List<string>? branchErrors = null;
            foreach (var node in anyOf.OfType<JsonObject>())
            {
                var nested = new List<string>();
                if (ValidateNode(instance, node, nested, path) && nested.Count == 0)
                    matches++;
                else if (branchErrors is null && BranchTargetsOp(node, instance))
                    branchErrors = nested;
            }
            if (matches == 0)
            {
                if (branchErrors is { Count: > 0 })
                {
                    foreach (var error in branchErrors)
                        errors.Add(error);
                }
                errors.Add($"{path} matches no published anyOf branch");
            }
        }
        return errors.Count == before;
    }

    private static bool BranchTargetsOp(JsonObject branch, JsonNode? instance)
    {
        if (instance is not JsonObject obj) return false;
        var op = Json.GetString(obj, "op");
        if (string.IsNullOrWhiteSpace(op)) return false;
        if (branch["properties"] is not JsonObject properties || properties["op"] is not JsonObject opSchema)
            return false;
        if (opSchema["const"] is JsonValue constant && constant.TryGetValue<string>(out var name))
            return string.Equals(name, op, StringComparison.OrdinalIgnoreCase);
        if (opSchema["enum"] is JsonArray enumerated)
        {
            return enumerated.Any(node =>
                node is JsonValue value && value.TryGetValue<string>(out var token) &&
                string.Equals(token, op, StringComparison.OrdinalIgnoreCase));
        }
        return false;
    }

    private static bool InstanceIsType(JsonNode? instance, string type) => type switch
    {
        "object" => instance is JsonObject,
        "array" => instance is JsonArray,
        "string" => instance is JsonValue value && value.TryGetValue<string>(out _),
        "boolean" => instance is JsonValue flag && flag.TryGetValue<bool>(out _),
        "integer" => instance is JsonValue number &&
                     ExcelDataOperationsContract.TryGetFiniteNumber(number, out var integral) &&
                     Math.Abs(integral - Math.Truncate(integral)) <= double.Epsilon,
        "number" => instance is JsonValue numeric &&
                    ExcelDataOperationsContract.TryGetFiniteNumber(numeric, out _),
        "null" => instance is null,
        _ => true,
    };

    private static bool JsonEquals(JsonNode? left, JsonNode? right)
    {
        if (left is null || right is null) return left is null && right is null;
        return string.Equals(left.ToJsonString(), right.ToJsonString(), StringComparison.Ordinal);
    }
}
