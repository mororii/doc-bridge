using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>표 구조와 기존 양식 필드를 직접 HWP Automation으로 편집한다.</summary>
public sealed partial class HwpAdapter
{
    private sealed record HwpTableDimensions(int Rows, int Columns);
    internal readonly record struct HwpRowHeightSpec(int Row, double HeightMm);

    /// <summary>
    /// 표 컨트롤의 논리 행/열 수를 커서를 움직이지 않고 읽는다. 한글 버전에 따라
    /// 속성 이름이 달라질 수 있어 알려진 이름을 차례로 검사하며, 읽지 못한 경우에는
    /// 실행 횟수 검증만 사용한다.
    /// </summary>
    private static HwpTableDimensions? TryReadTableDimensions(object hwpObject, int tableIndex)
    {
        dynamic hwp = hwpObject;
        dynamic? table = FindControl(hwp, "tbl", tableIndex);
        if (table is null) return null;
        return TryReadTableDimensionsOnControl(hwpObject, (object)table, tableIndex);
    }

    private static HwpTableDimensions? TryReadTableDimensionsOnControl(
        object hwpObject, object tableObject, int tableIndex)
    {
        dynamic table = tableObject;
        try
        {
            dynamic properties = table.Properties;
            static int? Read(dynamic source, params string[] names)
            {
                foreach (var name in names)
                {
                    try
                    {
                        var value = Convert.ToInt32(source.Item(name));
                        if (value > 0) return value;
                    }
                    catch { }
                    try
                    {
                        var value = Convert.ToInt32(source[name]);
                        if (value > 0) return value;
                    }
                    catch { }
                }
                return null;
            }

            var rows = Read(properties, "Rows", "RowCount");
            var columns = Read(properties, "Cols", "Columns", "ColumnCount");
            if (rows is not null && columns is not null)
                return new HwpTableDimensions(rows.Value, columns.Value);
        }
        catch { }
        return TryReadTableDimensionsByNavigation(hwpObject, tableObject, tableIndex);
    }

    private static string? CurrentTablePositionKey(object hwpObject)
    {
        dynamic hwp = hwpObject;
        try
        {
            dynamic position = hwp.CreateSet("ListParaPos");
            if (!(bool)hwp.GetPosBySet(position)) return null;
            return $"{Convert.ToInt32(position.Item("List"))}:{Convert.ToInt32(position.Item("Para"))}:{Convert.ToInt32(position.Item("Pos"))}";
        }
        catch { return null; }
    }

    /// <summary>
    /// 일부 HWP 2024 빌드는 table.Properties의 Rows/Cols를 노출하지 않는다.
    /// 첫 셀에서 아래 셀 위치를 표식으로 잡고 TableRightCell의 wrap 지점까지 세어 열 수를,
    /// TableLowerCell 성공 횟수로 행 수를 독립 검증한다. 원래 캐럿은 항상 복원한다.
    /// </summary>
    private static HwpTableDimensions? TryReadTableDimensionsByNavigation(
        object hwpObject, object tableObject, int tableIndex)
    {
        dynamic hwp = hwpObject;
        object? original = null;
        try
        {
            dynamic position = hwp.CreateSet("ListParaPos");
            if ((bool)hwp.GetPosBySet(position)) original = (object)position;
        }
        catch { }

        try
        {
            if (!SelectTableCellBlockOnControl(hwpObject, tableObject, tableIndex, 0, 0, 0, out _)) return null;
            try { hwp.HAction.Run("Cancel"); } catch { }
            var first = CurrentTablePositionKey(hwpObject);
            if (first is null) return null;

            string? secondRowFirst = null;
            if ((bool)hwp.HAction.Run("TableCellBlock") && (bool)hwp.HAction.Run("TableLowerCell"))
            {
                try { hwp.HAction.Run("Cancel"); } catch { }
                secondRowFirst = CurrentTablePositionKey(hwpObject);
            }
            try { hwp.HAction.Run("Cancel"); } catch { }

            if (!SelectTableCellBlockOnControl(hwpObject, tableObject, tableIndex, 0, 0, 0, out _)) return null;
            try { hwp.HAction.Run("Cancel"); } catch { }
            var columns = 1;
            for (; columns < 2000; columns++)
            {
                if (!(bool)hwp.HAction.Run("TableCellBlock") || !(bool)hwp.HAction.Run("TableRightCell")) break;
                try { hwp.HAction.Run("Cancel"); } catch { }
                var current = CurrentTablePositionKey(hwpObject);
                if (secondRowFirst is not null && string.Equals(current, secondRowFirst, StringComparison.Ordinal)) break;
                if (secondRowFirst is null && string.Equals(current, first, StringComparison.Ordinal)) break;
            }
            if (columns >= 2000) return null;

            if (!SelectTableCellBlockOnControl(hwpObject, tableObject, tableIndex, 0, 0, 0, out _)) return null;
            try { hwp.HAction.Run("Cancel"); } catch { }
            var rows = 1;
            while (rows < 2000)
            {
                if (!(bool)hwp.HAction.Run("TableCellBlock") || !(bool)hwp.HAction.Run("TableLowerCell")) break;
                rows++;
                try { hwp.HAction.Run("Cancel"); } catch { }
            }
            if (rows >= 2000) return null;
            return new HwpTableDimensions(rows, columns);
        }
        catch { return null; }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
            if (original is not null)
                try { hwp.SetPosBySet((dynamic)original); } catch { }
        }
    }

    private static bool SelectTableCellBlock(object hwpObject, int tableIndex, int row, int col, out string error) =>
        SelectTableCellBlock(hwpObject, tableIndex, row, col, null, out error);

    /// <summary>
    /// 표 셀을 선택한다. 병합 셀은 논리 row/col이 모호하므로 cellIndex(표의 실제 이동 순서)를 우선한다.
    /// 기존 row/col 호출도 유지하되 모든 이동 결과를 검사하여 잘못된 셀에 조용히 쓰는 일을 막는다.
    /// </summary>
    private static bool SelectTableCellBlock(
        object hwpObject, int tableIndex, int row, int col, int? cellIndex, out string error)
    {
        error = "";
        if (tableIndex < 0 || row < 0 || col < 0 || cellIndex < 0)
        {
            error = "tableIndex, row, col, cellIndex는 0 이상이어야 합니다";
            return false;
        }
        dynamic hwp = hwpObject;
        dynamic? table = FindControl(hwp, "tbl", tableIndex);
        if (table is null) { error = $"표 {tableIndex}을 찾을 수 없습니다"; return false; }
        return SelectTableCellBlockOnControl(hwpObject, (object)table, tableIndex, row, col, cellIndex, out error);
    }

    private static bool SelectTableCellBlockOnControl(
        object hwpObject, object tableObject, int tableIndex, int row, int col, int? cellIndex, out string error)
    {
        error = "";
        dynamic hwp = hwpObject;
        dynamic table = tableObject;
        if (!(bool)hwp.SetPosBySet(table.GetAnchorPos(0)))
        {
            error = $"표 {tableIndex}의 기준 위치로 이동하지 못했습니다";
            return false;
        }
        _ = hwp.HAction.Run("SelectCtrlFront");
        _ = hwp.HAction.Run("ShapeObjTextBoxEdit");
        if (!(bool)hwp.HAction.Run("TableCellBlock"))
        {
            error = $"표 {tableIndex}의 첫 셀을 선택하지 못했습니다";
            return false;
        }

        if (cellIndex is not null)
        {
            for (var i = 0; i < cellIndex.Value; i++)
            {
                if ((bool)hwp.HAction.Run("TableRightCell")) continue;
                error = $"표 {tableIndex}에 cellIndex={cellIndex} 셀이 없습니다";
                try { hwp.HAction.Run("Cancel"); } catch { }
                return false;
            }
        }
        else
        {
            for (var i = 0; i < row; i++)
            {
                if ((bool)hwp.HAction.Run("TableLowerCell")) continue;
                error = $"표 {tableIndex}에서 row={row}로 이동하지 못했습니다. 병합 표에는 cellIndex를 사용하세요";
                try { hwp.HAction.Run("Cancel"); } catch { }
                return false;
            }
            for (var i = 0; i < col; i++)
            {
                if ((bool)hwp.HAction.Run("TableRightCell")) continue;
                error = $"표 {tableIndex}에서 col={col}로 이동하지 못했습니다. 병합 표에는 cellIndex를 사용하세요";
                try { hwp.HAction.Run("Cancel"); } catch { }
                return false;
            }
        }
        try
        {
            if ((Convert.ToInt32(hwp.CurFieldState) & 0x0F) != 1)
            {
                if (IsCurrentTableCellFormula(hwpObject)) return true;
                error = "대상 셀 위치를 확인할 수 없습니다";
                return false;
            }
        }
        catch { }
        return true;
    }

    private static int? TryCountTableCellsOnControl(
        object hwpObject, object tableObject, int tableIndex, int maxCells = 20000)
    {
        dynamic hwp = hwpObject;
        object? original = null;
        try
        {
            dynamic position = hwp.CreateSet("ListParaPos");
            if ((bool)hwp.GetPosBySet(position)) original = (object)position;
        }
        catch { }
        try
        {
            if (!SelectTableCellBlockOnControl(hwpObject, tableObject, tableIndex, 0, 0, 0, out _)) return null;
            var count = 1;
            while (count < maxCells)
            {
                try { hwp.HAction.Run("Cancel"); } catch { }
                if (!(bool)hwp.HAction.Run("TableCellBlock") || !(bool)hwp.HAction.Run("TableRightCell"))
                    return count;
                count++;
            }
            return null;
        }
        catch { return null; }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
            if (original is not null)
                try { hwp.SetPosBySet((dynamic)original); } catch { }
        }
    }

    private static (int List, int Para)? PositionKey(dynamic set)
    {
        try
        {
            var list = Convert.ToInt32(set.Item("List"));
            var para = Convert.ToInt32(set.Item("Para"));
            return (list, para);
        }
        catch { return null; }
    }

    /// <summary>현재 표 셀 목록/문단에 한글 수식 컨트롤(%fmu)이 있는지 확인한다.</summary>
    private static bool IsCurrentTableCellFormula(object hwpObject)
    {
        dynamic hwp = hwpObject;
        try
        {
            foreach (var candidate in new object?[] { hwp.CurSelectedCtrl, hwp.ParentCtrl })
            {
                if (candidate is null) continue;
                dynamic control = candidate;
                if (string.Equals(Convert.ToString(control.CtrlID), "%fmu", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }
        catch { }

        (int List, int Para)? current = null;
        try
        {
            dynamic position = hwp.CreateSet("ListParaPos");
            if ((bool)hwp.GetPosBySet(position)) current = PositionKey(position);
        }
        catch { }
        return current is not null && FormulaCellPositions(hwpObject).Contains(current.Value);
    }

    private static HashSet<(int List, int Para)> FormulaCellPositions(object hwpObject)
    {
        dynamic hwp = hwpObject;
        var result = new HashSet<(int List, int Para)>();
        dynamic? ctrl = null;
        try { ctrl = hwp.HeadCtrl; } catch { }
        while (ctrl is not null)
        {
            try
            {
                if (string.Equals(Convert.ToString(ctrl.CtrlID), "%fmu", StringComparison.OrdinalIgnoreCase))
                {
                    object anchor = (object)ctrl.GetAnchorPos(0);
                    var key = PositionKey((dynamic)anchor);
                    if (key is not null) result.Add(key.Value);
                }
            }
            catch { }
            try { ctrl = ctrl.Next; } catch { ctrl = null; }
        }
        return result;
    }

    private static (int List, int Para)? CurrentListParaPosition(object hwpObject)
    {
        dynamic hwp = hwpObject;
        try
        {
            dynamic position = hwp.CreateSet("ListParaPos");
            return (bool)hwp.GetPosBySet(position) ? PositionKey(position) : null;
        }
        catch { return null; }
    }

    private static HwpWriteResult ExecTableInsertLine(dynamic hwp, JsonObject op, bool rows)
    {
        var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
        var row = Json.GetInt(op, "row") ?? 0;
        var col = Json.GetInt(op, "col") ?? 0;
        var count = Json.GetInt(op, "count") ?? 1;
        if (count is < 1 or > 20) throw new ArgumentOutOfRangeException(nameof(count), "count는 1~20입니다");
        var position = (Json.GetString(op, "position") ?? "after").ToLowerInvariant();
        var actionId = (rows, position) switch
        {
            (true, "before") => "TableInsertUpperRow",
            (true, "after") => "TableInsertLowerRow",
            (false, "before") => "TableInsertLeftColumn",
            (false, "after") => "TableInsertRightColumn",
            _ => throw new ArgumentException("position은 before|after 중 하나여야 합니다"),
        };
        dynamic? table = FindControl(hwp, "tbl", tableIndex);
        if (table is null) return new HwpWriteResult(false, $"table:{tableIndex}", $"표 {tableIndex}을 찾을 수 없습니다");
        var before = TryReadTableDimensions((object)hwp, tableIndex);
        var beforeCells = TryCountTableCellsOnControl((object)hwp, (object)table, tableIndex);
        var previousCells = beforeCells;
        int? perStepCellDelta = null;
        var cellDeltaVerified = beforeCells is not null;
        var completed = 0;
        string? failure = null;
        for (var index = 0; index < count; index++)
        {
            if (!SelectTableCellBlockOnControl((object)hwp, (object)table, tableIndex, row, col, null, out var error))
            {
                failure = $"{index + 1}/{count}회 대상 재선택 실패: {error}";
                break;
            }
            try
            {
                // HWP 2024 일부 빌드는 Count>1을 무시한다. 항상 한 줄씩 실행해
                // API 계약의 count와 실제 구조 변경 수가 같도록 보장한다.
                dynamic insert = hwp.HParameterSet.HTableInsertLine;
                hwp.HAction.GetDefault(actionId, insert.HSet);
                insert.Count = 1;
                if (!(bool)hwp.HAction.Execute(actionId, insert.HSet))
                {
                    failure = $"{index + 1}/{count}회 {actionId} 실행 실패";
                    break;
                }
                completed++;
            }
            finally { try { hwp.HAction.Run("Cancel"); } catch { } }
            var currentCells = TryCountTableCellsOnControl((object)hwp, (object)table, tableIndex);
            if (previousCells is null || currentCells is null)
                cellDeltaVerified = false;
            else
            {
                var delta = currentCells.Value - previousCells.Value;
                if (delta <= 0 || (perStepCellDelta is not null && delta != perStepCellDelta.Value))
                    cellDeltaVerified = false;
                perStepCellDelta ??= delta;
                previousCells = currentCells;
            }
        }

        var after = TryReadTableDimensions((object)hwp, tableIndex);
        int? expectedRows = before is null ? null : before.Rows + (rows ? count : 0);
        int? expectedColumns = before is null ? null : before.Columns + (rows ? 0 : count);
        // 실행 반환값만으로 성공 처리하면 Count 무시 빌드에서 거짓 성공이 된다.
        // 구조를 읽지 못한 경우도 검증 실패로 반환하여 호출자가 재확인하도록 한다.
        var dimensionVerified = before is not null && after is not null &&
            after.Rows == expectedRows && after.Columns == expectedColumns;
        var structureVerified = dimensionVerified ||
            (cellDeltaVerified && completed == count && perStepCellDelta is > 0);
        var ok = completed == count && structureVerified;
        var detail = failure ?? (structureVerified
            ? $"inserted {completed}/{count} {(rows ? "row(s)" : "column(s)")} {position} ({row},{col}); " +
              (dimensionVerified ? "dimension delta verified" : $"cell-count delta verified per step={perStepCellDelta}, total {beforeCells}->{previousCells}")
            : before is null || after is null
                ? $"inserted action completed {completed}/{count}, but table dimensions could not be read; success is not claimed"
                : $"structure readback mismatch: expected {expectedRows}x{expectedColumns}, actual {after.Rows}x{after.Columns}");
        return new HwpWriteResult(ok, $"table:{tableIndex}", detail,
            before is null ? null : $"{before.Rows}x{before.Columns}",
            after is null ? null : $"{after.Rows}x{after.Columns}");
    }

    private static bool SelectTableRowBlock(object hwpObject, int tableIndex, int row, out string error)
    {
        if (!SelectTableCellBlock(hwpObject, tableIndex, row, 0, out error)) return false;
        return ExtendSelectedTableRow(hwpObject, tableIndex, row, out error);
    }

    private static bool SelectTableRowBlock(
        object hwpObject, object tableObject, int tableIndex, int row, out string error)
    {
        if (!SelectTableCellBlockOnControl(hwpObject, tableObject, tableIndex, row, 0, null, out error)) return false;
        return ExtendSelectedTableRow(hwpObject, tableIndex, row, out error);
    }

    private static bool ExtendSelectedTableRow(object hwpObject, int tableIndex, int row, out string error)
    {
        dynamic hwp = hwpObject;
        error = "";
        if (!(bool)hwp.HAction.Run("TableCellBlockExtend"))
        {
            error = $"표 {tableIndex}의 {row}행 선택 확장에 실패했습니다";
            return false;
        }
        if (!(bool)hwp.HAction.Run("TableCellBlockRow"))
        {
            error = $"표 {tableIndex}의 {row}행 전체를 선택하지 못했습니다";
            return false;
        }
        return true;
    }

    private static double? ReadSelectedRowHeightMm(object hwpObject)
    {
        dynamic hwp = hwpObject;
        try
        {
            dynamic shape = hwp.HParameterSet.HShapeObject;
            hwp.HAction.GetDefault("TablePropertyDialog", shape.HSet);
            double units;
            try { units = Convert.ToDouble(shape.ShapeTableCell.Height); }
            catch { units = Convert.ToDouble(shape.HSet.Item("Height")); }
            var unitsPerMillimeter = Convert.ToDouble(hwp.MiliToHwpUnit(1.0));
            return unitsPerMillimeter > 0 ? units / unitsPerMillimeter : null;
        }
        catch { return null; }
    }

    /// <summary>
    /// 한컴 공식 TablePropertyDialog/ShapeTableCell.Height 경로로 행 높이를 mm 단위로 지정한다.
    /// 단일 셀만 고치면 같은 행의 다른 셀과 경계가 어긋날 수 있으므로 반드시 TableCellBlockRow로
    /// 행 전체를 선택하고, 적용 뒤 같은 행을 다시 선택해 실제 높이를 읽어 검증한다.
    /// </summary>
    private static HwpWriteResult ExecTableSetRowHeight(dynamic hwp, JsonObject op)
        => ExecTableSetRowHeightCore((object)hwp, op, null);

    private static HwpWriteResult ExecTableSetRowHeightCore(
        object hwpObject, JsonObject op, object? tableControl)
    {
        dynamic hwp = hwpObject;
        var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
        var row = Json.GetInt(op, "row") ?? throw new ArgumentException("table_set_row_height.row가 필요합니다");
        if (!TryJsonNumber(op, "heightMm", out var heightMm) || heightMm is < 4 or > 50)
            throw new ArgumentOutOfRangeException("heightMm", "table_set_row_height.heightMm는 4~50mm입니다");
        var selected = tableControl is null
            ? SelectTableRowBlock(hwpObject, tableIndex, row, out var error)
            : SelectTableRowBlock(hwpObject, tableControl, tableIndex, row, out error);
        if (!selected)
            return new HwpWriteResult(false, $"table:{tableIndex}/row:{row}", error);

        double? before = ReadSelectedRowHeightMm((object)hwp);
        dynamic shape = hwp.HParameterSet.HShapeObject;
        hwp.HAction.GetDefault("TablePropertyDialog", shape.HSet);
        shape.HSet.SetItem("ShapeType", 3);
        shape.HSet.SetItem("ShapeCellSize", 1);
        shape.ShapeTableCell.Height = hwp.MiliToHwpUnit(heightMm);
        var executed = (bool)hwp.HAction.Execute("TablePropertyDialog", shape.HSet);
        try { hwp.HAction.Run("Cancel"); } catch { }
        if (!executed)
            return new HwpWriteResult(false, $"table:{tableIndex}/row:{row}",
                "TablePropertyDialog 실행 실패", before?.ToString("0.00"), heightMm.ToString("0.00"));

        selected = tableControl is null
            ? SelectTableRowBlock(hwpObject, tableIndex, row, out error)
            : SelectTableRowBlock(hwpObject, tableControl, tableIndex, row, out error);
        if (!selected)
            return new HwpWriteResult(false, $"table:{tableIndex}/row:{row}",
                $"적용 후 행 재선택 실패: {error}", before?.ToString("0.00"), null);
        double? after = ReadSelectedRowHeightMm((object)hwp);
        try { hwp.HAction.Run("Cancel"); } catch { }
        // 한글은 셀 내용이 들어갈 최소 높이보다 작은 값은 자동으로 위쪽 보정한다.
        // 따라서 목표값보다 작지 않은지 검증하고 실제 적용값을 반드시 반환한다.
        var verified = after is not null && after.Value >= heightMm - 0.6;
        var heightDetail = after is not null && Math.Abs(after.Value - heightMm) <= 0.6
            ? $"row height {after:0.00}mm verified"
            : $"row height {after:0.00}mm verified (content minimum exceeded requested {heightMm:0.00}mm)";
        return new HwpWriteResult(verified, $"table:{tableIndex}/row:{row}",
            verified ? heightDetail : $"row height readback mismatch; requested minimum={heightMm:0.00}mm, actual={after:0.00}mm",
            before?.ToString("0.00"), after?.ToString("0.00"));
    }

    internal static IReadOnlyList<HwpRowHeightSpec> ParseRowHeightSpecs(JsonObject op)
    {
        var rows = Json.GetArr(op, "rows") ??
            throw new ArgumentException("table_set_row_heights.rows 배열이 필요합니다");
        if (rows.Count is < 1 or > 100)
            throw new ArgumentOutOfRangeException("rows", "table_set_row_heights.rows는 1~100개입니다");

        var result = new List<HwpRowHeightSpec>(rows.Count);
        var seen = new HashSet<int>();
        for (var index = 0; index < rows.Count; index++)
        {
            if (rows[index] is not JsonObject item)
                throw new ArgumentException($"table_set_row_heights.rows[{index}]는 객체여야 합니다");
            var row = Json.GetInt(item, "row") ??
                throw new ArgumentException($"table_set_row_heights.rows[{index}].row가 필요합니다");
            if (row < 0)
                throw new ArgumentOutOfRangeException("row", $"rows[{index}].row는 0 이상이어야 합니다");
            if (!TryJsonNumber(item, "heightMm", out var heightMm) || heightMm is < 4 or > 50)
                throw new ArgumentOutOfRangeException("heightMm", $"rows[{index}].heightMm는 4~50mm입니다");
            if (!seen.Add(row))
                throw new ArgumentException($"table_set_row_heights.rows에 row={row}가 중복되었습니다");
            result.Add(new HwpRowHeightSpec(row, heightMm));
        }
        return result;
    }

    private static IReadOnlyList<HwpWriteResult> ExecTableSetRowHeights(dynamic hwp, JsonObject op)
    {
        var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
        if (tableIndex < 0) throw new ArgumentOutOfRangeException("tableIndex");
        var specs = ParseRowHeightSpecs(op);
        dynamic? table = FindControl(hwp, "tbl", tableIndex);
        if (table is null)
            return new[] { new HwpWriteResult(false, $"table:{tableIndex}", $"표 {tableIndex}을 찾을 수 없습니다") };
        var dimensions = TryReadTableDimensions((object)hwp, tableIndex);
        if (dimensions is not null && specs.Any(spec => spec.Row >= dimensions.Rows))
            throw new ArgumentOutOfRangeException("rows",
                $"표 {tableIndex}의 유효 행은 0..{dimensions.Rows - 1}입니다");

        var results = new List<HwpWriteResult>(specs.Count);
        foreach (var spec in specs)
        {
            var single = new JsonObject
            {
                ["tableIndex"] = tableIndex,
                ["row"] = spec.Row,
                ["heightMm"] = spec.HeightMm,
            };
            var result = ExecTableSetRowHeightCore((object)hwp, single, (object)table);
            results.Add(result);
            if (!result.Ok) break;
        }
        return results;
    }

    private static HwpWriteResult ExecTableDeleteLine(dynamic hwp, JsonObject op, bool rows)
    {
        var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
        var row = Json.GetInt(op, "row") ?? 0;
        var col = Json.GetInt(op, "col") ?? 0;
        var count = Json.GetInt(op, "count") ?? 1;
        if (count is < 1 or > 20) throw new ArgumentOutOfRangeException(nameof(count), "count는 1~20입니다");
        dynamic? table = FindControl(hwp, "tbl", tableIndex);
        if (table is null) return new HwpWriteResult(false, $"table:{tableIndex}", $"표 {tableIndex}을 찾을 수 없습니다");
        var before = TryReadTableDimensions((object)hwp, tableIndex);
        var beforeCells = TryCountTableCellsOnControl((object)hwp, (object)table, tableIndex);
        var previousCells = beforeCells;
        int? perStepCellDelta = null;
        var cellDeltaVerified = beforeCells is not null;
        if (before is not null)
        {
            var available = rows ? before.Rows - row : before.Columns - col;
            var remaining = (rows ? before.Rows : before.Columns) - count;
            if (available < count)
                throw new ArgumentOutOfRangeException(nameof(count), $"삭제 시작 위치부터 남은 {(rows ? "행" : "열")}은 {available}개입니다");
            if (remaining < 1)
                throw new ArgumentOutOfRangeException(nameof(count), "표에는 최소 한 행과 한 열이 남아야 합니다");
        }
        var actionId = rows ? "TableDeleteRow" : "TableDeleteColumn";
        var completed = 0;
        string? failure = null;
        int? messageMode = null;
        try
        {
            try { messageMode = Convert.ToInt32(hwp.GetMessageBoxMode()); hwp.SetMessageBoxMode(0x00002000); } catch { }
            for (var index = 0; index < count; index++)
            {
                if (!SelectTableCellBlockOnControl((object)hwp, (object)table, tableIndex, row, col, null, out var error))
                {
                    failure = $"{index + 1}/{count}회 대상 재선택 실패: {error}";
                    break;
                }
                try
                {
                    dynamic delete = hwp.HParameterSet.HTableDeleteLine;
                    hwp.HAction.GetDefault(actionId, delete.HSet);
                    delete.Type = rows ? 0 : 1;
                    if (!(bool)hwp.HAction.Execute(actionId, delete.HSet))
                    {
                        failure = $"{index + 1}/{count}회 {actionId} 실행 실패";
                        break;
                    }
                    completed++;
                }
                finally { try { hwp.HAction.Run("Cancel"); } catch { } }
                var currentCells = TryCountTableCellsOnControl((object)hwp, (object)table, tableIndex);
                if (previousCells is null || currentCells is null)
                    cellDeltaVerified = false;
                else
                {
                    var delta = previousCells.Value - currentCells.Value;
                    if (delta <= 0 || (perStepCellDelta is not null && delta != perStepCellDelta.Value))
                        cellDeltaVerified = false;
                    perStepCellDelta ??= delta;
                    previousCells = currentCells;
                }
            }
        }
        finally { try { hwp.SetMessageBoxMode(messageMode ?? 0xFFFFF); } catch { } }

        var after = TryReadTableDimensions((object)hwp, tableIndex);
        int? expectedRows = before is null ? null : before.Rows - (rows ? count : 0);
        int? expectedColumns = before is null ? null : before.Columns - (rows ? 0 : count);
        var dimensionVerified = before is not null && after is not null &&
            after.Rows == expectedRows && after.Columns == expectedColumns;
        var structureVerified = dimensionVerified ||
            (cellDeltaVerified && completed == count && perStepCellDelta is > 0);
        var ok = completed == count && structureVerified;
        var detail = failure ?? (structureVerified
            ? $"deleted {completed}/{count} {(rows ? "row(s)" : "column(s)")} at ({row},{col}); " +
              (dimensionVerified ? "dimension delta verified" : $"cell-count delta verified per step={perStepCellDelta}, total {beforeCells}->{previousCells}")
            : before is null || after is null
                ? $"delete action completed {completed}/{count}, but table dimensions could not be read; success is not claimed"
                : $"structure readback mismatch: expected {expectedRows}x{expectedColumns}, actual {after.Rows}x{after.Columns}");
        return new HwpWriteResult(ok, $"table:{tableIndex}", detail,
            before is null ? null : $"{before.Rows}x{before.Columns}",
            after is null ? null : $"{after.Rows}x{after.Columns}");
    }

    private static HwpWriteResult ExecTableMergeCells(dynamic hwp, JsonObject op)
    {
        var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
        var startRow = Json.GetInt(op, "startRow") ?? throw new ArgumentException("startRow가 필요합니다");
        var startCol = Json.GetInt(op, "startCol") ?? throw new ArgumentException("startCol이 필요합니다");
        var endRow = Json.GetInt(op, "endRow") ?? throw new ArgumentException("endRow가 필요합니다");
        var endCol = Json.GetInt(op, "endCol") ?? throw new ArgumentException("endCol이 필요합니다");
        if (endRow < startRow || endCol < startCol || (endRow == startRow && endCol == startCol))
            throw new ArgumentException("병합 범위는 시작 셀보다 오른쪽/아래의 두 셀 이상이어야 합니다");
        if (!SelectTableCellBlock((object)hwp, tableIndex, startRow, startCol, out var error))
            return new HwpWriteResult(false, $"table:{tableIndex}", error);
        hwp.HAction.Run("TableCellBlockExtend");
        for (var i = startRow; i < endRow; i++) hwp.HAction.Run("TableLowerCell");
        for (var i = startCol; i < endCol; i++) hwp.HAction.Run("TableRightCell");
        var ok = (bool)hwp.HAction.Run("TableMergeCell");
        try { hwp.HAction.Run("Cancel"); } catch { }
        return new HwpWriteResult(ok, $"table:{tableIndex}", $"merged ({startRow},{startCol})-({endRow},{endCol})");
    }

    private static HwpWriteResult ExecSetFieldText(dynamic hwp, JsonObject op)
    {
        var name = Json.GetString(op, "name") ?? throw new ArgumentException("set_field_text.name이 필요합니다");
        var text = Json.GetString(op, "text") ?? "";
        if (!(bool)hwp.FieldExist(name)) return new HwpWriteResult(false, $"field:{name}", "필드를 찾을 수 없습니다");
        var before = (string)(hwp.GetFieldText(name) ?? "");
        hwp.PutFieldText(name, text);
        var after = (string)(hwp.GetFieldText(name) ?? "");
        var verified = string.Equals(after.TrimEnd('\u0002'), text, StringComparison.Ordinal);
        return new HwpWriteResult(verified, $"field:{name}", verified ? "field text replaced" : "field readback mismatch", before, after);
    }

    private static JsonObject InspectFields(dynamic hwp, int maxFields, bool includeValues)
    {
        string list;
        try { list = (string)(hwp.GetFieldList(1, 0) ?? ""); }
        catch { list = ""; }
        var names = list.Split('\u0002', StringSplitOptions.RemoveEmptyEntries);
        var fields = new JsonArray();
        foreach (var name in names.Take(maxFields))
        {
            var item = new JsonObject { ["name"] = name };
            if (includeValues)
            {
                try { item["text"] = ((string)(hwp.GetFieldText(name) ?? "")).TrimEnd('\u0002'); }
                catch (Exception ex) { item["error"] = ex.Message; }
            }
            fields.Add(item);
        }
        return new JsonObject
        {
            ["fields"] = fields,
            ["fieldCount"] = names.Length,
            ["truncated"] = names.Length > maxFields,
            ["valuesIncluded"] = includeValues,
        };
    }
}
