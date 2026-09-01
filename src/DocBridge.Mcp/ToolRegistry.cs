using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Mcp;

/// <summary>
/// MCP tool 레지스트리 (명령서 §6). tool 이름은 밑줄(_)만 사용한다.
/// 모든 handler는 DocBridgeHost의 단일 진입점으로 위임되므로
/// dry-run → snapshot → confirmToken → apply → readback 안전 순서가 강제된다.
/// </summary>
public sealed class ToolRegistry
{
    public sealed record ToolDef(string Name, string Description, JsonObject InputSchema,
        Func<JsonObject, JsonObject> Handler);

    private readonly List<ToolDef> _tools;

    public ToolRegistry(DocBridgeHost host)
    {
        var hwpJobs = new HwpOperationJobManager(host);
        JsonObject NoInput(string desc) => new()
        {
            ["type"] = "object",
            ["description"] = desc,
            ["properties"] = new JsonObject(),
        };

        JsonObject ApplyOpsSchema(string app, string allowedOps) => new()
        {
            ["type"] = "object",
            ["description"] = $"{app} 쓰기 ops 배치. 허용 op: {allowedOps}. " +
                              "dryRun=true로 diff+confirmToken을 받은 뒤, 사용자 승인 시 같은 ops를 dryRun=false+confirmToken으로 재호출한다.",
            ["properties"] = new JsonObject
            {
                ["ops"] = new JsonObject
                {
                    ["type"] = "array",
                    ["description"] = "Operation[] (공통 스키마 §7)",
                    ["items"] = new JsonObject { ["type"] = "object", ["properties"] = new JsonObject { ["op"] = new JsonObject { ["type"] = "string" } }, ["required"] = new JsonArray("op") },
                },
                ["dryRun"] = new JsonObject { ["type"] = "boolean", ["description"] = "true면 미적용 diff/confirmToken 발급" },
                ["confirmToken"] = new JsonObject { ["type"] = "string", ["description"] = "직전 dry-run이 발급한 토큰 (dryRun=false 시 필수)" },
                ["highRiskConfirm"] = new JsonObject { ["type"] = "boolean", ["description"] = "고위험 op 포함 시 사용자 명시 승인 표시" },
            },
            ["required"] = new JsonArray("ops", "dryRun"),
        };

        JsonObject ExcelApplyOpsSchema()
        {
            var schema = ApplyOpsSchema("excel",
                "set_values, set_formulas, insert_rows, insert_cols, format_range, find_replace, copy_sheet, " +
                "merge_cells, unmerge_cells, set_rows_hidden, set_cols_hidden, set_sheet_visibility");
            schema["description"] =
                "Excel 쓰기 배치. 쓰기 대상 시트는 절대로 활성 시트로 추정하지 않습니다. " +
                "set_values/set_formulas/format_range는 target.sheet 또는 '시트 이름'!A1 형식의 range가 필요하고, " +
                "insert_rows/insert_cols, 숨김/표시 작업 및 sheet 범위 find_replace는 target.sheet가 필요합니다. " +
                "merge_cells는 좌상단 외 셀에 값/수식이 있으면 데이터 손실 방지를 위해 거부합니다. " +
                "활성 시트와 마지막 표시 시트는 숨기지 않습니다. " +
                "find_replace의 target.scope='workbook'과 copy_sheet는 자체적으로 대상을 명시합니다. " +
                "dryRun=true로 diff+confirmToken을 받은 뒤 동일한 ops를 dryRun=false+confirmToken으로 다시 호출합니다.";

            var items = (JsonObject)((JsonObject)((JsonObject)schema["properties"]!)["ops"]!)["items"]!;
            items["properties"] = new JsonObject
            {
                ["op"] = new JsonObject
                {
                    ["type"] = "string",
                    ["enum"] = new JsonArray(
                        "set_values", "set_formulas", "insert_rows", "insert_cols",
                        "format_range", "find_replace", "copy_sheet",
                        "merge_cells", "unmerge_cells", "set_rows_hidden", "set_cols_hidden",
                        "set_sheet_visibility"),
                },
                ["range"] = new JsonObject
                {
                    ["type"] = "string",
                    ["description"] = "A1 범위. target.sheet를 생략할 때는 예: '공사 내역'!B2:D5처럼 시트명을 포함해야 합니다.",
                },
                ["values"] = new JsonObject
                {
                    ["type"] = "array",
                    ["description"] = "set_values용 직사각형 2차원 배열. Excel 숫자는 COM Double로 정규화됩니다.",
                    ["items"] = new JsonObject { ["type"] = "array" },
                },
                ["formulas"] = new JsonObject
                {
                    ["type"] = "array",
                    ["description"] = "set_formulas용 직사각형 2차원 수식 배열.",
                    ["items"] = new JsonObject { ["type"] = "array" },
                },
                ["row"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1 },
                ["col"] = new JsonObject
                {
                    ["description"] = "열 번호(1부터) 또는 열 문자(A, B, ...).",
                    ["oneOf"] = new JsonArray(
                        new JsonObject { ["type"] = "integer", ["minimum"] = 1 },
                        new JsonObject { ["type"] = "string" }),
                },
                ["count"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1 },
                ["hidden"] = new JsonObject
                {
                    ["type"] = "boolean",
                    ["description"] = "set_rows_hidden/set_cols_hidden에서 true=숨김, false=표시.",
                },
                ["visibility"] = new JsonObject
                {
                    ["type"] = "string",
                    ["enum"] = new JsonArray("visible", "hidden"),
                    ["description"] = "set_sheet_visibility의 일반 표시 상태. veryHidden은 지원하지 않습니다.",
                },
                ["target"] = new JsonObject
                {
                    ["type"] = "object",
                    ["description"] = "쓰기 대상. 시트 범위 쓰기에는 sheet를 명시합니다.",
                    ["properties"] = new JsonObject
                    {
                        ["sheet"] = new JsonObject { ["type"] = "string", ["description"] = "정확한 워크시트 이름." },
                        ["workbook"] = new JsonObject { ["type"] = "string", ["description"] = "열린 대상 workbook 이름 또는 절대 경로." },
                        ["scope"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("sheet", "workbook") },
                    },
                },
                ["style"] = new JsonObject { ["type"] = "object" },
                ["find"] = new JsonObject { ["type"] = "string" },
                ["replace"] = new JsonObject { ["type"] = "string" },
                ["options"] = new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["matchCase"] = new JsonObject { ["type"] = "boolean" },
                    },
                },
                ["sourceWorkbook"] = new JsonObject { ["type"] = "string" },
                ["sourceSheet"] = new JsonObject { ["type"] = "string" },
                ["targetSheet"] = new JsonObject { ["type"] = "string" },
                ["targetWorkbook"] = new JsonObject { ["type"] = "string" },
            };
            return schema;
        }

        JsonObject HwpApplyOpsSchema()
        {
            var schema = ApplyOpsSchema("hwp",
                "insert_text, append_text, insert_before_text, insert_after_text, replace_document_text, replace_selection, find_replace, " +
                "set_paragraph_style_basic, set_paragraph_format, format_paragraphs(bulk), set_page_setup, insert_break, " +
                "insert_table, table_cell_set_text, table_set_cells(batch), table_insert_rows/columns, table_delete_rows/columns(high-risk), " +
                "table_merge_cells, table_set_row_height(mm), table_set_row_heights(bulk), set_field_text, insert_picture, insert_page_number, " +
                "set_header_footer_text, export_pdf(high-risk)");
            schema["description"] =
                "한글 쓰기 배치. insert_text/replace_selection은 현재 커서, append_text는 문서 끝에 추가합니다. " +
                "모든 텍스트 쓰기는 기본적으로 기존 글자, 위·아래 문단, 반복 양식의 같은 라벨/값 셀을 비교해 글자와 문단 서식을 보존합니다. " +
                "insert_before_text/insert_after_text는 고유한 기준 문구 또는 occurrence로 지정한 항목의 앞뒤에 삽입합니다. " +
                "replace_document_text는 전체 본문을 교체합니다. text 안의 \\n은 실제 한글 문단으로 입력되므로 다중 문단을 한 op로 처리할 수 있습니다. " +
                "반복 글자+문단 서식은 format_paragraphs.items로 묶고, 여러 셀은 table_set_cells.cells, 여러 행 높이는 table_set_row_heights.rows로 묶어 한 op로 실행·검증합니다. " +
                "table_insert/delete_rows/columns의 count는 요청한 횟수만큼 한 줄씩 실행·검증합니다. insert_picture에 tableIndex와 row+col 또는 cellIndex를 주면 표 셀 안에 그림을 넣습니다. " +
                "dryRun은 같은 batch의 앞 op가 만든 본문과 표 구조를 뒤 op가 이어받아 순차 시뮬레이션합니다. " +
                "활성 문서는 대상을 생략하고, 여러 열린 문서 중 하나는 hwp_get_active_context.openDocuments의 documentRef를 모든 op에 동일하게 지정합니다. " +
                "파일 작업은 모든 op에 같은 절대 file을 지정하며 file과 documentRef는 함께 쓰지 않습니다. " +
                "PowerShell/Python COM 우회나 새 한글 프로세스 반복 실행은 금지됩니다. " +
                "dryRun=true로 diff+confirmToken을 받은 뒤 같은 ops로 적용합니다.";
            var properties = (JsonObject)((JsonObject)((JsonObject)schema["properties"]!)["ops"]!)["items"]!;
            properties["properties"] = new JsonObject
            {
                ["op"] = new JsonObject
                {
                    ["type"] = "string",
                    ["enum"] = new JsonArray(
                        "insert_text", "append_text", "insert_before_text", "insert_after_text",
                        "replace_document_text", "replace_selection", "find_replace",
                        "set_paragraph_style_basic", "set_paragraph_format", "format_paragraphs", "set_page_setup", "insert_break",
                        "insert_table", "table_cell_set_text", "table_set_cells", "table_insert_rows", "table_insert_columns",
                        "table_delete_rows", "table_delete_columns", "table_merge_cells", "table_set_row_height", "table_set_row_heights", "set_field_text",
                        "insert_picture", "insert_page_number", "set_header_footer_text", "export_pdf"),
                },
                ["text"] = new JsonObject
                {
                    ["type"] = "string",
                    ["description"] = "텍스트 쓰기 op에 사용할 본문. 줄바꿈(\\n)은 실제 문단으로 보존됩니다.",
                },
                ["startNewParagraph"] = new JsonObject
                {
                    ["type"] = "boolean",
                    ["description"] = "append_text 전용. 기본 true이며 기존 본문 뒤에서 새 문단으로 시작합니다.",
                },
                ["anchor"] = new JsonObject
                {
                    ["type"] = "string",
                    ["description"] = "insert_before_text/insert_after_text의 기준 문구. occurrence를 생략하면 문서에서 정확히 1개여야 합니다.",
                },
                ["occurrence"] = new JsonObject
                {
                    ["type"] = "integer",
                    ["minimum"] = 1,
                    ["description"] = "기준 문구가 여러 개일 때 사용할 1부터 시작하는 순번.",
                },
                ["matchCase"] = new JsonObject
                {
                    ["type"] = "boolean",
                    ["description"] = "기준 문구 및 find_replace의 대소문자 구분 여부.",
                },
                ["find"] = new JsonObject
                {
                    ["type"] = "string",
                    ["minLength"] = 1,
                    ["description"] = "find_replace에서 찾을 문자열. GetTextFile이 직렬화한 HTML 숫자 엔터티도 기본적으로 원문 Unicode로 복원합니다.",
                },
                ["replace"] = new JsonObject
                {
                    ["type"] = "string",
                    ["description"] = "find_replace에서 바꿀 문자열.",
                },
                ["scope"] = new JsonObject
                {
                    ["type"] = "object",
                    ["description"] = "find_replace 선택 범위. 생략 시 문서 전체. 문단은 startParagraph+endParagraph, 표 셀은 tableIndex+(row+col 또는 cellIndex)를 지정합니다.",
                    ["properties"] = new JsonObject
                    {
                        ["startParagraph"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                        ["endParagraph"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                        ["tableIndex"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                        ["row"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                        ["col"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                        ["cellIndex"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                    },
                },
                ["mode"] = new JsonObject
                {
                    ["type"] = "string",
                    ["enum"] = new JsonArray("paragraph", "inline"),
                    ["description"] = "paragraph(기본)은 기준 문단 앞/뒤에 새 문단으로 삽입하고 인접 문단 서식을 상속. inline은 기준 문구 바로 앞/뒤에 삽입.",
                },
                ["file"] = new JsonObject
                {
                    ["type"] = "string",
                    ["description"] = "선택적 절대 .hwp/.hwpx 경로. 생략하면 사용자가 이미 열어 둔 활성 문서이며 DocBridge는 빈 한글 창을 자동 실행하지 않습니다.",
                },
                ["documentRef"] = new JsonObject
                {
                    ["type"] = "string",
                    ["description"] = "선택 사항. hwp_get_active_context.summary.openDocuments가 반환한 documentRef 또는 instanceRef. 여러 열린 창/탭 중 정확한 라이브 문서를 선택합니다. file과 함께 쓰지 않습니다.",
                },
                ["target"] = new JsonObject { ["type"] = "object" },
                ["style"] = new JsonObject
                {
                    ["type"] = "object",
                    ["description"] = "선택 사항. 문맥에서 복사한 서식보다 우선하는 글자/문단 서식.",
                },
                ["items"] = new JsonObject
                {
                    ["type"] = "array",
                    ["minItems"] = 1,
                    ["maxItems"] = 100,
                    ["description"] = "format_paragraphs 전용. 각 항목은 target과 characterStyle/paragraphStyle 중 하나 이상을 가집니다.",
                    ["items"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["properties"] = new JsonObject
                        {
                            ["target"] = new JsonObject { ["type"] = "object" },
                            ["characterStyle"] = new JsonObject { ["type"] = "object" },
                            ["paragraphStyle"] = new JsonObject { ["type"] = "object" },
                        },
                    },
                },
                ["rows"] = new JsonObject
                {
                    ["type"] = "array",
                    ["minItems"] = 1,
                    ["maxItems"] = 500,
                    ["description"] = "insert_table의 표 행 배열 또는 table_set_row_heights의 [{row:0,heightMm:8.0}, ...] 배열입니다. 후자는 런타임에서 최대 100개로 제한됩니다.",
                },
                ["cells"] = new JsonObject
                {
                    ["type"] = "array",
                    ["minItems"] = 1,
                    ["maxItems"] = 500,
                    ["description"] = "table_set_cells 전용. 각 항목은 text와 row+col 또는 cellIndex를 가지며, 표 컨트롤을 한 번만 찾아 순서대로 적용·정확 재읽기합니다.",
                    ["items"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["properties"] = new JsonObject
                        {
                            ["row"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                            ["col"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                            ["cellIndex"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                            ["text"] = new JsonObject { ["type"] = "string" },
                            ["preserveStyle"] = new JsonObject { ["type"] = "boolean" },
                            ["style"] = new JsonObject { ["type"] = "object" },
                        },
                        ["required"] = new JsonArray("text"),
                    },
                },
                ["preserveStyle"] = new JsonObject
                {
                    ["type"] = "boolean",
                    ["description"] = "기본 true. 기존 값과 위·아래 문맥 및 반복 양식에서 글자/문단 서식을 추론해 보존합니다.",
                },
                ["styleSource"] = new JsonObject
                {
                    ["type"] = "object",
                    ["description"] = "문맥 후보가 충돌할 때만 사용. text+occurrence 또는 tableIndex+cellIndex로 복사할 서식 원본을 지정합니다.",
                },
                ["tableIndex"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                ["row"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                ["col"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0 },
                ["count"] = new JsonObject
                {
                    ["type"] = "integer", ["minimum"] = 1, ["maximum"] = 20,
                    ["description"] = "표 행/열 삽입·삭제 개수. 한글의 Count 무시 문제를 피하기 위해 한 줄씩 반복 실행하고 구조를 검증합니다.",
                },
                ["position"] = new JsonObject
                {
                    ["type"] = "string", ["enum"] = new JsonArray("before", "after"),
                    ["description"] = "표 행/열 삽입 위치.",
                },
                ["path"] = new JsonObject
                {
                    ["type"] = "string",
                    ["description"] = "insert_picture의 로컬 이미지 절대 경로.",
                },
                ["embedded"] = new JsonObject { ["type"] = "boolean" },
                ["sizeOption"] = new JsonObject
                {
                    ["type"] = "string", ["enum"] = new JsonArray("real", "specific", "cell", "cell-ratio"),
                    ["description"] = "그림 크기. 표 셀 대상의 기본값은 cell-ratio, 일반 삽입은 real입니다.",
                },
                ["widthMm"] = new JsonObject { ["type"] = "number", ["exclusiveMinimum"] = 0 },
                ["heightMm"] = new JsonObject
                {
                    ["type"] = "number",
                    ["exclusiveMinimum"] = 0,
                    ["description"] = "table_set_row_height는 4~50mm, insert_picture specific은 양수 높이(mm).",
                },
                ["effect"] = new JsonObject
                {
                    ["type"] = "string", ["enum"] = new JsonArray("original", "grayscale", "black-white"),
                },
                ["reverse"] = new JsonObject { ["type"] = "boolean" },
                ["watermark"] = new JsonObject { ["type"] = "boolean" },
                ["clearCell"] = new JsonObject
                {
                    ["type"] = "boolean",
                    ["description"] = "표 셀 그림 삽입 전 기존 셀 내용을 지울지 여부. 기본 false.",
                },
                ["cellIndex"] = new JsonObject
                {
                    ["type"] = "integer",
                    ["minimum"] = 0,
                    ["description"] = "표의 실제 셀 이동 순서(0부터). 병합 표에서는 row/col 대신 이 값을 사용합니다.",
                },
            };
            return schema;
        }

        _tools = new List<ToolDef>
        {
            // ---------- core (§6.1) ----------
            new("core_ping", "doc-bridge 서버 생존/버전/등록 어댑터 확인",
                NoInput("입력 없음"),
                _ => host.CorePing()),

            new("core_get_status", "실행 중인 어댑터, 연결된 프로그램, 현재 문서 요약",
                NoInput("입력 없음"),
                _ => host.CoreGetStatus()),

            new("core_disconnect", "서버를 종료하지 않고 지정 앱의 COM 연결을 해제합니다. Excel은 DocBridge가 만든 인스턴스만 안전 조건에서 종료하며 사용자 Excel은 절대 종료하지 않습니다.",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["app"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("excel") },
                    },
                    ["required"] = new JsonArray("app"),
                },
                a => host.CoreDisconnect(a)),

            new("core_get_capabilities", "Excel/HWP/AutoCAD의 지원 명령, 자동화 방식, 연결 상태, 안전 기능과 처리 한계를 작업 전에 조회",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["app"] = new JsonObject
                        {
                            ["type"] = "string",
                            ["enum"] = new JsonArray("excel", "hwp", "cad"),
                            ["description"] = "생략하면 세 앱 전체를 반환",
                        },
                    },
                },
                a => host.CoreGetCapabilities(a)),

            new("core_create_snapshot", "지정 app의 현재 문서 스냅샷 생성",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["app"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("excel", "hwp", "cad") },
                        ["reason"] = new JsonObject { ["type"] = "string" },
                    },
                    ["required"] = new JsonArray("app"),
                },
                a => host.CoreCreateSnapshot(a)),

            new("core_list_snapshots", "스냅샷 목록 조회",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["app"] = new JsonObject { ["type"] = "string" },
                        ["limit"] = new JsonObject { ["type"] = "integer" },
                    },
                },
                a => host.CoreListSnapshots(a)),

            new("core_restore_snapshot", "[고위험] 스냅샷 복원. confirmToken 없이 호출하면 dry-run으로 토큰을 발급하고, 토큰과 함께 재호출해야 실제 복원된다.",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["snapshotId"] = new JsonObject { ["type"] = "string" },
                        ["confirmToken"] = new JsonObject { ["type"] = "string" },
                    },
                    ["required"] = new JsonArray("snapshotId"),
                },
                a => host.CoreRestoreSnapshot(a)),

            // ---------- excel (§6.2) ----------
            new("excel_get_active_context", "실행 중인 Excel의 활성 workbook/worksheet, 선택 범위, 시트 목록, 사용 범위 요약. Excel이나 workbook이 없으면 새 창을 만들지 않고 오류를 반환합니다.",
                NoInput("입력 없음"),
                _ => host.GetActiveContext("excel")),

            new("excel_read_range", "범위 값/수식 및 선택적 병합·행열 숨김 상태 읽기",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["range"] = new JsonObject { ["type"] = "string", ["description"] = "예: A1:B3" },
                        ["workbook"] = new JsonObject { ["type"] = "string", ["description"] = "선택 사항: 열려 있는 workbook의 이름 또는 절대 경로. 경로만으로 닫힌 파일을 자동으로 열지 않습니다." },
                        ["allowOpenFile"] = new JsonObject
                        {
                            ["type"] = "boolean",
                            ["description"] = "기본 false. 사용자가 닫힌 파일을 열어 읽으라고 명시한 경우에만 true로 설정합니다. 존재하는 절대 workbook 경로가 함께 있어야 하며, 쓰기에는 사용할 수 없습니다.",
                        },
                        ["sheet"] = new JsonObject { ["type"] = "string" },
                        ["includeFormulas"] = new JsonObject { ["type"] = "boolean" },
                        ["includeStyles"] = new JsonObject { ["type"] = "boolean" },
                        ["includeLayout"] = new JsonObject
                        {
                            ["type"] = "boolean",
                            ["description"] = "true면 mergedAreas, 행/열 hidden 상태, 시트 visibility를 제한 범위 안에서 함께 반환합니다.",
                        },
                    },
                    ["required"] = new JsonArray("range"),
                },
                a => host.Read("excel", a)),

            new("excel_inspect", "Excel workbook 구조·표/차트/도형/피벗·수식 오류 또는 제한된 보기/모달 상태를 비파괴 진단",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["scope"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("scan", "objects", "errors", "diagnostics") },
                        ["workbook"] = new JsonObject { ["type"] = "string", ["description"] = "diagnostics 외 scope의 선택적 열린 workbook 이름 또는 절대 경로. 경로만으로 닫힌 파일을 자동으로 열지 않습니다." },
                        ["allowOpenFile"] = new JsonObject
                        {
                            ["type"] = "boolean",
                            ["description"] = "기본 false. 사용자가 닫힌 파일을 열어 비파괴 검사하라고 명시한 경우에만 true로 설정합니다. diagnostics에는 적용되지 않습니다.",
                        },
                        ["sheet"] = new JsonObject { ["type"] = "string" },
                        ["limit"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1, ["maximum"] = 2000 },
                    },
                    ["required"] = new JsonArray("scope"),
                },
                a => host.Read("excel", a)),

            new("excel_apply_ops", "Excel 쓰기 ops 적용 (dry-run → confirmToken → apply 안전 흐름)",
                ExcelApplyOpsSchema(),
                a => host.ApplyOps("excel", a)),

            new("excel_disconnect", "현재 Excel COM 연결을 즉시 해제합니다. 사용자 Excel에는 Quit을 호출하지 않고, DocBridge 소유 인스턴스만 저장되지 않은 변경이 없을 때 종료합니다.",
                NoInput("입력 없음"),
                _ => host.CoreDisconnect(new JsonObject { ["app"] = "excel" })),

            // ---------- hwp (§6.3) ----------
            new("hwp_plan_creation", "새 한글 문서의 제작 경로를 결정합니다. 새 일반 문서는 DOCX 우선, 기존 HWP/HWPX·한글 필드·복잡한 병합표·한글 전용 개체·원본 배치 보존은 HWP 직접 편집으로 고정합니다.",
                new JsonObject
                {
                    ["type"] = "object",
                    ["description"] = "파일을 만들거나 앱을 실행하지 않는 읽기 전용 정책 평가입니다. 새 문서 제작 전에 호출하고 반환된 mode와 workflow를 따릅니다.",
                    ["properties"] = new JsonObject
                    {
                        ["documentState"] = new JsonObject
                        {
                            ["type"] = "string",
                            ["enum"] = new JsonArray("new", "existing-hwp", "existing-hwpx"),
                            ["description"] = "새 문서는 new, 기존 한글 파일 편집은 확장자에 맞는 existing-hwp/existing-hwpx.",
                        },
                        ["hasExistingHwpTemplate"] = new JsonObject { ["type"] = "boolean" },
                        ["requiresNativeFields"] = new JsonObject { ["type"] = "boolean", ["description"] = "한글 필드·누름틀 등 네이티브 필드 필요 여부." },
                        ["requiresHwpOnlyObjects"] = new JsonObject { ["type"] = "boolean", ["description"] = "한글 전용 개체 필요 여부." },
                        ["requiresComplexMergedTables"] = new JsonObject { ["type"] = "boolean", ["description"] = "복잡한 병합표를 네이티브 구조로 유지할지 여부." },
                        ["mustPreserveOriginalLayout"] = new JsonObject { ["type"] = "boolean", ["description"] = "기존 HWP 원본 배치를 그대로 보존할지 여부." },
                        ["docxGeneratorAvailable"] = new JsonObject { ["type"] = "boolean", ["description"] = "검수 가능한 DOCX 생성 도구가 있으면 true(기본)." },
                    },
                    ["required"] = new JsonArray("documentState"),
                },
                a => host.HwpPlanCreation(a)),

            new("hwp_launch", "hwp_plan_creation 결정에 따라 네이티브 빈 문서를 시작하거나, 검수한 DOCX를 OOXML로 가져와 새 HWPX/HWP로 안전 변환합니다. 기존 출력은 덮어쓰지 않습니다.",
                new JsonObject
                {
                    ["type"] = "object",
                    ["description"] = "새 일반 문서는 creationMode=docx-first와 sourceFile/outputFile을 사용합니다. newDocument=true는 hwp_plan_creation이 native-hwp를 반환한 한글 전용 양식에서만 한 번 사용합니다. DOCX 경로는 한글 FileOpen(OOXML) 후 FileSaveAs(HWPX/HWP)로 변환하며 원본 해시 불변과 출력 검증 정보를 반환합니다.",
                    ["properties"] = new JsonObject
                    {
                        ["creationMode"] = new JsonObject
                        {
                            ["type"] = "string",
                            ["enum"] = new JsonArray("docx-first", "native-hwp"),
                            ["description"] = "hwp_plan_creation이 반환한 mode. 기존 호출 호환을 위해 생략 가능하지만 새 제작 작업에서는 명시합니다.",
                        },
                        ["newDocument"] = new JsonObject
                        {
                            ["type"] = "boolean",
                            ["description"] = "creationMode=native-hwp인 새 한글 전용 양식에서만 true. 새 일반 문서는 사용하지 않습니다.",
                        },
                        ["sourceFile"] = new JsonObject
                        {
                            ["type"] = "string",
                            ["description"] = "Absolute .docx path. Imports through HWP FileOpen with OOXML format.",
                        },
                        ["outputFile"] = new JsonObject
                        {
                            ["type"] = "string",
                            ["description"] = "Absolute new .hwpx or .hwp path. Defaults to the source name with .hwpx; existing files are refused.",
                        },
                        ["closeAfterImport"] = new JsonObject
                        {
                            ["type"] = "boolean",
                            ["description"] = "Batch verification mode. Close the imported tab after conversion. Default false keeps the HWP result open.",
                        },
                        ["expectedPageCount"] = new JsonObject
                        {
                            ["type"] = "integer",
                            ["minimum"] = 1,
                            ["description"] = "DOCX render page count. A mismatch is returned as verification.passed=false and must not be reported complete.",
                        },
                        ["expectedTableCount"] = new JsonObject
                        {
                            ["type"] = "integer",
                            ["minimum"] = 0,
                            ["description"] = "Expected native HWP table control count after import.",
                        },
                        ["requiredText"] = new JsonObject
                        {
                            ["type"] = "array",
                            ["items"] = new JsonObject { ["type"] = "string" },
                            ["description"] = "Critical phrases that must survive OOXML import exactly.",
                        },
                    },
                },
                a => host.HwpLaunch(a)),

            new("hwp_get_active_context", "모든 표시 한글 창과 탭을 summary.openDocuments로 열거하고, 연결된 활성 문서의 경로·선택 영역·텍스트 요약 조회",
                NoInput("입력 없음"),
                _ => host.GetActiveContext("hwp")),

            new("hwp_doctor", "한글 실행 없이 ProgID, 설치 버전, TypeLib GUID/등록 경로와 버전 불일치를 진단",
                NoInput("입력 없음. 상태가 CHECK_PASSED가 아니면 userAction을 확인하세요."),
                a => host.HwpDoctor(a)),

            new("hwp_repair_typelib", "[고위험] 설치된 Hwp.exe /RegServer를 실행해 잘못되거나 누락된 TypeLib 등록을 복구. UAC 승인과 한글/AI 클라이언트 재시작이 필요합니다.",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["confirm"] = new JsonObject { ["type"] = "boolean", ["description"] = "사용자가 TypeLib 재등록을 명시 승인했을 때만 true" },
                        ["hwpExecutable"] = new JsonObject { ["type"] = "string", ["description"] = "선택 사항: 사용할 Hwp.exe 절대 경로" },
                        ["elevate"] = new JsonObject { ["type"] = "boolean", ["description"] = "기본 true. 관리자 UAC로 실행" },
                    },
                    ["required"] = new JsonArray("confirm"),
                },
                a => host.HwpRepairTypeLib(a)),

            new("hwp_read_text", "한글 문서 텍스트와 표 셀/서식을 비파괴로 읽기. documentRef로 열린 창/탭을 선택하고, file로 특정 HWP/HWPX 파일을 지정할 수 있다.",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["file"] = new JsonObject { ["type"] = "string", ["description"] = "선택 사항: 읽을 .hwp/.hwpx 파일의 절대 경로. 생략하면 기존 활성 한글 창" },
                        ["documentRef"] = new JsonObject { ["type"] = "string", ["description"] = "선택 사항: hwp_get_active_context.summary.openDocuments의 documentRef 또는 instanceRef. file과 함께 쓰지 않음" },
                        ["scope"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("selection", "document", "bundle", "document_map", "structure", "fields", "tables") },
                        ["sections"] = new JsonObject
                        {
                            ["type"] = "array",
                            ["items"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("text", "document_map", "structure", "fields", "tables") },
                            ["description"] = "scope=bundle에서 한 COM 연결로 함께 읽을 항목. 기본 text+document_map+structure",
                        },
                        ["maxChars"] = new JsonObject { ["type"] = "integer" },
                        ["startParagraph"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0, ["description"] = "scope=document_map의 시작 문단(0부터)" },
                        ["maxParagraphs"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1, ["maximum"] = 2000 },
                        ["maxControls"] = new JsonObject { ["type"] = "integer" },
                        ["includePageCount"] = new JsonObject { ["type"] = "boolean" },
                        ["maxFields"] = new JsonObject { ["type"] = "integer" },
                        ["includeValues"] = new JsonObject { ["type"] = "boolean" },
                        ["tableIndex"] = new JsonObject { ["type"] = "integer", ["description"] = "scope=tables일 때 선택적 표 번호(0부터). 생략하면 모든 표" },
                        ["maxCells"] = new JsonObject { ["type"] = "integer", ["description"] = "scope=tables의 표당 최대 셀 수" },
                        ["includeStyles"] = new JsonObject { ["type"] = "boolean", ["description"] = "scope=tables에서 글자/문단 서식 메타데이터 포함(기본 true)" },
                    },
                    ["required"] = new JsonArray("scope"),
                },
                a => host.Read("hwp", a)),

            new("hwp_apply_ops", "한글 쓰기 ops 적용. 기존 값과 위·아래 문단 및 반복 양식의 같은 역할을 비교해 글자·문단 서식을 기본 보존하며 PowerShell COM 우회를 사용하지 않습니다.",
                HwpApplyOpsSchema(),
                a => host.ApplyOps("hwp", a)),

            new("hwp_submit_ops", "긴 한글 쓰기 배치를 비동기로 제출하고 즉시 jobId를 반환합니다. hwp_get_job으로 원래 적용 결과를 조회해 60초 client timeout 뒤의 중복 적용을 방지합니다.",
                HwpApplyOpsSchema(),
                a => hwpJobs.Submit(a)),

            new("hwp_get_job", "hwp_submit_ops가 반환한 작업의 queued/running/succeeded/failed 상태와 최종 결과를 조회합니다.",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject { ["jobId"] = new JsonObject { ["type"] = "string" } },
                    ["required"] = new JsonArray("jobId"),
                },
                a => hwpJobs.Get(a)),

            // ---------- cad (§6.4) ----------
            new("cad_launch", "AutoCAD를 내부 COM 명령으로 실행하거나 기존 인스턴스에 연결하고, 창을 표시한 뒤 편집 가능한 활성 도면을 보장",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["template"] = new JsonObject
                        {
                            ["type"] = "string",
                            ["enum"] = new JsonArray("acad.dwt", "acadiso.dwt"),
                            ["description"] = "활성 도면이 없을 때 사용할 기본 템플릿",
                        },
                    },
                },
                a => host.CadLaunch(a)),

            new("cad_get_active_context", "열린 도면과 활성 도면의 경량 상태 조회. 기본 basic은 대형 도면의 엔티티를 순회하지 않으며, summary만 레이어 미리보기와 최대 500개 엔티티 유형 표본을 읽음",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["detailLevel"] = new JsonObject
                        {
                            ["type"] = "string",
                            ["enum"] = new JsonArray("basic", "summary"),
                            ["default"] = "basic",
                            ["description"] = "basic: 도면/개수/단위만 읽고 엔티티·레이어를 순회하지 않음. summary: 레이어 최대 50개와 엔티티 최대 500개 유형 표본. 전체 자료는 응답의 nextActions에 따라 cad_query_entities 사용",
                        },
                    },
                },
                a => host.GetActiveContext("cad", a)),

            new("cad_query_entities", "도면 엔티티·배치/뷰포트·레이어·XREF 조회 또는 여러 도곽 영역 일괄 검증 (AutoCAD 미실행 시 file 인자로 DXF 분석)",
                new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = new JsonObject
                    {
                        ["scope"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("entities", "layouts", "layers", "xrefs", "window", "regions"), ["description"] = "생략하면 entities. layers는 전체 레이어, xrefs는 XREF 상태, window는 AutoCAD 네이티브 공간선택, regions는 ModelSpace 한 번 순회로 여러 도곽을 검증" },
                        ["contains"] = new JsonObject { ["type"] = "string", ["description"] = "scope=layers 이름 포함 필터" },
                        ["startsWith"] = new JsonObject { ["type"] = "string", ["description"] = "scope=layers 이름 접두사 필터" },
                        ["regions"] = new JsonObject { ["type"] = "array", ["description"] = "scope=regions: name, bounds, 선택적 entityTypes/layer/textContains/minCount/maxCount/boundsMode 배열" },
                        ["layer"] = new JsonObject { ["type"] = "string" },
                        ["entityType"] = new JsonObject { ["type"] = "string" },
                        ["document"] = new JsonObject { ["type"] = "string", ["description"] = "열린 DWG 이름/경로 패턴(*, ? 지원)" },
                        ["textContains"] = new JsonObject { ["type"] = "string" },
                        ["blockName"] = new JsonObject { ["type"] = "string" },
                        ["includeGeometry"] = new JsonObject { ["type"] = "boolean" },
                        ["countOnly"] = new JsonObject { ["type"] = "boolean", ["description"] = "엔티티 목록 없이 전체 일치 개수와 합산 경계만 반환" },
                        ["startIndex"] = new JsonObject { ["type"] = "integer", ["description"] = "ModelSpace 시작 인덱스(증분 검증용)" },
                        ["endIndex"] = new JsonObject { ["type"] = "integer", ["description"] = "ModelSpace 종료 인덱스" },
                        ["boundsMode"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("center", "inside", "intersect") },
                        ["bounds"] = new JsonObject
                        {
                            ["type"] = "object",
                            ["properties"] = new JsonObject
                            {
                                ["minX"] = new JsonObject { ["type"] = "number" },
                                ["minY"] = new JsonObject { ["type"] = "number" },
                                ["maxX"] = new JsonObject { ["type"] = "number" },
                                ["maxY"] = new JsonObject { ["type"] = "number" },
                            },
                        },
                        ["limit"] = new JsonObject { ["type"] = "integer" },
                        ["file"] = new JsonObject { ["type"] = "string", ["description"] = "DWG/DXF 파일 경로. DWG는 AutoCAD에서 읽기 전용으로 열어 조회" },
                    },
                },
                a => host.Read("cad", a)),

            new("cad_apply_ops", "CAD 쓰기 ops 적용. 도형·해치·수정·블록속성·배치/뷰포트·저장/출력을 ActiveX COM으로 직접 처리. XREF 클립만 AutoCAD 기본 XCLIP 명령을 사용하며 AutoLISP는 사용하지 않음.",
                ApplyOpsSchema("cad", "activate_document, set_layer_visibility/color, move/rotate/set_text, copy_entities_between_documents, insert_xref, zoom_window, draw_entities(lwpolyline/circle/block/text/hatch/line/arc/ellipse/point/mtext/dim_aligned/dim_rotated), copy/scale/mirror/offset_entities, set_entity_properties, set_block_attributes, configure_layout, create_viewport / 고위험: save_document, plot_pdf, delete_entities*, run_script_template"),
                a => host.ApplyOps("cad", a)),
        };
    }

    public IReadOnlyList<ToolDef> All => _tools;

    public ToolDef? Find(string name) => _tools.FirstOrDefault(t => t.Name == name);

    /// <summary>tools/list 응답용 명세 배열</summary>
    public JsonArray ListSpec()
    {
        var arr = new JsonArray();
        foreach (var t in _tools)
        {
            var readOnly = t.Name is "core_ping" or "core_get_status" or "core_get_capabilities" or "core_list_snapshots"
                or "excel_get_active_context" or "excel_read_range" or "excel_inspect"
                or "hwp_plan_creation" or "hwp_get_active_context" or "hwp_read_text" or "hwp_doctor" or "hwp_get_job"
                or "cad_get_active_context" or "cad_query_entities";
            var destructive = t.Name.EndsWith("_apply_ops", StringComparison.Ordinal)
                              || t.Name is "core_restore_snapshot" or "hwp_repair_typelib" or "hwp_submit_ops";
            arr.Add(new JsonObject
            {
                ["name"] = t.Name,
                ["description"] = t.Description,
                ["inputSchema"] = t.InputSchema.DeepClone(),
                // MCP 2025-06-18 tool hints. 클라이언트 승인을 대체하지 않으며 UI 분류에만 사용된다.
                ["annotations"] = new JsonObject
                {
                    ["readOnlyHint"] = readOnly,
                    ["destructiveHint"] = destructive,
                    ["idempotentHint"] = readOnly || t.Name is "core_disconnect" or "excel_disconnect",
                    ["openWorldHint"] = false,
                },
            });
        }
        return arr;
    }
}
