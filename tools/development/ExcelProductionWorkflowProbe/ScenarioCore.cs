namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class ScenarioCore
{
    public static List<PlannedBatch> E2(string artifactDir, string sourceXlsx)
    {
        const string sheet = "내역서";
        var xlsx = Path.Combine(artifactDir, "e2-cost-estimate.xlsx");
        var pdf = Path.Combine(artifactDir, "e2-cost-estimate.pdf");
        SourceIntegrity.RefuseIfProtectedPath(xlsx, sourceXlsx);

        var headers = new[] { "번호", "공종", "규격", "단위", "수량", "단가", "금액", "비고" };
        var values = new JsonArray
        {
            Arr("공사비 수량 산출 내역서 (기능검증용 가상 데이터)", null, null, null, null, null, null, null),
            Arr("현장", "수원시 하수관로 정비 공사 (가상 수량)", null, null, "기준일", "2026-09-10", null, null),
            Arr(null, null, null, null, null, null, null, null),
            Arr(headers.Cast<object?>().ToArray()),
        };
        var row = 5;
        var items = new List<int>();
        foreach (var (i, line) in AcceptanceNumbers.CostLines.Select((l, i) => (i + 1, l)))
        {
            values.Add(Arr(i, line.Name, "본/개소", "식", line.Qty, line.Unit, null, ""));
            items.Add(row);
            row++;
        }
        var subRow = row;
        values.Add(Arr(null, "소계", null, null, null, null, null, null));
        row++;
        var vatRow = row;
        values.Add(Arr(null, "부가세", null, null, null, null, null, "ROUND(소계*0.1,0)"));
        row++;
        var grandRow = row;
        values.Add(Arr(null, "합계", null, null, null, null, null, null));
        values.Add(Arr("작성", null, "검토", null, "승인", null, "결재", null));
        values.Add(Arr(null, null, null, null, null, null, null, null));
        values.Add(Arr(null, null, null, null, null, null, null, null));
        while (values.Count < 23)
            values.Add(Arr(null, null, null, null, null, null, null, null));
        values.Add(Arr("DIRTY", "오입력 999는 합계 제외", null, null, 999, 999, null, "computed G는 G5:G8만"));

        var formulas = new JsonArray();
        foreach (var r in items)
            formulas.Add(new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = $"G{r}", ["formulas"] = Arr(Arr($"=ROUND(E{r}*F{r},0)")) });
        var itemRefs = string.Join(",", items.Select(r => $"G{r}"));
        formulas.Add(new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = $"G{subRow}", ["formulas"] = Arr(Arr($"=SUM({itemRefs})")) });
        formulas.Add(new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = $"G{vatRow}", ["formulas"] = Arr(Arr($"=ROUND(G{subRow}*0.1,0)")) });
        formulas.Add(new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = $"G{grandRow}", ["formulas"] = Arr(Arr($"=G{subRow}+G{vatRow}")) });

        return
        [
            .. Start("e2", sheet, xlsx),
            Values("e2-values", sheet, "A1:H24", values),
            ApplyPlanner.Batch("e2", "e2-formulas", OpFamily.Values, formulas, "execute",
                "initial G5/G9/G10/G11 before qty mutation", capture: "initial"),
            ApplyPlanner.Batch("e2", "e2-merge-title", OpFamily.Merge, ApplyPlanner.Merges(sheet, ["A1:H1", "B2:D2"]), "token"),
            ApplyPlanner.Batch("e2", "e2-recalc-qty", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "E5",
                ["values"] = Arr(Arr(AcceptanceNumbers.CostQtyChanged)),
            }), "execute", "관로 수량 130 → 금액/소계/부가세/합계 갱신"),
            ApplyPlanner.Batch("e2", "e2-copy-scratch", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "copy_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A5:H5",
                ["destRange"] = "A20",
                ["mode"] = "formulas",
            }), "token", "scratch copy; live then delete_rows the scratch strip"),
            .. E2DocumentQa(sheet, xlsx),
            .. SaveCloseReopen("e2", xlsx, pdf, sheet),
        ];
    }

    public static List<PlannedBatch> E3(string artifactDir, string sourceXlsx, string pictureDir)
    {
        const string sheet = "작업일보";
        var xlsx = Path.Combine(artifactDir, "e3-daily-report.xlsx");
        var pdf = Path.Combine(artifactDir, "e3-daily-report.pdf");
        var pic1 = TinyPng.WriteSample(pictureDir, "e3-photo-1.png", 40, 90, 160);
        var pic2 = TinyPng.WriteSample(pictureDir, "e3-photo-2.png", 160, 90, 40);
        File.WriteAllText(Path.Combine(pictureDir, "README.txt"), "샘플 이미지. 현장 사진이 아님.\n");

        var longNote = "관로 터파기 및 부설 구간에서 기존 지장물 이설, 우수 유입 대비 양수, 야간 교통 통제 협의, 품질시험 의뢰, 인접 민원 대응을 포함한 당일 작업 내용을 한 셀에 기록하여 줄바꿈·행 높이·병합 가독성을 확인한다. " +
                       string.Join(" ", Enumerable.Repeat("추가 작업 메모.", 40));

        var values = new JsonArray
        {
            Arr("작업일보 (기능검증용 가상 데이터)", null, null, null, null, null),
            Arr("현장", "수원시 하수관로 정비", "작성일", "2026-09-10", "날씨", "맑음"),
            Arr("담당", "김현장", "작업구간", "고색처리분구 STA.0+120~0+180", "공종", "오수관로 신설"),
            Arr("구분", "인원", "장비", "작업량", "단위", "비고"),
            Arr("기능공", 6, 2, 18, "m", ""),
            Arr("보통인부", 4, 1, null, "", ""),
            Arr("관리", 3, null, null, "", ""),
            Arr("합계", null, null, null, "", "수식"),
            Arr("작업내용", longNote, null, null, null, null),
            Arr("특이사항", "샘플 사진 2장. 실제 현장이 아님.", null, null, null, null),
            Arr("작성", "", "검토", "", "현장대리인", ""),
            Arr("감리", "", "", "", "", ""),
        };

        return
        [
            .. Start("e3", sheet, xlsx),
            Values("e3-values", sheet, "A1:F12", values),
            ApplyPlanner.Batch("e3", "e3-formulas", OpFamily.Values, JsonUtil.Arr(
                new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = "B8", ["formulas"] = Arr(Arr("=B5+B6+B7")) },
                new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = "C8", ["formulas"] = Arr(Arr("=C5+C6")) }), "execute"),
            ApplyPlanner.Batch("e3", "e3-merge", OpFamily.Merge, ApplyPlanner.Merges(sheet, ["A1:F1", "B9:F9", "B10:F10", "A11:B11", "C11:D11", "E11:F11"]), "token"),
            ApplyPlanner.Batch("e3", "e3-pictures", OpFamily.Data, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "insert_sheet_picture",
                    ["target"] = JsonUtil.Target(sheet),
                    ["path"] = pic1,
                    ["name"] = "Photo1",
                    ["position"] = new JsonObject { ["left"] = 20, ["top"] = 220, ["width"] = 180, ["height"] = 120 },
                },
                new JsonObject
                {
                    ["op"] = "insert_sheet_picture",
                    ["target"] = JsonUtil.Target(sheet),
                    ["path"] = pic2,
                    ["name"] = "Photo2",
                    ["position"] = new JsonObject { ["left"] = 220, ["top"] = 220, ["width"] = 180, ["height"] = 120 },
                }), "token", "샘플 이미지"),
            ApplyPlanner.Batch("e3", "e3-layout", OpFamily.SheetLayout, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_row_heights",
                    ["target"] = JsonUtil.Target(sheet),
                    ["rows"] = JsonUtil.Arr(
                        new JsonObject { ["row"] = 1, ["heightPoints"] = 28 },
                        new JsonObject { ["row"] = 9, ["heightPoints"] = 90 },
                        new JsonObject { ["row"] = 11, ["heightPoints"] = 36 }),
                },
                new JsonObject
                {
                    ["op"] = "set_page_setup",
                    ["target"] = JsonUtil.Target(sheet),
                    ["page"] = new JsonObject
                    {
                        ["paperSize"] = "A4",
                        ["orientation"] = "portrait",
                        ["fitToWidth"] = 1,
                        ["fitToHeight"] = 1,
                        ["printArea"] = "A1:F12",
                    },
                }), "token"),
            ApplyPlanner.Batch("e3", "e3-wrap", OpFamily.Format, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "B9",
                ["style"] = new JsonObject { ["wrapText"] = true, ["verticalAlign"] = "top" },
            }), "execute"),
            .. RemainingOpWorkflows.E3(sheet, pictureDir),
            .. SaveCloseReopen("e3", xlsx, pdf, sheet),
        ];
    }

    public static List<PlannedBatch> E4(string artifactDir, string sourceXlsx)
    {
        const string sheet = "자재대장";
        var xlsx = Path.Combine(artifactDir, "e4-materials-ledger.xlsx");
        var pdf = Path.Combine(artifactDir, "e4-materials-ledger.pdf");
        var values = new JsonArray
        {
            Arr("자재 입출고 대장 (기능검증용 가상 데이터)", null, null, null, null, null, null, null, null),
            Arr("거래ID", "일자", "품목", "규격", "입고", "출고", "잔량", "담당", "비고"),
            Arr("TX-001", "2026-09-01", "PVC관", "D300", 120, 40, null, "이민수", "초도 입고"),
            Arr("TX-002", "2026-09-03", "맨홀", "1호", 10, 3, null, "박하린", ""),
            Arr("TX-003", "2026-09-05", "PVC관", "D300", 50, 65, null, "이민수", "야간"),
        };

        return
        [
            .. Start("e4", sheet, xlsx),
            Values("e4-values", sheet, "A1:I5", values),
            ApplyPlanner.Batch("e4", "e4-remain-formulas", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_formulas",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G3:G5",
                ["formulas"] = Arr(
                    Arr("=E3-F3"),
                    Arr("=E4-F4"),
                    Arr("=E5-F5")),
            }), "execute", "initial G3/G4/G5 remains before add-row", capture: "initial"),
            ApplyPlanner.Batch("e4", "e4-table", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "create_table",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A2:I5",
                ["name"] = "자재표",
                ["hasHeaders"] = true,
                ["styleName"] = "TableStyleMedium2",
                ["showTotals"] = true,
            }), "token"),
            ApplyPlanner.Batch("e4", "e4-validation", OpFamily.Data, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_data_validation",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "H3:H20",
                    ["type"] = "list",
                    ["formula1"] = "이민수,박하린,최감독",
                    ["inCellDropdown"] = true,
                },
                new JsonObject
                {
                    ["op"] = "set_data_validation",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "E3:F20",
                    ["type"] = "decimal",
                    ["operator"] = "greaterEqual",
                    ["formula1"] = "0",
                },
                new JsonObject
                {
                    ["op"] = "set_data_validation",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "B3:B20",
                    ["type"] = "date",
                    ["operator"] = "between",
                    ["formula1"] = "2026-01-01",
                    ["formula2"] = "2026-12-31",
                }), "token"),
            ApplyPlanner.Batch("e4", "e4-cf", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "add_conditional_format",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A3:I20",
                ["rule"] = new JsonObject { ["type"] = "expression", ["formula1"] = "=$F3>$E3" },
                ["style"] = new JsonObject { ["fillColor"] = "#F4B183" },
            }, new JsonObject
            {
                ["op"] = "add_conditional_format",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G3:G20",
                ["rule"] = new JsonObject { ["type"] = "cellValue", ["operator"] = "less", ["formula1"] = "10" },
                ["style"] = new JsonObject { ["fillColor"] = "#FF6B6B" },
            }), "token", "출고>입고 또는 잔량<안전재고(10)"),
            ApplyPlanner.Batch("e4", "e4-sort-table", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "sort_table",
                ["target"] = JsonUtil.Target(sheet),
                ["name"] = "자재표",
                ["keys"] = JsonUtil.Arr(
                    new JsonObject { ["column"] = "일자", ["order"] = "asc" },
                    new JsonObject { ["column"] = "거래ID", ["order"] = "asc" }),
            }), "token", "two keys; PVC관 is a duplicate 품목"),
            ApplyPlanner.Batch("e4", "e4-filter", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_auto_filter",
                ["target"] = JsonUtil.Target(sheet),
                ["name"] = "자재표",
                ["criteria"] = JsonUtil.Arr(new JsonObject { ["column"] = "품목", ["operator"] = "equals", ["value"] = "PVC관" }),
            }), "token", "0918 schema operator is equals, not equal"),
            ApplyPlanner.Batch("e4", "e4-clear-filter", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "clear_auto_filter",
                ["target"] = JsonUtil.Target(sheet),
                ["name"] = "자재표",
            }), "token", "filtered PVC intermediate is checkpointed before clear"),
            ApplyPlanner.Batch("e4", "e4-dirty-void", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A12:I12",
                ["values"] = Arr(Arr("VOID-999", "2026-09-09", "PVC관", "D300", 999, 0, null, "취소", "더티 행: 잔량 수식/표에 넣지 않음")),
            }), "execute", "cancelled 999 must not enter G3/G4/G6 remain totals"),
            ApplyPlanner.Batch("e4", "e4-add-row", OpFamily.Values, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_values",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A6:I6",
                    ["values"] = Arr(Arr("TX-004", "2026-09-08", "PVC관", "D300", 20, 5, null, "박하린", "추가입고")),
                },
                new JsonObject
                {
                    ["op"] = "set_formulas",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "G6",
                    ["formulas"] = Arr(Arr("=E6-F6")),
                }), "execute"),
            ApplyPlanner.Batch("e4", "e4-resize", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "resize_table",
                ["target"] = JsonUtil.Target(sheet),
                ["name"] = "자재표",
                ["range"] = "A2:I6",
            }), "token", "TX-004 is inside 자재표; VOID-999 stays out"),
            ApplyPlanner.Batch("e4", "e4-safety-stock-name", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "define_name",
                ["target"] = JsonUtil.Target(sheet),
                ["name"] = "안전재고",
                ["refersTo"] = "=10",
                ["scope"] = "workbook",
            }), "token"),
            .. RemainingOpWorkflows.E4(sheet, xlsx),
            .. SaveCloseReopen("e4", xlsx, pdf, sheet),
        ];
    }

    public static List<PlannedBatch> E5(string artifactDir, string sourceXlsx)
    {
        var xlsx = Path.Combine(artifactDir, "e5-progress-payment.xlsx");
        var pdf = Path.Combine(artifactDir, "e5-progress-payment.pdf");
        const string input = "기성입력";
        const string summary = "누계집계";

        return
        [
            .. Start("e5", input, xlsx),
            ApplyPlanner.Batch("e5", "e5-add-summary", OpFamily.Structure, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "add_sheet",
                ["name"] = summary,
                ["afterSheet"] = input,
            }), "token"),
            Values("e5-input", input, "A1:B5", new JsonArray
            {
                Arr("기성 입력 (기능검증용 가상 데이터)", null),
                Arr("계약금액", AcceptanceNumbers.ContractAmount),
                Arr("1월", AcceptanceNumbers.Jan),
                Arr("2월", AcceptanceNumbers.Feb),
                Arr("3월", AcceptanceNumbers.Mar),
            }),
            Values("e5-summary-labels", summary, "A1:B6", new JsonArray
            {
                Arr("기성 누계 집계", null),
                Arr("누계", null),
                Arr("누계율", null),
                Arr("잔액", null),
                Arr("입력시트", null),
                Arr("비고", "소계를 다시 더하지 않음. 더티 999000000은 이름/수식에 없음"),
            }),
            ApplyPlanner.Batch("e5", "e5-names", OpFamily.Data, JsonUtil.Arr(
                new JsonObject { ["op"] = "define_name", ["name"] = "ContractAmount", ["refersTo"] = $"='{input}'!$B$2", ["scope"] = "workbook" },
                new JsonObject { ["op"] = "define_name", ["name"] = "JanAmt", ["refersTo"] = $"='{input}'!$B$3", ["scope"] = "workbook" },
                new JsonObject { ["op"] = "define_name", ["name"] = "FebAmt", ["refersTo"] = $"='{input}'!$B$4", ["scope"] = "workbook" },
                new JsonObject { ["op"] = "define_name", ["name"] = "MarAmt", ["refersTo"] = $"='{input}'!$B$5", ["scope"] = "workbook" }), "token",
                "quoted absolute $B$2..$B$5 for fixed named cells; do not treat this as a relative-name test"),
            ApplyPlanner.Batch("e5", "e5-formulas", OpFamily.Values, JsonUtil.Arr(
                new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(summary), ["range"] = "B2", ["formulas"] = Arr(Arr("=JanAmt+FebAmt+MarAmt")) },
                new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(summary), ["range"] = "B3", ["formulas"] = Arr(Arr("=B2/ContractAmount")) },
                new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(summary), ["range"] = "B4", ["formulas"] = Arr(Arr("=ContractAmount-B2")) }),
                "execute", "initial B2/B3/B4 before Feb mutation"),
            ApplyPlanner.Batch("e5", "e5-calculate", OpFamily.Values, JsonUtil.Arr(
                new JsonObject { ["op"] = "calculate" }), "token",
                "named formulas need a public calculate before initial capture", capture: "initial"),
            ApplyPlanner.Batch("e5", "e5-widths", OpFamily.SheetLayout, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_column_widths",
                    ["target"] = JsonUtil.Target(input),
                    ["columns"] = JsonUtil.Arr(
                        new JsonObject { ["col"] = "A", ["count"] = 1, ["widthChars"] = 14 },
                        new JsonObject { ["col"] = "B", ["count"] = 1, ["widthChars"] = 16 }),
                },
                new JsonObject
                {
                    ["op"] = "set_column_widths",
                    ["target"] = JsonUtil.Target(summary),
                    ["columns"] = JsonUtil.Arr(
                        new JsonObject { ["col"] = "A", ["count"] = 1, ["widthChars"] = 14 },
                        new JsonObject { ["col"] = "B", ["count"] = 1, ["widthChars"] = 16 }),
                }), "token"),
            ApplyPlanner.Batch("e5", "e5-formats", OpFamily.Format, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(input),
                    ["range"] = "A1",
                    ["style"] = new JsonObject { ["fontName"] = "돋움", ["fontSize"] = 14, ["bold"] = true },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(input),
                    ["range"] = "B2:B5",
                    ["style"] = new JsonObject { ["numberFormat"] = "#,##0", ["fontName"] = "돋움", ["fontSize"] = 10 },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(summary),
                    ["range"] = "B2",
                    ["style"] = new JsonObject { ["numberFormat"] = "#,##0", ["fontName"] = "돋움", ["fontSize"] = 10 },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(summary),
                    ["range"] = "B3",
                    ["style"] = new JsonObject { ["numberFormat"] = "0%", ["fontName"] = "돋움", ["fontSize"] = 10 },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(summary),
                    ["range"] = "B4",
                    ["style"] = new JsonObject { ["numberFormat"] = "#,##0", ["fontName"] = "돋움", ["fontSize"] = 10 },
                }), "execute", "readable currency / percent formats, not values-only"),
            ApplyPlanner.Batch("e5", "e5-hyperlink", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_hyperlink",
                ["target"] = JsonUtil.Target(summary),
                ["range"] = "B5",
                ["subAddress"] = $"'{input}'!A1",
                ["textToDisplay"] = "입력으로 이동",
            }), "token"),
            ApplyPlanner.Batch("e5", "e5-rename-move", OpFamily.Rename, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "rename_sheet",
                ["target"] = JsonUtil.Target(input),
                ["newName"] = "월별기성",
            }), "token"),
            ApplyPlanner.Batch("e5", "e5-hyperlink-after-rename", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_hyperlink",
                ["target"] = JsonUtil.Target(summary),
                ["range"] = "B5",
                ["subAddress"] = "'월별기성'!A1",
                ["textToDisplay"] = "입력으로 이동",
            }), "token", "public retarget after rename; leftover 기성입력!A1 is a failed requirement"),
            ApplyPlanner.Batch("e5", "e5-recalc-feb", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target("월별기성"),
                ["range"] = "B4",
                ["values"] = Arr(Arr(AcceptanceNumbers.FebChanged)),
            }), "execute"),
            ApplyPlanner.Batch("e5", "e5-calculate-after-feb", OpFamily.Values, JsonUtil.Arr(
                new JsonObject { ["op"] = "calculate" }), "token",
                "public calculate after Feb 20M so 47M/47%/53M are live values"),
            .. RemainingOpWorkflows.E5(summary, "월별기성", xlsx),
            .. SaveCloseReopen("e5", xlsx, pdf, summary),
        ];
    }

    public static List<PlannedBatch> E6(string artifactDir, string sourceXlsx, string pictureDir)
    {
        const string sheet = "월간보고";
        var xlsx = Path.Combine(artifactDir, "e6-monthly-report.xlsx");
        var pdf = Path.Combine(artifactDir, "e6-monthly-report.pdf");
        var logo = TinyPng.WriteSample(pictureDir, "e6-logo.png", 20, 20, 20);
        var months = new[] { "1월", "2월", "3월", "4월", "5월", "6월" };
        var header = new JsonArray { Arr("월간 공정 보고서 (기능검증용 가상 데이터)", null, null, null) };
        header.Add(Arr("월", "계획누계율", "실적누계율", "차이"));
        for (var i = 0; i < 6; i++)
            header.Add(Arr(months[i], AcceptanceNumbers.PlanRates[i], AcceptanceNumbers.ActualRates[i], null));
        for (var i = 1; i <= 30; i++)
            header.Add(Arr($"작업항목 {i:00}", null, null, "다쪽 인쇄용 행"));
        for (var i = 39; i <= 70; i++)
            header.Add(Arr($"인쇄여백 {i:00}", null, null, "차트·2쪽 인쇄 영역"));

        var diffs = new JsonArray();
        for (var r = 3; r <= 8; r++)
            diffs.Add(new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = $"D{r}", ["formulas"] = Arr(Arr($"=C{r}-B{r}")) });

        return
        [
            .. Start("e6", sheet, xlsx),
            Values("e6-values", sheet, "A1:D70", header),
            ApplyPlanner.Batch("e6", "e6-diff", OpFamily.Values, diffs, "execute",
                "initial D5=3 before Mar mutation", capture: "initial"),
            ApplyPlanner.Batch("e6", "e6-diff-helper", OpFamily.Values, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_values",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "F2:G2",
                    ["values"] = Arr(Arr("월", "차이")),
                },
                new JsonObject
                {
                    ["op"] = "set_formulas",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "F3:G8",
                    ["formulas"] = new JsonArray
                    {
                        Arr("=A3", "=D3"),
                        Arr("=A4", "=D4"),
                        Arr("=A5", "=D5"),
                        Arr("=A6", "=D6"),
                        Arr("=A7", "=D7"),
                        Arr("=A8", "=D8"),
                    },
                }), "execute", "contiguous F2:G8 helper; union A2:A8,D2:D8 stays a data gap"),
            ApplyPlanner.Batch("e6", "e6-charts", OpFamily.Data, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "create_chart",
                    ["target"] = JsonUtil.Target(sheet),
                    ["name"] = "PlanActualLine",
                    ["chartType"] = "lineMarkers",
                    ["sourceRange"] = "A2:C8",
                    ["title"] = "계획/실적 누계율",
                    ["hasLegend"] = true,
                    ["legendPosition"] = "bottom",
                    ["position"] = new JsonObject { ["left"] = 10, ["top"] = 580, ["width"] = 360, ["height"] = 180 },
                    ["series"] = JsonUtil.Arr(
                        new JsonObject { ["values"] = "B2:B8", ["categories"] = "A2:A8", ["name"] = "계획누계율" },
                        new JsonObject { ["values"] = "C2:C8", ["categories"] = "A2:A8", ["name"] = "실적누계율" }),
                },
                new JsonObject
                {
                    ["op"] = "create_chart",
                    ["target"] = JsonUtil.Target(sheet),
                    ["name"] = "DiffColumn",
                    ["chartType"] = "columnClustered",
                    ["sourceRange"] = "A2:A8,D2:D8",
                    ["title"] = "월별 차이",
                    ["hasLegend"] = true,
                    ["position"] = new JsonObject { ["left"] = 10, ["top"] = 770, ["width"] = 360, ["height"] = 160 },
                    ["series"] = JsonUtil.Arr(
                        new JsonObject { ["values"] = "D2:D8", ["categories"] = "A2:A8", ["name"] = "수량" }),
                },
                new JsonObject
                {
                    ["op"] = "insert_sheet_picture",
                    ["target"] = JsonUtil.Target(sheet),
                    ["path"] = logo,
                    ["name"] = "Logo",
                    ["position"] = new JsonObject { ["left"] = 10, ["top"] = 4, ["width"] = 28, ["height"] = 28 },
                }), "token", "charts sit in printed A1:D70 below the A20 break; helper F:G stays outside print"),
            ApplyPlanner.Batch("e6", "e6-print", OpFamily.SheetLayout, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_column_widths",
                    ["target"] = JsonUtil.Target(sheet),
                    ["columns"] = JsonUtil.Arr(
                        new JsonObject { ["col"] = "A", ["count"] = 1, ["widthChars"] = 16 },
                        new JsonObject { ["col"] = "B", ["count"] = 3, ["widthChars"] = 12 }),
                },
                new JsonObject
                {
                    ["op"] = "set_row_heights",
                    ["target"] = JsonUtil.Target(sheet),
                    ["rows"] = JsonUtil.Arr(
                        new JsonObject { ["row"] = 1, ["autoFit"] = true },
                        new JsonObject { ["row"] = 9, ["count"] = 30, ["autoFit"] = true }),
                },
                new JsonObject
                {
                    ["op"] = "set_page_setup",
                    ["target"] = JsonUtil.Target(sheet),
                    ["page"] = new JsonObject
                    {
                        ["paperSize"] = "A4",
                        ["orientation"] = "portrait",
                        ["printArea"] = "A1:D70",
                        ["printTitleRows"] = "$1:$2",
                        ["centerHeader"] = "월간 공정 보고서",
                        ["centerFooter"] = "기능검증용 가상 데이터",
                    },
                }), "token", "0558 fixture uses canonical $1:$2; bare 1:2 remains a product titleRows normalization regression"),
            ApplyPlanner.Batch("e6", "e6-page-break", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_page_breaks",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A20",
            }), "token", "0558 COM cannot get Range.PageBreak (0x800A03EC); A1:D70 + title rows still force two content pages. Keep A20 as product regression."),
            ApplyPlanner.Batch("e6", "e6-recalc-mar", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "C5",
                ["values"] = Arr(Arr(AcceptanceNumbers.MarActualChanged)),
            }), "execute", "3월 실적 45 → 차이 5"),
            .. RemainingOpWorkflows.E6(sheet),
            .. SaveCloseReopen("e6", xlsx, pdf, sheet),
        ];
    }

    public static JsonObject ExpectedE2() => new()
    {
        ["initial"] = new JsonObject
        {
            ["G5"] = AcceptanceNumbers.CostLines[0].Amount,
            ["G9"] = AcceptanceNumbers.CostSubtotal,
            ["G10"] = AcceptanceNumbers.CostVat,
            ["G11"] = AcceptanceNumbers.CostGrand,
        },
        ["afterQty130"] = new JsonObject
        {
            ["G5"] = AcceptanceNumbers.CostAmountChanged,
            ["G9"] = AcceptanceNumbers.CostSubtotalChanged,
            ["G10"] = AcceptanceNumbers.CostVatChanged,
            ["G11"] = AcceptanceNumbers.CostGrandChanged,
        },
        ["pdfPagesExpected"] = 1,
        ["printArea"] = "A1:H14",
        ["dirtyRow"] = 24,
        ["formats"] = "merged title, metadata row, table borders, #,##0원 currency, signature blocks; DIRTY outside print area",
        ["typedTextFixture"] = "F2 ISO 2026-09-10 and empty-string blanks stay required; do not weaken to dotted date or null",
        ["mixedNumberFormatRow"] = "A26:D26 formats General|0.00|yyyy-mm-dd|quoted LIT then strings 2026-09-10|0012|1E3|=1+1; strict string + original format after reopen",
        ["heightRegression"] = "rows 15/16 request 28/22; report actual measured mapping",
        ["signatureBlocks"] = JsonUtil.Arr("A12:B14", "C12:D14", "E12:F14", "G12:H14"),
        ["approvalLabel"] = "G12 결재; H12 blank; do not leave a G-column gap",
        ["mergedBorderAcceptance"] = "pending next pin; 0732 is before final MergeArea correction",
        ["reopen"] = "save + close_workbook + open_workbook + export_pdf",
    };

    public static JsonObject ExpectedE3() => new()
    {
        ["people"] = AcceptanceNumbers.DailyPeople,
        ["equipment"] = AcceptanceNumbers.DailyEquipment,
        ["pictures"] = 2,
        ["picturesAreSamples"] = true,
        ["pdfPagesExpectedMin"] = 1,
        ["reopen"] = "save + close_workbook + open_workbook + export_pdf",
    };

    public static JsonObject ExpectedE4() => new()
    {
        ["pvcRemain"] = AcceptanceNumbers.PvcRemain,
        ["manholeRemain"] = AcceptanceNumbers.ManholeRemain,
        ["pvcRemainAfterAdd"] = AcceptanceNumbers.PvcRemainAfterAdd,
        ["initial"] = new JsonObject
        {
            ["G3"] = AcceptanceNumbers.PvcRemain,
            ["G4"] = AcceptanceNumbers.ManholeRemain,
            ["G5"] = -15,
        },
        ["afterAdd"] = new JsonObject
        {
            ["G3"] = AcceptanceNumbers.PvcRemain,
            ["G6"] = 15,
            ["pvcRemain"] = AcceptanceNumbers.PvcRemainAfterAdd,
        },
        ["table"] = "자재표",
        ["netIn"] = new JsonObject { ["J3"] = 80, ["J4"] = 7, ["J5"] = -15, ["J6"] = 15 },
        ["tableTotals"] = new JsonObject { ["입고"] = 200, ["출고"] = 113, ["순입고"] = 87 },
        ["revision"] = new JsonObject
        {
            ["sheet"] = "개정",
            ["afterInsert"] = new JsonObject { ["입고"] = 230, ["출고"] = 121 },
            ["afterDeleteTx003"] = new JsonObject { ["입고"] = 180, ["출고"] = 56 },
            ["ids"] = JsonUtil.Arr("TX-001", "TX-002", "TX-004", "TX-005"),
            ["removed"] = "TX-003",
            ["sortKeys"] = "품목 then 거래ID; PVC관 is a duplicate primary key",
        },
        ["pdfPagesExpected"] = 1,
        ["reopen"] = "save + close_workbook + open_workbook + export_pdf",
    };

    public static JsonObject ExpectedE5() => new()
    {
        ["initial"] = new JsonObject
        {
            ["B2"] = AcceptanceNumbers.Cum,
            ["B3"] = AcceptanceNumbers.Cum / AcceptanceNumbers.ContractAmount,
            ["B4"] = AcceptanceNumbers.Remain,
        },
        ["afterFeb20m"] = new JsonObject
        {
            ["B2"] = AcceptanceNumbers.CumChanged,
            ["B3"] = AcceptanceNumbers.CumChanged / AcceptanceNumbers.ContractAmount,
            ["B4"] = AcceptanceNumbers.RemainChanged,
        },
        ["formats"] = "column widths + 돋움 + #,##0 / 0%",
        ["namedFormulas"] = "='기성입력'!$B$2..$B$5 quoted absolute; RelativeActiveCell is a separate probe, not these totals",
        ["hyperlinkAfterRename"] = "'월별기성'!A1",
        ["pdfPagesExpected"] = 1,
        ["reopen"] = "save + close_workbook + open_workbook + export_pdf",
    };

    public static JsonObject ExpectedE6() => new()
    {
        ["plan"] = new JsonArray(AcceptanceNumbers.PlanRates.Select(v => JsonValue.Create(v)).ToArray()),
        ["actual"] = new JsonArray(AcceptanceNumbers.ActualRates.Select(v => JsonValue.Create(v)).ToArray()),
        ["initial"] = new JsonObject { ["D5"] = AcceptanceNumbers.ActualRates[2] - AcceptanceNumbers.PlanRates[2] },
        ["afterMar45"] = new JsonObject { ["D5"] = AcceptanceNumbers.MarDiffChanged },
        ["marChanged"] = AcceptanceNumbers.MarActualChanged,
        ["marDiffChanged"] = AcceptanceNumbers.MarDiffChanged,
        ["diffFormula"] = "actual - plan (C-B); initial 43-40=3; after 45-40=5",
        ["charts"] = JsonUtil.Arr("PlanActualLine", "DiffColumn"),
        ["printArea"] = "A1:D70",
        ["printTitleRows"] = "$1:$2",
        ["printTitleRowsRegression"] = "bare 1:2 vs Excel $1:$2; MatchesRequested normalizes printArea only",
        ["pageBreak"] = "A20",
        ["pdfPagesExpectedMin"] = 2,
        ["reopen"] = "save + close_workbook + open_workbook + export_pdf",
    };

    private static IEnumerable<PlannedBatch> Start(string id, string sheet, string xlsx)
    {
        yield return ApplyPlanner.Batch(id, id + "-create", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "create_workbook",
            ["sheetName"] = sheet,
        }), "token");
        yield return ApplyPlanner.Batch(id, id + "-save-as", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "save_workbook",
            ["output"] = xlsx,
        }), "token", "SaveAs before execute identity");
    }

    private static PlannedBatch Values(string id, string sheet, string range, JsonArray values)
    {
        var scenario = id.Split('-')[0];
        return ApplyPlanner.Batch(scenario, id, OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_values",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = range,
            ["values"] = values,
        }), "execute");
    }

    private static IEnumerable<PlannedBatch> E2DocumentQa(string sheet, string xlsx)
    {
        var pdf = Path.ChangeExtension(xlsx, ".pdf");
        var thin = new JsonObject { ["weight"] = "thin", ["lineStyle"] = "continuous", ["color"] = "#000000" };
        var medium = new JsonObject { ["weight"] = "medium", ["lineStyle"] = "continuous", ["color"] = "#1F4E79" };
        yield return ApplyPlanner.Batch("e2", "e2-move-dirty", OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_values",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = "A24:H24",
            ["values"] = Arr(Arr("DIRTY", "오입력 999는 합계 제외", null, null, 999, 999, null, "computed G는 G5:G8만")),
        }), "execute", "keep DIRTY fixture; place it below the designed print area");
        yield return ApplyPlanner.Batch("e2", "e2-clear-print-dirty", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "clear_range",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = "A13:H13",
            ["what"] = "contents",
        }), "token", "owned 0559 had DIRTY on row 13 inside the printed used range");
        yield return ApplyPlanner.Batch("e2", "e2-sig-merge", OpFamily.Merge, ApplyPlanner.Merges(sheet,
            ["A12:B14", "C12:D14", "E12:F14", "G12:H14"]), "token",
            "four contiguous blocks; 결재 is G12 and H12 is blank. Do not replay historical H12:H14 / empty-string 0820 merge");
        var cellBox = new JsonObject
        {
            ["top"] = thin.DeepClone(),
            ["bottom"] = thin.DeepClone(),
            ["left"] = thin.DeepClone(),
            ["right"] = thin.DeepClone(),
        };
        yield return ApplyPlanner.Batch("e2", "e2-format-title", OpFamily.Format, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A1:H1",
                ["style"] = new JsonObject
                {
                    ["fontName"] = "돋움",
                    ["fontSize"] = 16,
                    ["bold"] = true,
                    ["fontColor"] = "#FFFFFF",
                    ["fillColor"] = "#1F4E79",
                    ["horizontalAlign"] = "center",
                    ["verticalAlign"] = "center",
                    ["wrapText"] = true,
                },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A2:H2",
                ["style"] = new JsonObject
                {
                    ["fontName"] = "돋움",
                    ["fontSize"] = 10,
                    ["fillColor"] = "#D6DCE4",
                    ["verticalAlign"] = "center",
                    ["wrapText"] = true,
                    ["borders"] = new JsonObject { ["outline"] = thin.DeepClone() },
                },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A2",
                ["style"] = new JsonObject { ["bold"] = true, ["horizontalAlign"] = "center" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "E2",
                ["style"] = new JsonObject { ["bold"] = true, ["horizontalAlign"] = "right" },
            }), "execute", "title + separate metadata; F2 value not rewritten");
        yield return ApplyPlanner.Batch("e2", "e2-format-header", OpFamily.Format, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "format_range",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = "A4:H4",
            ["style"] = new JsonObject
            {
                ["fontName"] = "돋움",
                ["fontSize"] = 10,
                ["bold"] = true,
                ["fontColor"] = "#FFFFFF",
                ["fillColor"] = "#305496",
                ["horizontalAlign"] = "center",
                ["verticalAlign"] = "center",
                ["borders"] = cellBox.DeepClone(),
            },
        }), "execute", "header fill; cell-edge borders, not borders.all (0558 readback mismatch)");
        yield return ApplyPlanner.Batch("e2", "e2-format-body", OpFamily.Format, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "format_range",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = "A5:H11",
            ["style"] = new JsonObject
            {
                ["fontName"] = "돋움",
                ["fontSize"] = 10,
                ["verticalAlign"] = "center",
                ["borders"] = cellBox.DeepClone(),
            },
        }), "execute");
        yield return ApplyPlanner.Batch("e2", "e2-format-numbers", OpFamily.Format, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "E5:E8",
                ["style"] = new JsonObject { ["numberFormat"] = "#,##0.00", ["horizontalAlign"] = "right" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "F5:F8",
                ["style"] = new JsonObject { ["numberFormat"] = "#,##0", ["horizontalAlign"] = "right" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G5:G11",
                ["style"] = new JsonObject { ["numberFormat"] = "#,##0", ["horizontalAlign"] = "right" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "B9:B11",
                ["style"] = new JsonObject { ["bold"] = true, ["horizontalAlign"] = "right" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G9:G10",
                ["style"] = new JsonObject { ["bold"] = true, ["fillColor"] = "#DEEBF7" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G11",
                ["style"] = new JsonObject { ["bold"] = true, ["fillColor"] = "#FFF2CC" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "H10",
                ["style"] = new JsonObject
                {
                    ["wrapText"] = true,
                    ["fontSize"] = 8,
                    ["verticalAlign"] = "center",
                    ["horizontalAlign"] = "left",
                },
            }), "execute", "#,##0 grouping only; avoid quoted-원 mixed text-run NumberFormat");
        yield return ApplyPlanner.Batch("e2", "e2-format-sig", OpFamily.Format, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A12:H14",
                ["style"] = new JsonObject
                {
                    ["fontName"] = "돋움",
                    ["fontSize"] = 10,
                    ["bold"] = true,
                    ["horizontalAlign"] = "center",
                    ["verticalAlign"] = "center",
                    ["wrapText"] = true,
                },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A12:B14",
                ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "C12:D14",
                ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "E12:F14",
                ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G12:H14",
                ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
            }), "execute", "four contiguous signature outlines; G12:H14 closes the 승인/결재 gap");
        yield return ApplyPlanner.Batch("e2", "e2-layout-widths", OpFamily.SheetLayout, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "set_column_widths",
                ["target"] = JsonUtil.Target(sheet),
                ["columns"] = JsonUtil.Arr(
                    new JsonObject { ["col"] = "A", ["widthChars"] = 8 },
                    new JsonObject { ["col"] = "B", ["widthChars"] = 26 },
                    new JsonObject { ["col"] = "C", ["widthChars"] = 10 },
                    new JsonObject { ["col"] = "D", ["widthChars"] = 8 },
                    new JsonObject { ["col"] = "E", ["widthChars"] = 10 },
                    new JsonObject { ["col"] = "F", ["widthChars"] = 14 },
                    new JsonObject { ["col"] = "G", ["widthChars"] = 16 },
                    new JsonObject { ["col"] = "H", ["widthChars"] = 22 }),
            }), "token");
        yield return ApplyPlanner.Batch("e2", "e2-layout-heights", OpFamily.SheetLayout, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "set_row_heights",
                ["target"] = JsonUtil.Target(sheet),
                ["rows"] = JsonUtil.Arr(
                    new JsonObject { ["row"] = 1, ["autoFit"] = true },
                    new JsonObject { ["row"] = 2, ["autoFit"] = true },
                    new JsonObject { ["row"] = 3, ["heightPoints"] = 8 },
                    new JsonObject { ["row"] = 10, ["autoFit"] = true },
                    new JsonObject { ["row"] = 12, ["count"] = 3, ["autoFit"] = true }),
            }), "token", "exact heightPoints failed 0558 readback on wrapped/merged rows; autoFit instead");
        yield return ApplyPlanner.Batch("e2", "e2-layout-page", OpFamily.SheetLayout, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "set_page_setup",
                ["target"] = JsonUtil.Target(sheet),
                ["page"] = new JsonObject
                {
                    ["paperSize"] = "A4",
                    ["orientation"] = "portrait",
                    ["fitToWidth"] = 1,
                    ["fitToHeight"] = 1,
                    ["printArea"] = "A1:H14",
                    ["leftMarginMm"] = 12,
                    ["rightMarginMm"] = 12,
                    ["topMarginMm"] = 12,
                    ["bottomMarginMm"] = 12,
                    ["centerHeader"] = "공사비 수량 산출 내역서",
                    ["centerFooter"] = "기능검증용 가상 데이터",
                },
            }), "token", "DIRTY row 24 and scratch row 20 stay outside A1:H14");
        yield return ApplyPlanner.Batch("e2", "e2-sig-gap-relabel", OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_values",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = "G12:H12",
            ["values"] = Arr(Arr("결재", null)),
        }), "execute", "0820 correction: only G12:H12=['결재',null]. Leave B12/D12/F12 empty-string constants for the merge-guard regression. Do not rewrite A12:H14.");
        yield return ApplyPlanner.Batch("e2", "e2-sig-gap-merge", OpFamily.Merge, ApplyPlanner.Merges(sheet,
            ["A12:B14", "C12:D14", "E12:F14", "G12:H14"]), "token",
            "corrected four-block merge; not the historical H12:H14 0820 batch");
        yield return ApplyPlanner.Batch("e2", "e2-sig-gap-format", OpFamily.Format, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A12:H14",
                ["style"] = new JsonObject
                {
                    ["fontName"] = "돋움",
                    ["fontSize"] = 10,
                    ["bold"] = true,
                    ["horizontalAlign"] = "center",
                    ["verticalAlign"] = "center",
                    ["wrapText"] = true,
                },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A12:B14",
                ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "C12:D14",
                ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "E12:F14",
                ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G12:H14",
                ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
            }), "execute", "contiguous blue outer borders; no G-column gap");
        yield return ApplyPlanner.Batch("e2", "e2-sig-gap-save", OpFamily.Lifecycle, JsonUtil.Arr(
            new JsonObject { ["op"] = "save_workbook", ["output"] = xlsx, ["overwrite"] = true }), "token",
            "persist G12:H14 gap fix before mixed-format-prep; do not auto-close");
        yield return ApplyPlanner.Batch("e2", "e2-sig-gap-pdf", OpFamily.Lifecycle, JsonUtil.Arr(
            new JsonObject { ["op"] = "export_pdf", ["output"] = pdf, ["sheet"] = sheet, ["overwrite"] = true }), "token",
            "public PDF checkpoint before known General NumberFormat failure; do not close E2/E6");
        yield return ApplyPlanner.Batch("e2", "e2-mixed-format-prep", OpFamily.Format, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A26",
                ["style"] = new JsonObject { ["numberFormat"] = "General" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "B26",
                ["style"] = new JsonObject { ["numberFormat"] = "0.00" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "C26",
                ["style"] = new JsonObject { ["numberFormat"] = "yyyy-mm-dd" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "D26",
                ["style"] = new JsonObject { ["numberFormat"] = "\"LIT\"" },
            }), "execute", "formats first; mixed NumberFormat row is outside print A1:H14");
        yield return ApplyPlanner.Batch("e2", "e2-mixed-typed-strings", OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_values",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = "A26:D26",
            ["values"] = Arr(Arr("2026-09-10", "0012", "1E3", "=1+1")),
        }), "execute", "strict JSON strings; empty-string blanks and F2 ISO stay required separately");
        yield return ApplyPlanner.Batch("e2", "e2-height-regression", OpFamily.SheetLayout, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "set_row_heights",
                ["target"] = JsonUtil.Target(sheet),
                ["rows"] = JsonUtil.Arr(
                    new JsonObject { ["row"] = 15, ["heightPoints"] = 28 },
                    new JsonObject { ["row"] = 16, ["heightPoints"] = 22 }),
            }), "token", "0732 requested 28/22; record actual/measured mapping, do not hide with a large epsilon");
    }

    private static PlannedBatch LayoutA4(string id, string sheet, string printArea) =>
        ApplyPlanner.Batch(id, id + "-page", OpFamily.SheetLayout, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_page_setup",
            ["target"] = JsonUtil.Target(sheet),
            ["page"] = new JsonObject
            {
                ["paperSize"] = "A4",
                ["orientation"] = "portrait",
                ["printArea"] = printArea,
            },
        }), "token");

    internal static IEnumerable<PlannedBatch> SaveCloseReopen(string id, string xlsx, string pdf, string sheet)
    {
        yield return ApplyPlanner.Batch(id, id + "-save", OpFamily.Lifecycle, JsonUtil.Arr(
            new JsonObject { ["op"] = "save_workbook", ["output"] = xlsx, ["overwrite"] = true }), "token");
        yield return ApplyPlanner.Batch(id, id + "-close", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "close_workbook",
            ["saveChanges"] = false,
        }), "token", "required save-close-reopen; missing-op is a failed requirement");
        yield return ApplyPlanner.Batch(id, id + "-reopen", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "open_workbook",
            ["path"] = xlsx,
        }), "token");
        yield return ApplyPlanner.Batch(id, id + "-export-pdf", OpFamily.Lifecycle, JsonUtil.Arr(
            new JsonObject { ["op"] = "export_pdf", ["output"] = pdf, ["sheet"] = sheet, ["overwrite"] = true }), "token");
    }

    private static JsonArray Arr(params object?[] cells)
    {
        var a = new JsonArray();
        foreach (var c in cells) a.Add(JsonUtil.FromClr(c));
        return a;
    }
}
