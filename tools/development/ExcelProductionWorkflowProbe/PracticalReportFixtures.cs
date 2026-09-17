namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Drop-in public DocBridge PlannedBatch sequences for useful E3/E6/E8/E9 documents.
/// Copy these three staging files into the harness project and point
/// <c>ScenarioRunner.Plan/Expected</c> at the methods below. Do not author workbooks
/// through openpyxl/COM/raster. Pictures are labelled TinyPng samples, not site photos.
/// </summary>
internal static class PracticalReportFixtures
{
    public static List<PlannedBatch> E3(string artifactDir, string sourceXlsx, string pictureDir)
    {
        const string sheet = "작업일보";
        var xlsx = Path.Combine(artifactDir, "e3-daily-report.xlsx");
        var pdf = Path.Combine(artifactDir, "e3-daily-report.pdf");
        SourceIntegrity.RefuseIfProtectedPath(xlsx, sourceXlsx);
        var pic1 = TinyPng.WriteSample(pictureDir, "e3-photo-1.png", 40, 90, 160);
        var pic2 = TinyPng.WriteSample(pictureDir, "e3-photo-2.png", 160, 90, 40);
        File.WriteAllText(Path.Combine(pictureDir, "e3-sample-images.txt"),
            "e3-photo-1.png / e3-photo-2.png 는 TinyPng 라벨 샘플이다. 현장 사진이 아님.\n",
            new UTF8Encoding(false));

        var longNote = "관로 터파기 및 부설 구간에서 기존 지장물 이설, 우수 유입 대비 양수, 야간 교통 통제 협의, 품질시험 의뢰, 인접 민원 대응을 포함한 당일 작업 내용을 기록한다. " +
                       string.Join(" ", Enumerable.Repeat("추가 작업 메모.", 40));

        var thin = Edge("thin", "#000000");
        var medium = Edge("medium", "#1F4E79");
        var cellBox = new JsonObject
        {
            ["top"] = thin.DeepClone(),
            ["bottom"] = thin.DeepClone(),
            ["left"] = thin.DeepClone(),
            ["right"] = thin.DeepClone(),
        };
        var e3Objects = PracticalLayoutMetrics.E3Objects();
        var photo1 = PracticalLayoutMetrics.Position(e3Objects[0]);
        var photo2 = PracticalLayoutMetrics.Position(e3Objects[1]);

        return
        [
            .. Start("e3", sheet, xlsx),
            Values("e3", "e3-values", sheet, "A1:F10", new JsonArray
            {
                Arr("작업일보 (기능검증용 가상 데이터)", null, null, null, null, null),
                Arr("현장", "수원시 하수관로 정비", "작성일", "2026-09-10", "날씨", "맑음"),
                Arr("담당", "김현장", "작업구간", "고색처리분구 STA.0+120~0+180", "공종", "오수관로 신설"),
                Arr("구분", "인원", "장비", "작업량", "단위", "비고"),
                Arr("기능공", 6, 2, 18, "m", null),
                Arr("보통인부", 4, 1, null, null, null),
                Arr("관리", 3, null, null, null, null),
                Arr("합계", null, null, null, null, null),
                Arr("작업내용", longNote, null, null, null, null),
                Arr("사진", "현장 기록 사진(샘플 2장). 실제 현장이 아님.", null, null, null, null),
            }),
            Values("e3", "e3-signatures", sheet, "A21:F23", new JsonArray
            {
                Arr("작성", null, "검토", null, "현장대리인", null),
                Arr(null, null, null, null, null, null),
                Arr(null, null, null, null, null, null),
            }),
            ApplyPlanner.Batch("e3", "e3-formulas", OpFamily.Values, JsonUtil.Arr(
                new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = "B8", ["formulas"] = Arr(Arr("=B5+B6+B7")) },
                new JsonObject { ["op"] = "set_formulas", ["target"] = JsonUtil.Target(sheet), ["range"] = "C8", ["formulas"] = Arr(Arr("=C5+C6")) }),
                "execute", "B8 people=13, C8 equipment=3"),
            ApplyPlanner.Batch("e3", "e3-merge", OpFamily.Merge, ApplyPlanner.Merges(sheet,
                ["A1:F1", "B9:F9", "B10:F10", "A21:B23", "C21:D23", "E21:F23"]), "token",
                "title + wrapped note + photo caption + three signature blocks after the photo band"),
            ApplyPlanner.Batch("e3", "e3-format-title", OpFamily.Format, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A1:F1",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 16,
                        ["bold"] = true,
                        ["fontColor"] = "#FFFFFF",
                        ["fillColor"] = "#1F4E79",
                        ["horizontalAlign"] = "center",
                        ["verticalAlign"] = "center",
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A2:F3",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["fillColor"] = "#D6DCE4",
                        ["verticalAlign"] = "center",
                        ["wrapText"] = true,
                        ["borders"] = new JsonObject { ["outline"] = thin.DeepClone() },
                    },
                }), "execute", "title/metadata wrap; D3 station text uses row3=42"),
            ApplyPlanner.Batch("e3", "e3-format-table", OpFamily.Format, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A4:F4",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["bold"] = true,
                        ["fontColor"] = "#FFFFFF",
                        ["fillColor"] = "#305496",
                        ["horizontalAlign"] = "center",
                        ["borders"] = cellBox.DeepClone(),
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A5:F8",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["verticalAlign"] = "center",
                        ["borders"] = cellBox.DeepClone(),
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A8:F8",
                    ["style"] = new JsonObject { ["bold"] = true, ["fillColor"] = "#DEEBF7" },
                }), "execute"),
            ApplyPlanner.Batch("e3", "e3-wrap", OpFamily.Format, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "B9",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["wrapText"] = true,
                        ["verticalAlign"] = "top",
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "B10",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 9,
                        ["wrapText"] = true,
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A21:F23",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["bold"] = true,
                        ["horizontalAlign"] = "center",
                        ["verticalAlign"] = "center",
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A21:B23",
                    ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "C21:D23",
                    ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "E21:F23",
                    ["style"] = new JsonObject { ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() } },
                }), "execute", "long Korean wrap + readable signatures after photos"),
            ApplyPlanner.Batch("e3", "e3-pictures", OpFamily.Data, JsonUtil.Arr(
                Picture(sheet, pic1, "Photo1", photo1),
                Picture(sheet, pic2, "Photo2", photo2)), "token",
                "two independent TinyPng samples in reserved space below B9; not site photos"),
            ApplyPlanner.Batch("e3", "e3-layout", OpFamily.SheetLayout, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_column_widths",
                    ["target"] = JsonUtil.Target(sheet),
                    ["columns"] = JsonUtil.Arr(
                        new JsonObject { ["col"] = "A", ["widthChars"] = 14 },
                        new JsonObject { ["col"] = "B", ["widthChars"] = 16 },
                        new JsonObject { ["col"] = "C", ["widthChars"] = 14 },
                        new JsonObject { ["col"] = "D", ["widthChars"] = 12 },
                        new JsonObject { ["col"] = "E", ["widthChars"] = 14 },
                        new JsonObject { ["col"] = "F", ["widthChars"] = 16 }),
                },
                new JsonObject
                {
                    ["op"] = "set_row_heights",
                    ["target"] = JsonUtil.Target(sheet),
                    ["rows"] = E3HeightOps(),
                },
                new JsonObject
                {
                    ["op"] = "set_page_setup",
                    ["target"] = JsonUtil.Target(sheet),
                    ["page"] = new JsonObject
                    {
                        ["paperSize"] = "A4",
                        ["orientation"] = "portrait",
                        ["scale"] = 100,
                        ["printArea"] = "A1:F23",
                        ["leftMarginMm"] = 12,
                        ["rightMarginMm"] = 12,
                        ["topMarginMm"] = 12,
                        ["bottomMarginMm"] = 12,
                        ["centerHeader"] = "작업일보",
                        ["centerFooter"] = "기능검증용 가상 데이터 · 샘플 이미지",
                    },
                }), "token", "integer scale 100; omit fitToWidth/fitToHeight so native defaults cannot change page count"),
            .. ScenarioCore.SaveCloseReopen("e3", xlsx, pdf, sheet),
        ];
    }

    public static List<PlannedBatch> E6(string artifactDir, string sourceXlsx, string pictureDir)
    {
        const string sheet = "월간보고";
        var xlsx = Path.Combine(artifactDir, "e6-monthly-report.xlsx");
        var pdf = Path.Combine(artifactDir, "e6-monthly-report.pdf");
        SourceIntegrity.RefuseIfProtectedPath(xlsx, sourceXlsx);
        var sample = TinyPng.WriteSample(pictureDir, "e6-logo.png", 20, 20, 20);
        File.WriteAllText(Path.Combine(pictureDir, "e6-sample-images.txt"),
            "e6-logo.png 는 TinyPng 라벨 샘플이다. 현장 사진이 아님.\n",
            new UTF8Encoding(false));

        var months = new[] { "1월", "2월", "3월", "4월", "5월", "6월" };
        var table = new JsonArray
        {
            Arr("월간 공정 보고서 (기능검증용 가상 데이터)", null, null, null),
            Arr("월", "계획누계율", "실적누계율", "차이"),
        };
        for (var i = 0; i < 6; i++)
            table.Add(Arr(months[i], AcceptanceNumbers.PlanRates[i], AcceptanceNumbers.ActualRates[i], null));
        table.Add(Arr("계획 및 실적 누계 추이", "1~6월 누계율. 기능검증용 가상 데이터.", null, null));

        var diffs = new JsonArray();
        for (var r = 3; r <= 8; r++)
        {
            diffs.Add(new JsonObject
            {
                ["op"] = "set_formulas",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = $"D{r}",
                ["formulas"] = Arr(Arr($"=C{r}-B{r}")),
            });
        }

        var thin = Edge("thin", "#000000");
        var medium = Edge("medium", "#1F4E79");
        var cellBox = new JsonObject
        {
            ["top"] = thin.DeepClone(),
            ["bottom"] = thin.DeepClone(),
            ["left"] = thin.DeepClone(),
            ["right"] = thin.DeepClone(),
        };
        var e6Objects = PracticalLayoutMetrics.E6Objects();
        var chart1 = PracticalLayoutMetrics.Position(e6Objects[0]);
        var chart2 = PracticalLayoutMetrics.Position(e6Objects[1]);
        var mark = PracticalLayoutMetrics.Position(e6Objects[2]);

        return
        [
            .. Start("e6", sheet, xlsx),
            Values("e6", "e6-values", sheet, "A1:D9", table),
            Values("e6", "e6-page2-copy", sheet, "A22:D23", new JsonArray
            {
                Arr("월별 편차 분석", null, null, null),
                Arr("계획 대비 실적 차이", "현장 기록 마크(샘플). 실제 현장이 아님.", null, null),
            }),
            Values("e6", "e6-analysis", sheet, "A36:D45", new JsonArray
            {
                Arr("분석", "3월 실적이 계획을 앞선 구간을 확인하고, 4월 이후 차이를 좁히는 조치를 기록한다.", null, null),
                Arr("요지", "1~2월은 계획 대비 지연, 3월은 실적 43(변경 후 45)으로 계획 40을 상회한다. 4월은 실적 57 / 계획 60으로 다시 부족하다. 5~6월 누계는 계획에 근접한다.", null, null),
                Arr("원인", "3월 야간 작업 투입과 자재 선행 반입으로 실적이 앞당겨졌다. 4월은 우천과 교통 통제로 굴착이 중단된 날이 있다.", null, null),
                Arr("리스크", "6월 100 마감을 위해 4월 부족분 3p를 5월 작업에 흡수하지 못하면 누계 곡선이 다시 벌어진다.", null, null),
                Arr("조치", "1) 4월 우천 만회 야간 2일 편성 2) D300 관 선반입 확인 3) 교통 통제 협의 재개", null, null),
                Arr("담당", "공정: 김현장 / 자재: 이민수 / 안전: 최감독", null, null),
                Arr("결재", "기능검증용 가상 데이터.", null, null),
                Arr("작성", null, "검토", null),
                Arr(null, null, null, null),
                Arr(null, null, null, null),
            }),
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
                }), "execute", "F2:G8 helper remains valid; DiffColumn uses union A2:A8,D2:D8 + series"),
            ApplyPlanner.Batch("e6", "e6-merge", OpFamily.Merge, ApplyPlanner.Merges(sheet,
                PracticalLayoutMetrics.E6MergeRanges), "token",
                "A22:C22 / B23:C23 leave D22:D23 empty; analysis wrap 36-41; signatures 43-45"),
            ApplyPlanner.Batch("e6", "e6-format", OpFamily.Format, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A1:D1",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 16,
                        ["bold"] = true,
                        ["fontColor"] = "#FFFFFF",
                        ["fillColor"] = "#1F4E79",
                        ["horizontalAlign"] = "center",
                        ["verticalAlign"] = "center",
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A2:D2",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["bold"] = true,
                        ["fontColor"] = "#FFFFFF",
                        ["fillColor"] = "#305496",
                        ["horizontalAlign"] = "center",
                        ["borders"] = cellBox.DeepClone(),
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A3:D8",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["horizontalAlign"] = "center",
                        ["borders"] = cellBox.DeepClone(),
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "B3:D8",
                    ["style"] = new JsonObject { ["numberFormat"] = "0" },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A9:D9",
                    ["style"] = new JsonObject { ["fontName"] = "돋움", ["fontSize"] = 9, ["wrapText"] = true },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A22:C23",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 11,
                        ["bold"] = true,
                        ["fillColor"] = "#D6DCE4",
                        ["wrapText"] = true,
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A36:D42",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["wrapText"] = true,
                        ["verticalAlign"] = "top",
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A36:A42",
                    ["style"] = new JsonObject { ["bold"] = true, ["fillColor"] = "#DEEBF7" },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A43:B45",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["bold"] = true,
                        ["horizontalAlign"] = "center",
                        ["verticalAlign"] = "center",
                        ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() },
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "C43:D45",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["bold"] = true,
                        ["horizontalAlign"] = "center",
                        ["verticalAlign"] = "center",
                        ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() },
                    },
                }), "execute", "numeric table + useful page-2 analysis; no print-padding labels"),
            ApplyPlanner.Batch("e6", "e6-charts", OpFamily.Data, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "create_chart",
                    ["target"] = JsonUtil.Target(sheet),
                    ["name"] = "PlanActualLine",
                    ["chartType"] = "lineMarkers",
                    ["sourceRange"] = "A2:C8",
                    ["title"] = "계획 및 실적 누계 추이",
                    ["hasLegend"] = true,
                    ["legendPosition"] = "bottom",
                    ["position"] = chart1.DeepClone(),
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
                    ["title"] = "월별 편차 분석",
                    ["hasLegend"] = true,
                    ["legendPosition"] = "bottom",
                    ["position"] = chart2.DeepClone(),
                    ["series"] = JsonUtil.Arr(
                        new JsonObject { ["values"] = "D2:D8", ["categories"] = "A2:A8", ["name"] = "수량" }),
                },
                Picture(sheet, sample, "Logo", mark)), "token",
                "data pin: series values/categories/name applied; DiffColumn union A2:A8,D2:D8; helper F2:G8 stays valid"),
            ApplyPlanner.Batch("e6", "e6-print", OpFamily.SheetLayout, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_column_widths",
                    ["target"] = JsonUtil.Target(sheet),
                    ["columns"] = JsonUtil.Arr(
                        new JsonObject { ["col"] = "A", ["widthChars"] = 22 },
                        new JsonObject { ["col"] = "B", ["count"] = 3, ["widthChars"] = 20 }),
                },
                new JsonObject
                {
                    ["op"] = "set_row_heights",
                    ["target"] = JsonUtil.Target(sheet),
                    ["rows"] = E6HeightOps(),
                },
                new JsonObject
                {
                    ["op"] = "set_page_setup",
                    ["target"] = JsonUtil.Target(sheet),
                    ["page"] = new JsonObject
                    {
                        ["paperSize"] = "A4",
                        ["orientation"] = "portrait",
                        ["scale"] = 100,
                        ["printArea"] = "A1:D45",
                        ["printTitleRows"] = "$1:$2",
                        ["leftMarginMm"] = 12,
                        ["rightMarginMm"] = 12,
                        ["topMarginMm"] = 12,
                        ["bottomMarginMm"] = 12,
                        ["centerHeader"] = "월간 공정 보고서",
                        ["centerFooter"] = "기능검증용 가상 데이터",
                    },
                }), "token",
                "integer scale 100, no fit keys; chart 340pt must sit inside narrowest MDW sandwich of A-D 22/20/20/20; helper F:G outside print"),
            ApplyPlanner.Batch("e6", "e6-page-break", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_page_breaks",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A22",
                ["orientation"] = "row",
            }), "token",
                "empty boundary between page layouts. 0750 0x800A03EC on cell.PageBreak remains an open product regression; keep this op."),
            ApplyPlanner.Batch("e6", "e6-recalc-mar", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "C5",
                ["values"] = Arr(Arr(AcceptanceNumbers.MarActualChanged)),
            }), "execute", "3월 실적 45 → D5 차이 5"),
            .. ScenarioCore.SaveCloseReopen("e6", xlsx, pdf, sheet),
        ];
    }

    public static List<PlannedBatch> E8(string artifactDir, string sourceXlsx)
    {
        const string sheet = "입력양식";
        var xlsx = Path.Combine(artifactDir, "e8-protected-form.xlsx");
        var pdf = Path.ChangeExtension(xlsx, ".pdf");
        SourceIntegrity.RefuseIfProtectedPath(xlsx, sourceXlsx);
        var thin = Edge("thin", "#000000");
        var medium = Edge("medium", "#1F4E79");
        var cellBox = new JsonObject
        {
            ["top"] = thin.DeepClone(),
            ["bottom"] = thin.DeepClone(),
            ["left"] = thin.DeepClone(),
            ["right"] = thin.DeepClone(),
        };

        return
        [
            .. Start("e8", sheet, xlsx),
            Values("e8", "e8-values", sheet, "A1:B3", new JsonArray
            {
                Arr("수량 (입력)", 10),
                Arr("단가 (입력)", 1000),
                Arr("금액 (수식 잠금)", null),
            }),
            Values("e8", "e8-form-chrome", sheet, "D1:G8", new JsonArray
            {
                Arr("현장 수량 입력 양식 (기능검증용 가상 데이터)", null, null, null),
                Arr("노란 칸은 입력, 금액은 자동 계산. 기능검증용 가상 데이터.", null, null, null),
                Arr("현장", "수원시 하수관로 정비", "작성일", "2026-09-10"),
                Arr("구분", "금일 입력", "상태", "작성중"),
                Arr("비고", "입력 후 작성·검토란에 서명한다.", null, null),
                Arr(null, null, null, null),
                Arr("작성", null, "검토", null),
                Arr(null, null, null, null),
            }),
            ApplyPlanner.Batch("e8", "e8-formula", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_formulas",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "B3",
                ["formulas"] = new JsonArray { new JsonArray { "=B1*B2" } },
            }), "execute", "B3=10000 from 10*1000"),
            ApplyPlanner.Batch("e8", "e8-merge", OpFamily.Merge, ApplyPlanner.Merges(sheet,
                ["D1:G1", "D2:G2", "E5:G5", "D7:E8", "F7:G8"]), "token"),
            ApplyPlanner.Batch("e8", "e8-format-form", OpFamily.Format, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A1:B3",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["borders"] = cellBox.DeepClone(),
                        ["verticalAlign"] = "center",
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A1:A3",
                    ["style"] = new JsonObject { ["bold"] = true, ["fillColor"] = "#D6DCE4", ["locked"] = true },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "B1:B2",
                    ["style"] = new JsonObject
                    {
                        ["locked"] = false,
                        ["fillColor"] = "#FFF2CC",
                        ["numberFormat"] = "#,##0",
                        ["horizontalAlign"] = "right",
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "B3",
                    ["style"] = new JsonObject
                    {
                        ["locked"] = true,
                        ["bold"] = true,
                        ["fillColor"] = "#DEEBF7",
                        ["numberFormat"] = "#,##0",
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "D1:G2",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 12,
                        ["bold"] = true,
                        ["fillColor"] = "#1F4E79",
                        ["fontColor"] = "#FFFFFF",
                        ["wrapText"] = true,
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "D3:G8",
                    ["style"] = new JsonObject { ["fontName"] = "돋움", ["fontSize"] = 10, ["wrapText"] = true },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "D7:E8",
                    ["style"] = new JsonObject
                    {
                        ["bold"] = true,
                        ["horizontalAlign"] = "center",
                        ["verticalAlign"] = "center",
                        ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() },
                    },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "F7:G8",
                    ["style"] = new JsonObject
                    {
                        ["bold"] = true,
                        ["horizontalAlign"] = "center",
                        ["verticalAlign"] = "center",
                        ["borders"] = new JsonObject { ["outline"] = medium.DeepClone() },
                    },
                }), "execute", "practical form chrome; B1:B2 unlocked inputs; B3 locked formula"),
            ApplyPlanner.Batch("e8", "e8-validation-note", OpFamily.Data, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_data_validation",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "B1:B2",
                    ["type"] = "decimal",
                    ["operator"] = "greaterEqual",
                    ["formula1"] = "0",
                },
                new JsonObject
                {
                    ["op"] = "set_cell_note",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "B3",
                    ["text"] = "잠긴 수식 =B1*B2. 직접 입력 거부.",
                    ["visible"] = false,
                }), "token"),
            ApplyPlanner.Batch("e8", "e8-layout", OpFamily.SheetLayout, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_column_widths",
                    ["target"] = JsonUtil.Target(sheet),
                    ["columns"] = JsonUtil.Arr(
                        new JsonObject { ["col"] = "A", ["widthChars"] = 18 },
                        new JsonObject { ["col"] = "B", ["widthChars"] = 14 },
                        new JsonObject { ["col"] = "C", ["widthChars"] = 3 },
                        new JsonObject { ["col"] = "D", ["widthChars"] = 12 },
                        new JsonObject { ["col"] = "E", ["count"] = 3, ["widthChars"] = 18 }),
                },
                new JsonObject
                {
                    ["op"] = "set_row_heights",
                    ["target"] = JsonUtil.Target(sheet),
                    ["rows"] = JsonUtil.Arr(
                        new JsonObject { ["row"] = 1, ["heightPoints"] = 24 },
                        new JsonObject { ["row"] = 2, ["heightPoints"] = 30 },
                        new JsonObject { ["row"] = 7, ["count"] = 2, ["heightPoints"] = 24 }),
                },
                new JsonObject
                {
                    ["op"] = "set_page_setup",
                    ["target"] = JsonUtil.Target(sheet),
                    ["page"] = new JsonObject
                    {
                        ["paperSize"] = "A4",
                        ["orientation"] = "portrait",
                        ["printArea"] = "A1:G8",
                        ["centerFooter"] = "기능검증용 가상 데이터",
                    },
                }), "token"),
            ApplyPlanner.Batch("e8", "e8-protect", OpFamily.Protect, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "protect_sheet",
                ["target"] = JsonUtil.Target(sheet),
            }), "token", "no password field; secrets never logged"),
            ApplyPlanner.Batch("e8", "e8-unprotect", OpFamily.Protect, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "unprotect_sheet",
                ["target"] = JsonUtil.Target(sheet),
            }), "token", "edit cycle: unprotect before touching the locked-formula neighborhood"),
            ApplyPlanner.Batch("e8", "e8-unprotect-edit", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G4",
                ["values"] = Arr(Arr("재보호 준비")),
            }), "execute", "do not rewrite B1/B2/B3; LiveVerifier still owns those cells"),
            ApplyPlanner.Batch("e8", "e8-reprotect", OpFamily.Protect, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "protect_sheet",
                ["target"] = JsonUtil.Target(sheet),
            }), "token", "final state must stay protected so locked B3 write is refused"),
            .. ScenarioCore.SaveCloseReopen("e8", xlsx, pdf, sheet),
        ];
    }

    public static List<PlannedBatch> E9(string artifactDir, string sourceXlsx, string fixtureDir)
    {
        const string sheet = "정리";
        var xlsx = Path.Combine(artifactDir, "e9-csv-cleanup.xlsx");
        var pdf = Path.ChangeExtension(xlsx, ".pdf");
        SourceIntegrity.RefuseIfProtectedPath(xlsx, sourceXlsx);
        Directory.CreateDirectory(fixtureDir);
        var csv = PracticalCsvRecordParser.WriteSampleFile(fixtureDir);
        JsonUtil.Write(Path.Combine(fixtureDir, "e9-oracle.json"), PracticalCsvRecordParser.Oracle(csv));

        return
        [
            .. Start("e9", sheet, xlsx),
            ApplyPlanner.Batch("e9", "e9-import", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "import_csv",
                ["path"] = csv,
                ["target"] = JsonUtil.Target(sheet),
                ["destination"] = "A1",
            }), "token",
                "ApplyImportCsv uses ExcelCsvContract.Parse onto 정리!A1; this grid feeds dedupe/sort/export"),
            ApplyPlanner.Batch("e9", "e9-wrap-note", OpFamily.Format, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "A2:G4",
                    ["style"] = new JsonObject { ["numberFormat"] = "@" },
                },
                new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheet),
                    ["range"] = "D2:D4",
                    ["style"] = new JsonObject
                    {
                        ["fontName"] = "돋움",
                        ["fontSize"] = 10,
                        ["wrapText"] = true,
                        ["verticalAlign"] = "top",
                        ["numberFormat"] = "@",
                    },
                }), "execute", "keep imported text (@) including 0012 and =1+1; wrap notes only; no value rewrite"),
            ApplyPlanner.Batch("e9", "e9-text-to-columns", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "text_to_columns",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "G2:G4",
                ["destination"] = "I2",
                ["dataType"] = "delimited",
                ["comma"] = false,
                ["tab"] = false,
                ["semicolon"] = false,
                ["space"] = false,
                ["other"] = true,
                ["otherChar"] = "|",
            }), "token", "split imported packed column G; destination I stays off A:C"),
            ApplyPlanner.Batch("e9", "e9-dedupe", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "remove_duplicates",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A1:G4",
                ["hasHeaders"] = true,
            }), "token", "unique data rows must stay 2 for LiveVerifier A1:C4"),
            ApplyPlanner.Batch("e9", "e9-sort", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "sort_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A1:G3",
                ["hasHeaders"] = true,
                ["keys"] = JsonUtil.Arr(new JsonObject { ["column"] = "id", ["order"] = "asc" }),
            }), "token"),
            ApplyPlanner.Batch("e9", "e9-export", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "export_csv",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A1:G3",
                ["output"] = Path.Combine(artifactDir, "e9-clean.csv"),
                ["overwrite"] = true,
            }), "token", "export the imported-then-cleaned table, not a parser-authored grid"),
            ApplyPlanner.Batch("e9", "e9-layout", OpFamily.SheetLayout, JsonUtil.Arr(
                new JsonObject
                {
                    ["op"] = "set_column_widths",
                    ["target"] = JsonUtil.Target(sheet),
                    ["columns"] = JsonUtil.Arr(
                        new JsonObject { ["col"] = "A", ["widthChars"] = 8 },
                        new JsonObject { ["col"] = "B", ["widthChars"] = 16 },
                        new JsonObject { ["col"] = "C", ["widthChars"] = 8 },
                        new JsonObject { ["col"] = "D", ["widthChars"] = 14 },
                        new JsonObject { ["col"] = "E", ["count"] = 3, ["widthChars"] = 12 }),
                },
                new JsonObject
                {
                    ["op"] = "set_row_heights",
                    ["target"] = JsonUtil.Target(sheet),
                    ["rows"] = JsonUtil.Arr(new JsonObject { ["row"] = 2, ["count"] = 2, ["heightPoints"] = 30 }),
                },
                new JsonObject
                {
                    ["op"] = "set_page_setup",
                    ["target"] = JsonUtil.Target(sheet),
                    ["page"] = new JsonObject
                    {
                        ["paperSize"] = "A4",
                        ["orientation"] = "portrait",
                        ["printArea"] = "A1:G3",
                        ["centerFooter"] = "기능검증용 가상 데이터",
                    },
                }), "token"),
            .. ScenarioCore.SaveCloseReopen("e9", xlsx, pdf, sheet),
        ];
    }

    public static JsonObject ExpectedE3() => new()
    {
        ["people"] = AcceptanceNumbers.DailyPeople,
        ["equipment"] = AcceptanceNumbers.DailyEquipment,
        ["acceptanceCells"] = new JsonObject { ["B8"] = 13, ["C8"] = 3 },
        ["pictures"] = 2,
        ["picturesAreSamples"] = true,
        ["pictureNames"] = JsonUtil.Arr("Photo1", "Photo2"),
        ["wrap"] = "B9",
        ["printArea"] = "A1:F23",
        ["layout"] = PracticalLayoutMetrics.E3Bounds(),
        ["signatures"] = "A21:F23; photo/signature tops from cumulative row heights",
        ["pdfPagesExpectedMin"] = 1,
        ["liveVerifierHooks"] = JsonUtil.Arr(
            "keep B8/C8 numeric checks",
            "keep pictures-2 and wrap-B9",
            "RangeOf A1:F12 still covers the table/note; print area is now A1:F23",
            "do not treat TinyPng samples as site photos"),
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
        ["chartSources"] = new JsonObject
        {
            ["PlanActualLine"] = "A2:C8 + series B2:B8/C2:C8 categories A2:A8",
            ["DiffColumn"] = "union A2:A8,D2:D8 + series values D2:D8 categories A2:A8 name 수량; F2:G8 helper remains valid",
        },
        ["dataContract"] = "cursor-data/pivot-chart-20260911/REPORT.md",
        ["printArea"] = "A1:D45",
        ["printTitleRows"] = "$1:$2",
        ["pageBreak"] = "A22",
        ["pageBreakRegression"] = "0750 cell.PageBreak 0x800A03EC; keep the op. LiveVerifier does not assert the address.",
        ["printAreaVerifierHook"] = "replace print-area-A1D70 with A1:D45; no 인쇄여백/작업항목 padding rows",
        ["pdfPagesExpectedMin"] = 2,
        ["layout"] = PracticalLayoutMetrics.E6Bounds(),
        ["liveVerifierHooks"] = JsonUtil.Arr(
            "keep initial D5=3 and after C5=45 D5=5",
            "keep charts-2 and RangeOf A1:D8",
            "update print-area-A1D70 → A1:D45",
            "visual: both charts fully inside A-D, all six months, unobscured repeated $1:$2, useful page-2 analysis/sign-off",
            "0750 manual-break and 0752 lifecycle failures stay open"),
        ["reopen"] = "save + close_workbook + open_workbook + export_pdf",
    };

    public static JsonObject ExpectedE8() => new()
    {
        ["amount"] = 10_000,
        ["acceptanceCells"] = new JsonObject { ["B1"] = 10, ["B2"] = 1000, ["B3"] = 10_000 },
        ["protect"] = "locked B3 write refused; unlocked B1 write allowed; no password field",
        ["cycle"] = "protect_sheet → unprotect_sheet → set_values G4 → protect_sheet",
        ["finalState"] = "protected, B1:B2 unlocked, B3 locked formula =B1*B2",
        ["liveVerifierHooks"] = JsonUtil.Arr(
            "keep B3=10000, locked-b3-write-refused, unlocked-b1-write-allowed",
            "RangeOf A1:B3 unchanged",
            "b3-still-10000 after B1=11 assumes no recalc; do not rewrite B1/B3 in the fixture"),
        ["reopen"] = "save + close_workbook + open_workbook + export_pdf",
    };

    public static JsonObject ExpectedE9()
    {
        var oracle = PracticalCsvRecordParser.OracleFromCanonical();
        return new JsonObject
        {
            ["uniqueRows"] = 2,
            ["oracleRole"] = "read-only; never set_values this matrix onto 정리",
            ["acceptanceSource"] = "import_csv 정리!A1 (ExcelCsvContract.Parse) → text_to_columns G → remove_duplicates → sort → export_csv A1:G3",
            ["csvFeatures"] = JsonUtil.Arr("quoted comma", "quoted CRLF", "Korean", "leading-zero string", "formula-prefixed literal"),
            ["acceptanceGrid"] = "정리!A1:G3 after imported dedupe/sort; A1:C4 still has 2 data rows",
            ["sku"] = JsonUtil.Arr("0012", "0003"),
            ["literals"] = JsonUtil.Arr("=1+1", "=2+2"),
            ["textToColumns"] = "imported G2:G4 → I2 otherChar=|",
            ["oracle"] = oracle,
            ["liveVerifierHooks"] = JsonUtil.Arr(
                "keep unique-data-rows-2 and exported-csv-unique-data-2",
                "RangeOf A1:C4 unchanged",
                "assert imported cells against oracle: quoted comma, quoted CRLF, 0012, =1+1",
                "missing import_csv / export_csv / remove_duplicates / text_to_columns is a failed requirement"),
            ["failedRequirementUnless"] = "import_csv / export_csv / remove_duplicates / text_to_columns advertised",
        };
    }

    /// <summary>
    /// Optional one-line harness hook: <c>PracticalReportFixtures.TryPlan(id, ...) ?? existing</c>.
    /// </summary>
    public static List<PlannedBatch>? TryPlan(string id, string artifactDir, string sourceXlsx, string pictureOrFixtureDir) =>
        id switch
        {
            "e3" => E3(artifactDir, sourceXlsx, pictureOrFixtureDir),
            "e6" => E6(artifactDir, sourceXlsx, pictureOrFixtureDir),
            "e8" => E8(artifactDir, sourceXlsx),
            "e9" => E9(artifactDir, sourceXlsx, pictureOrFixtureDir),
            _ => null,
        };

    public static JsonObject? TryExpected(string id) => id switch
    {
        "e3" => ExpectedE3(),
        "e6" => ExpectedE6(),
        "e8" => ExpectedE8(),
        "e9" => ExpectedE9(),
        _ => null,
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

    private static PlannedBatch Values(string scenario, string id, string sheet, string range, JsonArray values) =>
        ApplyPlanner.Batch(scenario, id, OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_values",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = range,
            ["values"] = values,
        }), "execute");

    private static JsonObject Picture(string sheet, string path, string name, JsonObject position) => new()
    {
        ["op"] = "insert_sheet_picture",
        ["target"] = JsonUtil.Target(sheet),
        ["path"] = path,
        ["name"] = name,
        ["position"] = position.DeepClone(),
    };

    private static JsonObject Edge(string weight, string color) => new()
    {
        ["weight"] = weight,
        ["lineStyle"] = "continuous",
        ["color"] = color,
    };

    private static JsonArray E3HeightOps()
    {
        var heights = PracticalLayoutMetrics.E3RowHeights();
        var ops = new JsonArray();
        for (var i = 0; i < heights.Length; i++)
            ops.Add(new JsonObject { ["row"] = i + 1, ["heightPoints"] = heights[i] });
        return ops;
    }

    private static JsonArray E6HeightOps()
    {
        var heights = PracticalLayoutMetrics.E6RowHeights();
        var ops = new JsonArray();
        for (var i = 0; i < heights.Length; i++)
            ops.Add(new JsonObject { ["row"] = i + 1, ["heightPoints"] = heights[i] });
        return ops;
    }

    private static JsonArray Arr(params object?[] cells)
    {
        var a = new JsonArray();
        foreach (var c in cells) a.Add(JsonUtil.FromClr(c));
        return a;
    }
}
