using System.Diagnostics;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class HwpAdapter
{
    internal sealed record HwpDocxImportRequest(
        string SourceFile,
        string OutputFile,
        bool CloseAfterImport,
        int? ExpectedPageCount,
        int? ExpectedTableCount,
        IReadOnlyList<string> RequiredText);

    internal static HwpDocxImportRequest? ParseDocxImportRequest(JsonObject args)
    {
        var creationMode = Json.GetString(args, "creationMode")?.Trim().ToLowerInvariant();
        if (creationMode is not null && creationMode is not ("docx-first" or "native-hwp"))
            throw new HwpAutomationException(
                "HWP_CREATION_MODE_INVALID",
                $"creationMode는 docx-first 또는 native-hwp여야 합니다: {creationMode}",
                "새 문서 제작 전에 hwp_plan_creation으로 경로를 결정하세요.");

        var source = Json.GetString(args, "sourceFile");
        var output = Json.GetString(args, "outputFile");
        if (string.IsNullOrWhiteSpace(source))
        {
            if (creationMode == "docx-first")
                throw new HwpAutomationException(
                    "HWP_DOCX_SOURCE_REQUIRED",
                    "creationMode=docx-first에는 sourceFile이 필요합니다.",
                    "렌더 검수한 DOCX의 절대 경로를 sourceFile에 지정하세요.");
            if (!string.IsNullOrWhiteSpace(output))
                throw new HwpAutomationException(
                    "HWP_DOCX_SOURCE_REQUIRED",
                    "outputFile을 지정했지만 sourceFile이 없습니다.",
                    "sourceFile에 절대 경로의 .docx 파일을 지정하세요.");
            return null;
        }

        if (creationMode == "native-hwp")
            throw new HwpAutomationException(
                "HWP_CREATION_MODE_CONFLICT",
                "creationMode=native-hwp와 DOCX sourceFile을 함께 사용할 수 없습니다.",
                "DOCX 가져오기는 creationMode=docx-first를 사용하세요.");

        if (Json.GetBool(args, "newDocument"))
            throw new HwpAutomationException(
                "HWP_LAUNCH_MODE_CONFLICT",
                "sourceFile 가져오기와 newDocument=true를 동시에 사용할 수 없습니다.",
                "DOCX 변환은 sourceFile과 outputFile만 지정하세요.");
        if (!Path.IsPathFullyQualified(source))
            throw new HwpAutomationException(
                "HWP_DOCX_ABSOLUTE_PATH_REQUIRED",
                $"sourceFile은 절대 경로여야 합니다: {source}",
                "드라이브 문자로 시작하는 전체 경로를 지정하세요.");

        var sourcePath = CanonicalHwpPath(source);
        if (!string.Equals(Path.GetExtension(sourcePath), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new HwpAutomationException(
                "HWP_DOCX_SOURCE_FORMAT_INVALID",
                $"DOCX 우선 변환의 sourceFile은 .docx여야 합니다: {sourcePath}",
                "Word DOCX로 새 문서를 만든 뒤 그 경로를 지정하세요.");
        if (!File.Exists(sourcePath))
            throw new HwpAutomationException(
                "HWP_DOCX_SOURCE_NOT_FOUND",
                $"sourceFile을 찾을 수 없습니다: {sourcePath}",
                "DOCX 파일 경로를 확인하세요.");

        var outputValue = string.IsNullOrWhiteSpace(output)
            ? Path.ChangeExtension(sourcePath, ".hwpx")
            : output;
        if (!Path.IsPathFullyQualified(outputValue))
            throw new HwpAutomationException(
                "HWP_OUTPUT_ABSOLUTE_PATH_REQUIRED",
                $"outputFile은 절대 경로여야 합니다: {outputValue}",
                "드라이브 문자로 시작하는 전체 경로를 지정하세요.");
        var outputPath = CanonicalHwpPath(outputValue);
        var outputExtension = Path.GetExtension(outputPath);
        if (!string.Equals(outputExtension, ".hwpx", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(outputExtension, ".hwp", StringComparison.OrdinalIgnoreCase))
            throw new HwpAutomationException(
                "HWP_OUTPUT_FORMAT_INVALID",
                $"outputFile은 .hwpx 또는 .hwp여야 합니다: {outputPath}",
                "기본값인 HWPX를 권장합니다.");
        if (File.Exists(outputPath))
            throw new HwpAutomationException(
                "HWP_OUTPUT_EXISTS",
                $"출력 파일이 이미 있어 덮어쓰지 않습니다: {outputPath}",
                "다른 outputFile 이름을 지정하세요.");

        var expectedPageCount = Json.GetInt(args, "expectedPageCount");
        if (expectedPageCount is <= 0)
            throw new HwpAutomationException(
                "HWP_EXPECTED_PAGE_COUNT_INVALID",
                "expectedPageCount는 1 이상이어야 합니다.");
        var expectedTableCount = Json.GetInt(args, "expectedTableCount");
        if (expectedTableCount is < 0)
            throw new HwpAutomationException(
                "HWP_EXPECTED_TABLE_COUNT_INVALID",
                "expectedTableCount는 0 이상이어야 합니다.");
        var requiredText = (Json.GetArr(args, "requiredText") ?? new JsonArray())
            .Select(node => node?.GetValue<string>()?.Trim() ?? "")
            .Where(value => value.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        return new HwpDocxImportRequest(
            sourcePath,
            outputPath,
            Json.GetBool(args, "closeAfterImport"),
            expectedPageCount,
            expectedTableCount,
            requiredText);
    }

    internal static string HwpAutomationFormatForPath(string path) =>
        Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".hwpx" => "HWPX",
            ".docx" => "OOXML",
            _ => "HWP",
        };

    private JsonObject ImportDocxAsNativeHwp(
        dynamic hwp,
        object app,
        HwpDocxImportRequest request,
        ForegroundInteractionGuard foreground)
    {
        var totalTimer = Stopwatch.StartNew();
        var sourceHashBefore = FileHash(request.SourceFile);
        var openMilliseconds = 0L;
        var saveMilliseconds = 0L;
        var openedByThisCall = false;
        var saved = false;
        JsonObject? compatibilityAdjustment = null;

        try
        {
            EnsureFileAutomationSecurity(hwp);
            var outputDirectory = Path.GetDirectoryName(request.OutputFile);
            if (string.IsNullOrWhiteSpace(outputDirectory))
                throw new HwpAutomationException(
                    "HWP_OUTPUT_DIRECTORY_INVALID",
                    $"출력 폴더를 확인할 수 없습니다: {request.OutputFile}");
            Directory.CreateDirectory(outputDirectory);

            var timer = Stopwatch.StartNew();
            if (!OpenDocumentWithFormat(hwp, request.SourceFile, "OOXML"))
                throw new HwpAutomationException(
                    "HWP_DOCX_IMPORT_FAILED",
                    $"한글 OOXML 가져오기가 실패했습니다: {request.SourceFile}",
                    "DOCX를 Word에서 한 번 저장한 뒤 다시 시도하거나 HWP 호환성이 높은 요소로 간소화하세요.");
            timer.Stop();
            openMilliseconds = timer.ElapsedMilliseconds;
            openedByThisCall = true;

            if (ActiveDoc(hwp) is null)
                throw new HwpAutomationException(
                    "HWP_DOCX_IMPORT_NO_DOCUMENT",
                    "OOXML 가져오기 후 활성 한글 문서를 확인할 수 없습니다.");

            compatibilityAdjustment = TryCompactTrailingBlankPage(
                hwp,
                request.ExpectedPageCount);

            timer.Restart();
            SaveActiveDoc(hwp, request.OutputFile, overwrite: false);
            timer.Stop();
            saveMilliseconds = timer.ElapsedMilliseconds;
            saved = true;

            if (!File.Exists(request.OutputFile) || new FileInfo(request.OutputFile).Length == 0)
                throw new HwpAutomationException(
                    "HWP_OUTPUT_VERIFICATION_FAILED",
                    $"변환 파일이 생성되지 않았거나 비어 있습니다: {request.OutputFile}");

            var sourceHashAfter = FileHash(request.SourceFile);
            if (!string.Equals(sourceHashBefore, sourceHashAfter, StringComparison.OrdinalIgnoreCase))
                throw new HwpAutomationException(
                    "HWP_DOCX_SOURCE_CHANGED",
                    $"변환 중 원본 DOCX 해시가 바뀌어 작업을 성공으로 판정하지 않습니다: {request.SourceFile}");

            var identity = CaptureDocumentIdentity(hwp);
            var documentText = "";
            try { documentText = GetDocText(hwp); } catch { }
            var textLength = documentText.Length;
            var tableCount = 0;
            try { tableCount = CountTableControls(hwp); } catch { }
            var pageCount = 0;
            try { pageCount = Convert.ToInt32(hwp.PageCount); } catch { }
            var missingRequiredText = request.RequiredText
                .Where(value => !documentText.Contains(value, StringComparison.Ordinal))
                .ToArray();
            var pageCountMatches = request.ExpectedPageCount is null ||
                                   request.ExpectedPageCount == pageCount;
            var tableCountMatches = request.ExpectedTableCount is null ||
                                    request.ExpectedTableCount == tableCount;
            var verificationPassed = pageCountMatches && tableCountMatches &&
                                     missingRequiredText.Length == 0;

            var windowHandle = RotHelper.HwpWindowHandle(app);
            if (windowHandle != 0) foreground.TrackTargetWindow(windowHandle);
            else foreground.TrackTargetProcess(RotHelper.ProcessIdFromWindowHandle(windowHandle));
            if (request.CloseAfterImport)
            {
                CloseActiveDoc(hwp);
                openedByThisCall = false;
            }
            else
            {
                try { hwp.XHwpWindows.Active_XHwpWindow.Visible = true; } catch { }
                if (_ownsAttached && HwpUiFailureDetector.WaitForFailure(TimeSpan.FromMilliseconds(750)) is { } failure)
                    throw HwpUiInitializationException(failure);
                KeepOwnedLiveDocumentOpen(hwp);
                openedByThisCall = false;
            }

            totalTimer.Stop();
            var warnings = new List<string>();
            if (!pageCountMatches)
                warnings.Add($"페이지 수 불일치: expected={request.ExpectedPageCount}, actual={pageCount}");
            if (!tableCountMatches)
                warnings.Add($"표 수 불일치: expected={request.ExpectedTableCount}, actual={tableCount}");
            if (missingRequiredText.Length > 0)
                warnings.Add($"필수 텍스트 누락: {string.Join(", ", missingRequiredText)}");

            return new JsonObject
            {
                ["ok"] = true,
                ["app"] = App,
                ["documentRef"] = request.CloseAfterImport
                    ? request.OutputFile
                    : Json.GetString(identity, "documentRef") ?? request.OutputFile,
                ["summary"] = new JsonObject
                {
                    ["createdDocument"] = true,
                    ["creationMode"] = "docx-first",
                    ["creationPolicyVersion"] = HwpCreationPolicy.PolicyVersion,
                    ["sourceFile"] = request.SourceFile,
                    ["sourceFormat"] = "OOXML",
                    ["sourceSha256"] = sourceHashAfter,
                    ["sourceUnchanged"] = true,
                    ["outputFile"] = request.OutputFile,
                    ["outputFormat"] = HwpAutomationFormatForPath(request.OutputFile),
                    ["outputSha256"] = FileHash(request.OutputFile),
                    ["outputBytes"] = new FileInfo(request.OutputFile).Length,
                    ["textLength"] = textLength,
                    ["tableCount"] = tableCount,
                    ["pageCount"] = pageCount,
                    ["verification"] = new JsonObject
                    {
                        ["passed"] = verificationPassed,
                        ["expectedPageCount"] = request.ExpectedPageCount,
                        ["actualPageCount"] = pageCount,
                        ["pageCountMatches"] = pageCountMatches,
                        ["expectedTableCount"] = request.ExpectedTableCount,
                        ["actualTableCount"] = tableCount,
                        ["tableCountMatches"] = tableCountMatches,
                        ["requiredTextCount"] = request.RequiredText.Count,
                        ["missingRequiredText"] = Json.ToArray(missingRequiredText),
                    },
                    ["connectionMode"] = _connectionMode,
                    ["instanceRef"] = Json.GetString(identity, "instanceRef"),
                    ["visible"] = !request.CloseAfterImport,
                    ["closedAfterImport"] = request.CloseAfterImport,
                    ["compatibilityAdjustment"] = compatibilityAdjustment,
                    ["timingMs"] = new JsonObject
                    {
                        ["openOoxml"] = openMilliseconds,
                        ["saveNative"] = saveMilliseconds,
                        ["total"] = totalTimer.ElapsedMilliseconds,
                    },
                    ["instruction"] = verificationPassed
                        ? request.CloseAfterImport
                            ? "변환본을 경로로 지정해 읽거나 편집하세요."
                            : "열린 변환본을 documentRef로 지정해 검증·편집하세요."
                        : "변환은 완료됐지만 품질 게이트가 실패했습니다. 완료로 보고하지 말고 HWP PDF 렌더와 구조 차이를 수정하세요.",
                },
                ["warnings"] = Json.ToArray(warnings),
            };
        }
        catch
        {
            if (openedByThisCall)
            {
                try { CloseActiveDoc(hwp); } catch { }
            }
            if (saved && File.Exists(request.OutputFile))
            {
                // Retain only the newly-created output for diagnosis. Existing files are
                // never overwritten, and the original DOCX is never a save target.
            }
            throw;
        }
    }

    private static bool OpenDocumentWithFormat(dynamic hwp, string file, string format)
    {
        object? actionObject = null;
        object? parameterObject = null;
        object? hSetObject = null;
        try
        {
            dynamic action = hwp.HAction;
            dynamic parameters = hwp.HParameterSet.HFileSaveAs;
            dynamic hSet = parameters.HSet;
            actionObject = (object)action;
            parameterObject = (object)parameters;
            hSetObject = (object)hSet;
            action.GetDefault("FileOpen", hSet);
            parameters.OpenFileName = file;
            try { parameters.OpenFormat = format; } catch { }
            return (bool)action.Execute("FileOpen", hSet);
        }
        finally
        {
            RotHelper.ReleaseComObject(hSetObject);
            RotHelper.ReleaseComObject(parameterObject);
            RotHelper.ReleaseComObject(actionObject);
        }
    }

    internal static bool ShouldCompactTrailingBlankPage(
        int actualPageCount,
        int? expectedPageCount,
        string trailingParagraphText) =>
        expectedPageCount is > 0 &&
        actualPageCount == expectedPageCount.Value + 1 &&
        string.IsNullOrWhiteSpace(
            trailingParagraphText.Trim('\r', '\n', '\0', '\u0002', '\u0003'));

    private static JsonObject TryCompactTrailingBlankPage(dynamic hwp, int? expectedPageCount)
    {
        var beforePageCount = 0;
        try { beforePageCount = Convert.ToInt32(hwp.PageCount); } catch { }
        var result = new JsonObject
        {
            ["name"] = "compact-trailing-blank-page",
            ["attempted"] = false,
            ["applied"] = false,
            ["expectedPageCount"] = expectedPageCount,
            ["beforePageCount"] = beforePageCount,
            ["afterPageCount"] = beforePageCount,
        };

        if (expectedPageCount is null || beforePageCount != expectedPageCount.Value + 1)
        {
            result["reason"] = "page-count-not-eligible";
            return result;
        }

        var trailingParagraphText = "";
        try
        {
            hwp.HAction.Run("MoveDocEnd");
            if ((bool)hwp.HAction.Run("MoveSelParaBegin"))
                trailingParagraphText = GetSelectionText(hwp);
        }
        catch { }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
        }

        result["trailingParagraphTextLength"] = trailingParagraphText.Length;
        if (!ShouldCompactTrailingBlankPage(
                beforePageCount,
                expectedPageCount,
                trailingParagraphText))
        {
            result["reason"] = "last-paragraph-not-empty";
            return result;
        }

        result["attempted"] = true;
        try
        {
            hwp.HAction.Run("MoveDocEnd");
            var characterApplied = ApplyCharShape(hwp, new JsonObject
            {
                ["fontSize"] = 1.0,
            });
            var paragraphApplied = ApplyParagraphShape(hwp, new JsonObject
            {
                ["spaceBeforePt"] = 0.0,
                ["spaceAfterPt"] = 0.0,
                ["lineSpacingPercent"] = 50.0,
                ["widowOrphan"] = false,
                ["keepWithNext"] = false,
                ["keepLinesTogether"] = false,
                ["pageBreakBefore"] = false,
            });
            try { hwp.HAction.Run("MoveDocBegin"); } catch { }
            var afterPageCount = beforePageCount;
            try { afterPageCount = Convert.ToInt32(hwp.PageCount); } catch { }
            result["characterApplied"] = characterApplied;
            result["paragraphApplied"] = paragraphApplied;
            result["afterPageCount"] = afterPageCount;
            result["applied"] = characterApplied && paragraphApplied &&
                                afterPageCount == expectedPageCount.Value;
            result["reason"] = result["applied"]!.GetValue<bool>()
                ? "trailing-empty-paragraph-compacted"
                : "compaction-did-not-reduce-page-count";
        }
        catch (Exception ex)
        {
            result["reason"] = "compaction-error";
            result["error"] = ex.Message;
        }

        return result;
    }
}
