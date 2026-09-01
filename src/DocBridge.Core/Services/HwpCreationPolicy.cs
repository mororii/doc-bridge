using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// 새 한글 문서를 DOCX 우선 경로로 만들지, HWP Automation으로 직접 만들지
/// 결정하는 부작용 없는 정책 평가기다. 실제 파일 생성 전에 AI 클라이언트가
/// 동일한 기준으로 경로를 선택할 수 있도록 결과와 품질 게이트를 함께 반환한다.
/// </summary>
public static class HwpCreationPolicy
{
    public const string PolicyVersion = "hybrid-v1";

    private static readonly HashSet<string> SupportedDocumentStates =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "new",
            "existing-hwp",
            "existing-hwpx",
        };

    public static JsonObject Evaluate(JsonObject? args)
    {
        args ??= new JsonObject();
        var documentState = (Json.GetString(args, "documentState") ?? "new").Trim().ToLowerInvariant();
        if (!SupportedDocumentStates.Contains(documentState))
            return Json.ErrorResult(
                $"documentState는 new, existing-hwp, existing-hwpx 중 하나여야 합니다: {documentState}",
                "hwp");

        var hasExistingHwpTemplate = Json.GetBool(args, "hasExistingHwpTemplate");
        var requiresNativeFields = Json.GetBool(args, "requiresNativeFields");
        var requiresHwpOnlyObjects = Json.GetBool(args, "requiresHwpOnlyObjects");
        var requiresComplexMergedTables = Json.GetBool(args, "requiresComplexMergedTables");
        var mustPreserveOriginalLayout = Json.GetBool(args, "mustPreserveOriginalLayout");
        var docxGeneratorAvailable = !args.ContainsKey("docxGeneratorAvailable") ||
                                     Json.GetBool(args, "docxGeneratorAvailable");

        var reasons = new JsonArray();
        if (documentState is "existing-hwp" or "existing-hwpx")
            reasons.Add("기존 HWP/HWPX 문서는 원본 구조와 주변 서식을 보존해 직접 편집해야 합니다.");
        if (hasExistingHwpTemplate)
            reasons.Add("기존 한글 템플릿을 사용합니다.");
        if (requiresNativeFields)
            reasons.Add("한글 필드·누름틀 등 네이티브 필드가 필요합니다.");
        if (requiresHwpOnlyObjects)
            reasons.Add("한글 전용 개체가 필요합니다.");
        if (requiresComplexMergedTables)
            reasons.Add("복잡한 병합표를 한글 네이티브 구조로 유지해야 합니다.");
        if (mustPreserveOriginalLayout)
            reasons.Add("기존 한글 원본과 동일한 배치를 보존해야 합니다.");
        if (!docxGeneratorAvailable)
            reasons.Add("검수 가능한 DOCX 생성 도구를 사용할 수 없습니다.");

        var nativeHwp = documentState is "existing-hwp" or "existing-hwpx" ||
                        hasExistingHwpTemplate ||
                        requiresNativeFields ||
                        requiresHwpOnlyObjects ||
                        requiresComplexMergedTables ||
                        mustPreserveOriginalLayout ||
                        !docxGeneratorAvailable;

        if (!nativeHwp)
            reasons.Add("새 일반 문서는 문단·표·그림·단순 머리말/꼬리말을 DOCX에서 먼저 완성하는 편이 빠르고 안정적입니다.");

        var workflow = nativeHwp
            ? new JsonArray(
                "hwp_doctor와 hwp_get_active_context로 대상 문서를 확인합니다.",
                documentState == "new"
                    ? "hwp_launch에 creationMode=native-hwp, newDocument=true를 지정해 빈 문서를 한 번만 만듭니다."
                    : "기존 문서의 documentRef를 고정하고 HWP 직접 편집을 사용합니다.",
                "hwp_apply_ops의 dry-run, apply, postEditReread 순서를 지킵니다.",
                "표·쪽 수·핵심 문구를 다시 읽고 최종 PDF를 시각 검수합니다.")
            : new JsonArray(
                "DOCX에서 A4·표 너비·행 높이·글꼴·문단 간격을 완성합니다.",
                "DOCX를 PDF/PNG로 렌더하고 모든 쪽을 검수합니다.",
                "hwp_launch에 creationMode=docx-first, sourceFile, 새 outputFile, expectedPageCount, expectedTableCount, requiredText를 지정합니다.",
                "sourceUnchanged와 verification.passed를 확인하고 변환된 HWP/HWPX를 다시 읽습니다.",
                "최종 HWP PDF를 시각 검수한 뒤 필요한 미세 수정만 HWP 직접 편집으로 처리합니다.");

        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = "hwp",
            ["policyVersion"] = PolicyVersion,
            ["mode"] = nativeHwp ? "native-hwp" : "docx-first",
            ["wordComRequired"] = false,
            ["reasons"] = reasons,
            ["workflow"] = workflow,
            ["qualityGates"] = nativeHwp
                ? new JsonArray("dry-run/apply 일치", "postEditReread 검증", "표·쪽 수·핵심 문구 재확인", "최종 PDF 전쪽 시각 검수")
                : new JsonArray("DOCX 전쪽 렌더 검수", "원본 DOCX SHA-256 불변", "페이지 수 일치", "표 수 일치", "필수 문구 보존", "빈 경고", "최종 HWP PDF 전쪽 시각 검수"),
            ["fallback"] = nativeHwp
                ? "한글 전용 기능이 제거된 새 일반 문서로 요구사항이 바뀔 때만 DOCX 우선 경로를 다시 평가합니다."
                : "변환 품질 게이트가 실패하면 완료로 보고하지 말고 DOCX를 조정해 새 출력 이름으로 다시 변환합니다.",
        };
    }
}
