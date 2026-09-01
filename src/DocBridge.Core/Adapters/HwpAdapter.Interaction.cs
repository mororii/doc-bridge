using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class HwpAdapter
{
    private sealed record HwpViewState(
        int List, int Paragraph, int Position,
        bool HasSelection,
        int StartList, int StartParagraph, int StartPosition,
        int EndList, int EndParagraph, int EndPosition);

    private sealed class HwpInteractionState
    {
        private object? _application;
        private string? _originalDocumentId;
        private string? _targetDocumentId;
        private HwpViewState? _originalView;
        private HwpViewState? _targetView;
        private bool _restored;

        public bool InternalDocumentSwitched { get; private set; }
        public bool OriginalStateRestored { get; private set; } = true;

        public void CaptureOriginal(object application)
        {
            if (_originalDocumentId is not null) return;
            _application = application;
            dynamic hwp = application;
            var doc = ActiveDoc(hwp);
            if (doc is null) return;
            try { _originalDocumentId = Convert.ToString(doc.DocumentID); } catch { }
            _originalView = CaptureView(hwp);
        }

        public void CaptureTarget(object application)
        {
            _application = application;
            dynamic hwp = application;
            var doc = ActiveDoc(hwp);
            if (doc is null) return;
            try { _targetDocumentId = Convert.ToString(doc.DocumentID); } catch { }
            _targetView = CaptureView(hwp);
            InternalDocumentSwitched = !string.IsNullOrWhiteSpace(_originalDocumentId) &&
                !string.Equals(_originalDocumentId, _targetDocumentId, StringComparison.Ordinal);
        }

        public void Restore(ForegroundInteractionGuard? foreground = null)
        {
            if (_restored) return;
            _restored = true;
            if (_application is null) return;
            try
            {
                dynamic hwp = _application;
                if (!string.IsNullOrWhiteSpace(_targetDocumentId) &&
                    ActivateDocumentById(hwp, _targetDocumentId!))
                {
                    foreground?.Checkpoint(stopOnConcurrentInput: false);
                    var state = string.Equals(_targetDocumentId, _originalDocumentId, StringComparison.Ordinal)
                        ? _originalView ?? _targetView
                        : _targetView;
                    if (state is not null) OriginalStateRestored &= RestoreView(hwp, state);
                }

                if (InternalDocumentSwitched && !string.IsNullOrWhiteSpace(_originalDocumentId))
                {
                    OriginalStateRestored &= ActivateDocumentById(hwp, _originalDocumentId!);
                    foreground?.Checkpoint(stopOnConcurrentInput: false);
                    if (_originalView is not null) OriginalStateRestored &= RestoreView(hwp, _originalView);
                }
            }
            catch
            {
                OriginalStateRestored = false;
            }
        }

        private static HwpViewState? CaptureView(dynamic hwp)
        {
            try
            {
                int list = 0, paragraph = 0, position = 0;
                hwp.GetPos(out list, out paragraph, out position);
                int sl = 0, sp = 0, so = 0, el = 0, ep = 0, eo = 0;
                var hasSelection = false;
                try
                {
                    _ = hwp.GetSelectedPos(out sl, out sp, out so, out el, out ep, out eo);
                    hasSelection = sl != el || sp != ep || so != eo;
                }
                catch { }
                return new HwpViewState(list, paragraph, position, hasSelection,
                    sl, sp, so, el, ep, eo);
            }
            catch { return null; }
        }

        private static bool RestoreView(dynamic hwp, HwpViewState state)
        {
            try
            {
                try { hwp.HAction.Run("Cancel"); } catch { }
                if (state.HasSelection)
                {
                    hwp.SetPos(state.StartList, state.StartParagraph, state.StartPosition);
                    try
                    {
                        if (Convert.ToBoolean(hwp.SelectText(
                                state.StartParagraph, state.StartPosition,
                                state.EndParagraph, state.EndPosition)))
                            return true;
                    }
                    catch { }
                }
                hwp.SetPos(state.List, state.Paragraph, state.Position);
                return true;
            }
            catch { return false; }
        }

        private static bool ActivateDocumentById(dynamic hwp, string documentId)
        {
            try
            {
                dynamic documents = hwp.XHwpDocuments;
                var active = ActiveDoc(hwp);
                try
                {
                    if (active is not null && string.Equals(
                            Convert.ToString(active.DocumentID), documentId, StringComparison.Ordinal))
                        return true;
                }
                catch { }

                var count = Convert.ToInt32(documents.Count);
                for (var index = 0; index < count; index++)
                {
                    dynamic document = documents.Item(index);
                    try
                    {
                        if (!string.Equals(Convert.ToString(document.DocumentID), documentId,
                                StringComparison.Ordinal)) continue;
                        document.SetActive_XHwpDocument();
                        return true;
                    }
                    finally { RotHelper.ReleaseComObject((object)document); }
                }
            }
            catch { }
            return false;
        }
    }

    private static JsonObject CompleteHwpInteraction(
        ForegroundInteractionGuard foreground, HwpInteractionState documentState)
    {
        documentState.Restore(foreground);
        foreground.Checkpoint(stopOnConcurrentInput: false);
        var result = foreground.Complete();
        result["internalDocumentSwitched"] = documentState.InternalDocumentSwitched;
        result["originalStateRestored"] = documentState.OriginalStateRestored;
        return result;
    }

    private static void TrackHwpInteraction(
        object application, ForegroundInteractionGuard? foreground, HwpInteractionState? documentState,
        bool captureTarget)
    {
        var windowHandle = RotHelper.HwpWindowHandle(application);
        if (windowHandle != 0) foreground?.TrackTargetWindow(windowHandle);
        else foreground?.TrackTargetProcess(RotHelper.ProcessIdFromWindowHandle(windowHandle));
        documentState?.CaptureOriginal(application);
        if (captureTarget) documentState?.CaptureTarget(application);
    }
}
