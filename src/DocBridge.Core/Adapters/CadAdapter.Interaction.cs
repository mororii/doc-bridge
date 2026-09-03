using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class CadAdapter
{
    private sealed class CadInteractionState
    {
        private object? _application;
        private object? _originalDocument;
        private string _originalDocumentIdentity = "";
        private string _originalLayout = "";
        private int? _originalActiveSpace;
        private bool? _originalMSpace;
        private double _originalViewCenterX;
        private double _originalViewCenterY;
        private double _originalViewSize;
        private bool _hasOriginalView;
        private bool _restored;

        public bool InternalDocumentSwitched { get; private set; }
        public bool OriginalStateRestored { get; private set; } = true;

        public void CaptureOriginal(object application)
        {
            if (_application is not null) return;
            _application = application;
            try
            {
                dynamic app = application;
                var document = ActiveDoc(app);
                if (document is null) return;
                _originalDocument = (object)document;
                _originalDocumentIdentity = CadDocumentIdentity(document);
                try { _originalLayout = Convert.ToString(document.ActiveLayout.Name, CultureInfo.InvariantCulture) ?? ""; }
                catch { }
                try { _originalActiveSpace = Convert.ToInt32(document.ActiveSpace, CultureInfo.InvariantCulture); }
                catch { }
                try { _originalMSpace = Convert.ToBoolean(document.MSpace, CultureInfo.InvariantCulture); }
                catch { }
                try
                {
                    var center = (Array)document.GetVariable("VIEWCTR");
                    _originalViewCenterX = Convert.ToDouble(center.GetValue(0), CultureInfo.InvariantCulture);
                    _originalViewCenterY = Convert.ToDouble(center.GetValue(1), CultureInfo.InvariantCulture);
                    _originalViewSize = Convert.ToDouble(document.GetVariable("VIEWSIZE"), CultureInfo.InvariantCulture);
                    _hasOriginalView = _originalViewSize > 0;
                }
                catch { }
            }
            catch { }
        }

        public void Restore()
        {
            if (_restored) return;
            _restored = true;
            if (_application is null || _originalDocument is null) return;
            try
            {
                dynamic app = _application;
                dynamic original = _originalDocument;
                var current = ActiveDoc(app);
                InternalDocumentSwitched = current is not null &&
                    !string.Equals(CadDocumentIdentity(current), _originalDocumentIdentity,
                        StringComparison.OrdinalIgnoreCase);

                if (InternalDocumentSwitched) original.Activate();
                if (!string.IsNullOrWhiteSpace(_originalLayout))
                {
                    try
                    {
                        if (!string.Equals((string)original.ActiveLayout.Name, _originalLayout, StringComparison.OrdinalIgnoreCase))
                            original.ActiveLayout = original.Layouts.Item(_originalLayout);
                    }
                    catch { OriginalStateRestored = false; }
                }
                if (_originalActiveSpace is not null)
                {
                    try
                    {
                        if (Convert.ToInt32(original.ActiveSpace) != _originalActiveSpace.Value)
                            original.ActiveSpace = _originalActiveSpace.Value;
                    }
                    catch { OriginalStateRestored = false; }
                }
                if (_originalMSpace is not null)
                {
                    try
                    {
                        if (Convert.ToBoolean(original.MSpace) != _originalMSpace.Value)
                            original.MSpace = _originalMSpace.Value;
                    }
                    catch { /* MSpace is unavailable in model space on some versions. */ }
                }
                if (_hasOriginalView)
                {
                    try
                    {
                        var center = (Array)original.GetVariable("VIEWCTR");
                        var x = Convert.ToDouble(center.GetValue(0), CultureInfo.InvariantCulture);
                        var y = Convert.ToDouble(center.GetValue(1), CultureInfo.InvariantCulture);
                        var size = Convert.ToDouble(original.GetVariable("VIEWSIZE"), CultureInfo.InvariantCulture);
                        if (Math.Abs(x - _originalViewCenterX) > 1e-8 ||
                            Math.Abs(y - _originalViewCenterY) > 1e-8 || Math.Abs(size - _originalViewSize) > 1e-8)
                            app.ZoomCenter(Point(_originalViewCenterX, _originalViewCenterY, 0), _originalViewSize);
                    }
                    catch { OriginalStateRestored = false; }
                }

                current = ActiveDoc(app);
                OriginalStateRestored &= current is not null &&
                    string.Equals(CadDocumentIdentity(current), _originalDocumentIdentity,
                        StringComparison.OrdinalIgnoreCase);
            }
            catch { OriginalStateRestored = false; }
        }
    }

    private static string CadDocumentIdentity(dynamic document)
    {
        try
        {
            var fullName = Convert.ToString(document.FullName, CultureInfo.InvariantCulture) ?? "";
            if (!string.IsNullOrWhiteSpace(fullName)) return fullName;
        }
        catch { }
        try { return Convert.ToString(document.Name, CultureInfo.InvariantCulture) ?? ""; }
        catch { return ""; }
    }

    private static long CadWindowHandle(object application)
    {
        try { return Convert.ToInt64(((dynamic)application).HWND, CultureInfo.InvariantCulture); }
        catch
        {
            try { return Convert.ToInt64(((dynamic)application).Hwnd, CultureInfo.InvariantCulture); }
            catch { return 0; }
        }
    }

    private static void TrackCadInteraction(
        object application, ForegroundInteractionGuard foreground, CadInteractionState? state)
    {
        var handle = CadWindowHandle(application);
        if (handle != 0) foreground.TrackTargetWindow(handle);
        else foreground.TrackTargetProcess(RotHelper.ProcessIdFromWindowHandle(handle));
        state?.CaptureOriginal(application);
    }

    private static JsonObject CompleteCadInteraction(
        ForegroundInteractionGuard foreground, CadInteractionState state)
    {
        state.Restore();
        foreground.Checkpoint(stopOnConcurrentInput: false);
        var result = foreground.Complete();
        result["internalDocumentSwitched"] = state.InternalDocumentSwitched;
        result["originalStateRestored"] = state.OriginalStateRestored;
        return result;
    }
}
