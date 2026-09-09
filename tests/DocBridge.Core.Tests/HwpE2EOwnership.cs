using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;

namespace DocBridge.Core.Tests;

/// <summary>
/// Test-only HWP E2E COM helpers. Product source is not referenced for teardown
/// ownership: this type never calls process.Kill or XHwpWindows.Close.
/// </summary>
internal static class HwpE2EOwnership
{
    internal static HashSet<int> CurrentHwpProcessIds()
    {
        var ids = new HashSet<int>();
        foreach (var process in Process.GetProcessesByName("Hwp"))
        {
            using (process)
            {
                try { ids.Add(process.Id); } catch { }
            }
        }
        return ids;
    }

    internal static bool TryFileNew(dynamic hwp)
    {
        try { return CoerceActionResult(hwp.HAction.Run("FileNew")); }
        catch { return false; }
    }

    internal static bool? TryIsEmpty(object hwp)
    {
        try { return (bool)((dynamic)hwp).IsEmpty; }
        catch { return null; }
    }

    internal static bool TryReadInventory(dynamic hwp, out HwpE2EDocumentSnapshot snapshot)
    {
        snapshot = HwpE2EDocumentSnapshot.Empty;
        object? documentsObject = null;
        try
        {
            documentsObject = (object)hwp.XHwpDocuments;
            dynamic documents = documentsObject;
            var count = Convert.ToInt32(documents.Count);
            var ids = new List<string?>(count);
            for (var index = 0; index < count; index++)
            {
                object? document = null;
                try
                {
                    document = (object)documents.Item(index);
                    try { ids.Add(Convert.ToString(((dynamic)document).DocumentID)); }
                    catch { ids.Add(null); }
                }
                catch
                {
                    ids.Add(null);
                }
                finally { ReleaseComReference(document); }
            }

            string? activeId = null;
            object? activeObject = null;
            try
            {
                activeObject = (object)documents.Active_XHwpDocument;
                activeId = Convert.ToString(((dynamic)activeObject).DocumentID);
            }
            catch { activeId = null; }
            finally { ReleaseComReference(activeObject); }

            bool? isEmpty = TryIsEmpty((object)hwp);
            string unusedError;
            return HwpE2EOwnershipPolicy.TryValidateInventory(
                count, ids, activeId, isEmpty, out snapshot, out unusedError);
        }
        catch
        {
            snapshot = HwpE2EDocumentSnapshot.Empty;
            return false;
        }
        finally { ReleaseComReference(documentsObject); }
    }

    internal static void InsertSeedText(dynamic hwp, string text)
    {
        object? actionObject = null;
        object? parameterObject = null;
        object? hSetObject = null;
        try
        {
            dynamic act = hwp.HAction;
            dynamic ps = hwp.HParameterSet.HInsertText;
            dynamic hSet = ps.HSet;
            actionObject = (object)act;
            parameterObject = (object)ps;
            hSetObject = (object)hSet;
            act.GetDefault("InsertText", hSet);
            ps.Text = text;
            if (!CoerceActionResult(act.Execute("InsertText", hSet)))
                throw new InvalidOperationException("HWP E2E seed InsertText failed on a proven owned tab.");
        }
        finally
        {
            ReleaseComReference(hSetObject);
            ReleaseComReference(parameterObject);
            ReleaseComReference(actionObject);
        }
    }

    internal static bool TryCloseDocumentById(dynamic hwp, string documentId)
    {
        if (string.IsNullOrWhiteSpace(documentId)) return false;
        if (!TryActivateDocumentById(hwp, documentId)) return false;
        try { return CoerceActionResult(hwp.HAction.Run("FileClose")); }
        catch { return false; }
    }

    internal static bool TrySaveDocumentById(dynamic hwp, string documentId, string path)
    {
        if (string.IsNullOrWhiteSpace(documentId) || !TryActivateDocumentById(hwp, documentId))
            return false;
        return TrySaveActiveDocument(hwp, path);
    }

    internal static bool TrySaveActiveDocument(dynamic hwp, string path)
    {
        object? actionObject = null;
        object? parameterObject = null;
        object? hSetObject = null;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            dynamic act = hwp.HAction;
            dynamic sa = hwp.HParameterSet.HFileSaveAs;
            dynamic hSet = sa.HSet;
            actionObject = (object)act;
            parameterObject = (object)sa;
            hSetObject = (object)hSet;
            act.GetDefault("FileSaveAs", hSet);
            sa.SaveFileName = path;
            try { sa.SaveFormat = "HWP"; } catch { }
            try { sa.SaveOverWrite = true; } catch { }
            return CoerceActionResult(act.Execute("FileSaveAs", hSet));
        }
        catch { return false; }
        finally
        {
            ReleaseComReference(hSetObject);
            ReleaseComReference(parameterObject);
            ReleaseComReference(actionObject);
        }
    }

    internal static IReadOnlyList<string> CopyArtifacts(string artifactDir, IEnumerable<string> sources)
    {
        Directory.CreateDirectory(artifactDir);
        var copied = new List<string>();
        foreach (var source in sources)
        {
            if (string.IsNullOrWhiteSpace(source) || !File.Exists(source)) continue;
            var dest = Path.Combine(
                artifactDir,
                $"{Path.GetFileNameWithoutExtension(source)}-{DateTime.UtcNow:yyyyMMddTHHmmssfff}-{Guid.NewGuid():N}{Path.GetExtension(source)}");
            File.Copy(source, dest, overwrite: false);
            copied.Add(dest);
            Console.Error.WriteLine("HwpE2E: retained artifact: " + dest);
        }
        return copied;
    }

    /// <summary>
    /// Product <see cref="HwpAdapter.Dispose"/> Clears, XHwpWindows.Close, and Kill when
    /// the injected factory path set _ownsAttached. Tests must neutralize that before Dispose.
    /// </summary>
    internal static bool TryNeutralizeProductTeardown(HwpAdapter adapter)
    {
        try
        {
            const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic;
            var type = typeof(HwpAdapter);
            var owns = type.GetField("_ownsAttached", flags);
            var pid = type.GetField("_ownedProcessId", flags);
            if (owns is null || pid is null) return false;
            owns.SetValue(adapter, false);
            pid.SetValue(adapter, 0);
            return true;
        }
        catch { return false; }
    }

    internal static void DisposeAdapterWithoutKilling(HwpAdapter? adapter)
    {
        if (adapter is null) return;
        if (!TryNeutralizeProductTeardown(adapter))
        {
            Console.Error.WriteLine(
                "HwpE2E: skipped HwpAdapter.Dispose because product teardown flags could not be neutralized (would Kill/Close-all).");
            return;
        }

        try { adapter.Dispose(); }
        catch (Exception ex)
        {
            Console.Error.WriteLine("HwpE2E: neutralized adapter Dispose reported: " + ex.Message);
        }
    }

    /// <summary>
    /// One ReleaseComObject per acquired RCW. Never FinalReleaseComObject on HWP
    /// document/action objects that may still be cached by the application.
    /// </summary>
    internal static void ReleaseComReference(object? value)
    {
        if (value is null) return;
        try
        {
            if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
        catch { }
    }

    private static bool TryActivateDocumentById(dynamic hwp, string documentId)
    {
        object? documentsObject = null;
        try
        {
            documentsObject = (object)hwp.XHwpDocuments;
            dynamic documents = documentsObject;
            if (ActiveDocumentIdEquals(documents, documentId))
                return true;

            var count = Convert.ToInt32(documents.Count);
            var activated = false;
            for (var index = 0; index < count; index++)
            {
                object? document = null;
                try
                {
                    document = (object)documents.Item(index);
                    dynamic item = document;
                    if (!string.Equals(Convert.ToString(item.DocumentID), documentId, StringComparison.Ordinal))
                        continue;
                    item.SetActive_XHwpDocument();
                    activated = true;
                    break;
                }
                catch { }
                finally { ReleaseComReference(document); }
            }

            return activated && ActiveDocumentIdEquals(documents, documentId);
        }
        catch { return false; }
        finally { ReleaseComReference(documentsObject); }
    }

    private static bool ActiveDocumentIdEquals(dynamic documents, string documentId)
    {
        object? activeObject = null;
        try
        {
            activeObject = (object)documents.Active_XHwpDocument;
            return string.Equals(
                Convert.ToString(((dynamic)activeObject).DocumentID),
                documentId,
                StringComparison.Ordinal);
        }
        catch { return false; }
        finally { ReleaseComReference(activeObject); }
    }

    private static bool CoerceActionResult(object? result) => result switch
    {
        bool value => value,
        byte value => value != 0,
        short value => value != 0,
        int value => value != 0,
        long value => value != 0,
        null => false,
        _ => Convert.ToBoolean(result),
    };
}

/// <summary>
/// Forwards host contracts but never forwards <see cref="HwpAdapter.Dispose"/> while
/// product teardown still Kills or blanket-closes windows.
/// </summary>
internal sealed class HwpE2ESafeAdapter : IAppAdapter, IHwpAutomationAdapter, IPreviewReuseAdapter
{
    private readonly HwpAdapter _inner;
    private readonly Action _beforeInnerDispose;

    public HwpE2ESafeAdapter(HwpAdapter inner, Action beforeInnerDispose)
    {
        _inner = inner;
        _beforeInnerDispose = beforeInnerDispose;
    }

    public string App => _inner.App;
    public AdapterStatus GetStatus() => _inner.GetStatus();
    public JsonObject GetCapabilities() => _inner.GetCapabilities();
    public ContextResult GetActiveContext() => _inner.GetActiveContext();
    public JsonObject Read(JsonObject args) => _inner.Read(args);
    public ApplyPreview Preview(IReadOnlyList<JsonObject> ops) => _inner.Preview(ops);
    public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId) => _inner.Apply(ops, snapshotId);
    public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
        => _inner.CaptureSnapshot(snapshotDir, metadata, ops);
    public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata)
        => _inner.RestoreSnapshot(snapshotDir, metadata);
    public JsonObject Launch(JsonObject args) => _inner.Launch(args);
    public JsonObject Doctor(JsonObject args) => _inner.Doctor(args);
    public JsonObject RepairTypeLib(JsonObject args) => _inner.RepairTypeLib(args);
    public JsonObject ValidatePreviewReuse(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
        => _inner.ValidatePreviewReuse(snapshotDir, metadata, ops);

    public void Dispose()
    {
        try { _beforeInnerDispose(); }
        catch (Exception ex)
        {
            Console.Error.WriteLine("HwpE2E: owned-session release reported: " + ex.Message);
        }

        HwpE2EOwnership.DisposeAdapterWithoutKilling(_inner);
    }
}
