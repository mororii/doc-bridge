using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class SnapshotLookupTests : IDisposable
{
    private readonly TestHome _home = new();
    public void Dispose() => _home.Dispose();

    [Fact]
    public void Get_probes_exact_id_and_does_not_parse_sibling_history()
    {
        var svc = new SnapshotService(_home.Options);
        SnapshotInfo? target = null;
        for (var i = 0; i < 40; i++)
        {
            var info = CreatePlain(svc, "fake", $"fake://doc-{i}");
            if (i == 7) target = info;
        }

        CreatePlain(svc, "excel", "excel://other");

        var before = svc.MetadataFilesRead;
        var got = svc.Get(target!.SnapshotId);
        Assert.NotNull(got);
        Assert.Equal(target.SnapshotId, got!.Value.Info.SnapshotId);
        Assert.Equal("fake", got.Value.Info.App);
        Assert.Equal(1, svc.MetadataFilesRead - before);

        before = svc.MetadataFilesRead;
        Assert.Null(svc.Get("no-such-snapshot-id"));
        Assert.Equal(0, svc.MetadataFilesRead - before);
    }

    [Fact]
    public void Get_with_known_app_does_not_open_other_app_metadata()
    {
        var svc = new SnapshotService(_home.Options);
        var fake = CreatePlain(svc, "fake", "fake://doc");
        var excel = CreatePlain(svc, "excel", "excel://doc");

        var before = svc.MetadataFilesRead;
        var got = svc.Get(fake.SnapshotId, "fake");
        Assert.NotNull(got);
        Assert.Equal(fake.SnapshotId, got!.Value.Info.SnapshotId);
        Assert.Equal(1, svc.MetadataFilesRead - before);

        before = svc.MetadataFilesRead;
        Assert.Null(svc.Get(fake.SnapshotId, "excel"));
        Assert.Equal(0, svc.MetadataFilesRead - before);

        before = svc.MetadataFilesRead;
        Assert.NotNull(svc.Get(excel.SnapshotId, "excel"));
        Assert.Equal(1, svc.MetadataFilesRead - before);
    }

    [Fact]
    public void Get_rejects_unsafe_path_components()
    {
        var svc = new SnapshotService(_home.Options);
        CreatePlain(svc, "fake", "fake://doc");

        var before = svc.MetadataFilesRead;
        Assert.Null(svc.Get(".."));
        Assert.Null(svc.Get("../secret"));
        Assert.Null(svc.Get("fake/../excel"));
        Assert.Null(svc.Get("a\\b"));
        Assert.Null(svc.Get("id:evil"));
        Assert.Null(svc.Get("ok-id", ".."));
        Assert.Null(svc.Get("ok-id", "fake/../excel"));
        Assert.Equal(0, svc.MetadataFilesRead - before);
        Assert.False(SnapshotService.IsSafeSnapshotId(".."));
        Assert.False(SnapshotService.IsSafeSnapshotId("a/b"));
        Assert.False(SnapshotService.IsSafeSnapshotId("app:name"));
    }

    [Fact]
    public void Get_fails_closed_when_metadata_id_or_app_mismatches_folder()
    {
        var svc = new SnapshotService(_home.Options);
        var info = CreatePlain(svc, "fake", "fake://doc");

        RewriteMetadata(info.Dir, meta =>
        {
            meta["snapshotId"] = "not-this-folder";
            meta["app"] = "fake";
        });
        var before = svc.MetadataFilesRead;
        Assert.Null(svc.Get(info.SnapshotId));
        Assert.Equal(1, svc.MetadataFilesRead - before);

        RewriteMetadata(info.Dir, meta =>
        {
            meta["snapshotId"] = info.SnapshotId;
            meta["app"] = "excel";
        });
        Assert.Null(svc.Get(info.SnapshotId));
        Assert.Null(svc.Get(info.SnapshotId, "fake"));
    }

    [Fact]
    public void Legacy_snapshot_without_reuse_keys_still_gets_by_id()
    {
        var svc = new SnapshotService(_home.Options);
        var info = CreatePlain(svc, "fake", "legacy://doc");
        var metaPath = Path.Combine(info.Dir, SnapshotService.MetadataFile);
        var meta = JsonNode.Parse(File.ReadAllText(metaPath))!.AsObject();
        Assert.Null(Json.GetInt(meta, "snapshotReuseVersion"));
        Assert.Null(Json.GetString(meta, "opsHash"));

        var got = svc.Get(info.SnapshotId);
        Assert.NotNull(got);
        Assert.Equal(info.SnapshotId, got!.Value.Info.SnapshotId);
        Assert.Equal("fake", got.Value.Info.App);
        Assert.Equal("legacy://doc", got.Value.Info.DocumentRef);
    }

    [Fact]
    public void Reusable_lookup_reads_only_the_newest_directory_window()
    {
        var svc = new SnapshotService(_home.Options);
        const string opsHash = "ops-window";
        var preview = PreviewA1();
        SnapshotInfo? hiddenMatch = null;
        for (var i = 0; i < 8; i++)
            hiddenMatch = Plant("fake", $"20260101-000000-{i:x8}", "fake://doc", opsHash, preview);
        for (var i = 0; i < 20; i++)
            Plant("fake", $"20261231-000000-{i:x8}", "fake://doc", "other-hash", preview);

        var before = svc.MetadataFilesRead;
        Assert.Null(svc.FindLatestReusableCandidate(
            "fake", "fake://doc", opsHash, SameRef, searchLimit: 20));
        Assert.InRange(svc.MetadataFilesRead - before, 1, 20);

        var found = svc.FindLatestReusableCandidate(
            "fake", "fake://doc", opsHash, SameRef, searchLimit: 40);
        Assert.NotNull(found);
        Assert.Equal(hiddenMatch!.SnapshotId, found!.Value.Info.SnapshotId);
    }

    [Fact]
    public void Get_accepts_legacy_metadata_when_id_and_app_fields_are_absent()
    {
        var svc = new SnapshotService(_home.Options);
        var info = CreatePlain(svc, "fake", "legacy://missing-fields");
        RewriteMetadata(info.Dir, meta =>
        {
            meta.Remove("snapshotId");
            meta.Remove("app");
        });

        var got = svc.Get(info.SnapshotId, "fake");
        Assert.NotNull(got);
        Assert.Equal(info.SnapshotId, got!.Value.Info.SnapshotId);
        Assert.Equal("fake", got.Value.Info.App);
    }

    [Fact]
    public void Reusable_lookup_counts_malformed_metadata_against_search_limit()
    {
        var svc = new SnapshotService(_home.Options);
        var preview = PreviewA1();
        Plant("fake", "20260101-000000-00000001", "fake://doc", "ops-hidden", preview);
        for (var i = 0; i < 25; i++)
        {
            var bad = Plant("fake", $"20261231-000000-{i:x8}", "fake://doc", "ops-hidden", preview);
            RewriteMetadata(bad.Dir, meta => meta["snapshotId"] = "not-the-folder");
        }

        var before = svc.MetadataFilesRead;
        Assert.Null(svc.FindLatestReusableCandidate(
            "fake", "fake://doc", "ops-hidden", SameRef, searchLimit: 20));
        Assert.Equal(20, svc.MetadataFilesRead - before);
    }

    [Fact]
    public void Reusable_lookup_skips_identity_mismatch_inside_the_window()
    {
        var svc = new SnapshotService(_home.Options);
        var preview = PreviewA1();
        var match = CreateReusable(svc, "fake", "fake://doc", "ops-ok", preview);
        RewriteMetadata(match.Dir, meta => meta["app"] = "excel");

        var before = svc.MetadataFilesRead;
        Assert.Null(svc.FindLatestReusableCandidate(
            "fake", "fake://doc", "ops-ok", SameRef, searchLimit: 5));
        Assert.Equal(1, svc.MetadataFilesRead - before);
    }

    [Fact]
    public void Reusable_lookup_rereads_directory_metadata_and_does_not_cache()
    {
        var svc = new SnapshotService(_home.Options);
        var preview = PreviewA1();
        var match = CreateReusable(svc, "fake", "fake://doc", "ops-live", preview);

        var first = svc.FindLatestReusableCandidate("fake", "fake://doc", "ops-live", SameRef);
        Assert.NotNull(first);
        Assert.Equal(match.SnapshotId, first!.Value.Info.SnapshotId);

        RewriteMetadata(match.Dir, meta => meta["opsHash"] = "ops-changed");

        var second = svc.FindLatestReusableCandidate("fake", "fake://doc", "ops-live", SameRef);
        Assert.Null(second);
    }

    [Fact]
    public void Reusable_lookup_rejects_unsafe_app_without_reading_metadata()
    {
        var svc = new SnapshotService(_home.Options);
        CreateReusable(svc, "fake", "fake://doc", "ops-live", PreviewA1());
        var before = svc.MetadataFilesRead;
        Assert.Null(svc.FindLatestReusableCandidate("..", "fake://doc", "ops-live", SameRef));
        Assert.Null(svc.FindLatestReusableCandidate("fake/../fake", "fake://doc", "ops-live", SameRef));
        Assert.Equal(0, svc.MetadataFilesRead - before);
    }

    private SnapshotInfo Plant(
        string app, string id, string documentRef, string? opsHash, ApplyPreview? preview)
    {
        var dir = Path.Combine(_home.Options.SnapshotsDir, app, id);
        Directory.CreateDirectory(dir);
        var meta = new JsonObject
        {
            ["snapshotId"] = id,
            ["createdAt"] = "2026-01-01T00:00:00+00:00",
            ["app"] = app,
            ["documentRef"] = documentRef,
            ["reason"] = "planted",
        };
        File.WriteAllText(Path.Combine(dir, "state.json"), "{}");
        if (opsHash is not null && preview is not null)
        {
            meta["snapshotReuseVersion"] = 1;
            ApplyPreviewArtifact.StoreInMetadata(meta, opsHash, preview);
        }

        File.WriteAllText(Path.Combine(dir, SnapshotService.MetadataFile), meta.ToJsonString());
        return new SnapshotInfo(id, "2026-01-01T00:00:00+00:00", app, documentRef, "planted", dir);
    }

    private static SnapshotInfo CreatePlain(SnapshotService svc, string app, string documentRef) =>
        svc.Create(app, "unit-test", documentRef,
            (dir, _) => File.WriteAllText(Path.Combine(dir, "state.json"), "{}"));

    private static SnapshotInfo CreateReusable(
        SnapshotService svc, string app, string documentRef, string opsHash, ApplyPreview preview) =>
        svc.Create(app, "dry-run", documentRef, (dir, meta) =>
        {
            File.WriteAllText(Path.Combine(dir, "state.json"), "{}");
            meta["snapshotReuseVersion"] = 1;
            ApplyPreviewArtifact.StoreInMetadata(meta, opsHash, preview);
        });

    private static ApplyPreview PreviewA1()
    {
        var preview = new ApplyPreview();
        preview.Affected.Add(new AffectedRef("cell", "A1"));
        return preview;
    }

    private static bool SameRef(string? expected, string? current) =>
        string.Equals(expected, current, StringComparison.Ordinal);

    private static void RewriteMetadata(string snapshotDir, Action<JsonObject> edit)
    {
        var path = Path.Combine(snapshotDir, SnapshotService.MetadataFile);
        var meta = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        edit(meta);
        File.WriteAllText(path, meta.ToJsonString());
    }
}
