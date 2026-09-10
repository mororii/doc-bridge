# ExcelSnapshotStrategyProbe

Guarded .NET 8 experiment. Not a product binary. Requires `DOCBRIDGE_E2E=1`.

This directory is the durable contract for the probe. Dated folders under `output/` are run artifacts, not the interface spec.

```text
DocBridge.ExcelSnapshotStrategyProbe.exe --output <dir> --scale 100|1000|7000 --repeat 1..3 [--case mixed-fill|uniform-bold] [--deferred-restore]
```

`--case` defaults to `mixed-fill` and requests `fillColor` so capture hits the coupled-color path. `uniform-bold` is an explicit Bold-only control. `--output` must not be a user Documents path. Creates a new owned Excel PID only. Does not set DisplayAlerts. Never Kill.

`--deferred-restore` is optional and leaves the default capture-benchmark path unchanged. It is a fixture-only SaveCopyAs checkpoint + deferred production-native extract/restore proof for this generated `.xlsx` (scale 100|1000). Not a generic rollout. Result JSON then labels the run as a local experimental probe implementation, not a published v0.4.20 public-baseline package.

Audit expectation `value.kind` is the COM `Value2` CLR type (`string|number|bool|blank`). `formula` is JSON null only when `HasFormula` is false. The C4 dry-run is unknown-style-key rejection (`fillTint`), not a native RGB+tint snapshot proof.

## Artifacts written under `--output`

| Role | File name |
|---|---|
| Metrics + identity + HRESULT + cleanup/PID flags | `excel-snapshot-strategy-probe-result.json` |
| Owned saved baseline (last SaveAs) | `excel-snapshot-strategy-target.xlsx` |
| Independent disk copy of that baseline | `excel-snapshot-strategy-baseline-copy.xlsx` |
| Seeded semantic catalog (not native-before) | `excel-snapshot-strategy-expected.json` |
| Per-cell native-before observations (pre-SaveCopyAs) | `excel-snapshot-strategy-native-before.json` |
| Claude audit expectation `docbridge-excel-audit-expectation/1` | `excel-snapshot-strategy-audit-expectation.json` |
| In-memory SaveCopyAs (unique per capture repeat) | `savecopyas-rep{n}-{guid}.xlsx` |
| Deferred checkpoint / corrupt copies (`--deferred-restore`) | `deferred-checkpoint-{guid}.xlsx`, `deferred-corrupt-{guid}.xlsx` |

Scoped production captures go to `probe-home/captures/scoped-rep{n}-{guid}/` (path also on stdout as `INTERFACE scopedCaptureDir=`). Host apply/restore smoke snapshots go to `probe-home/snapshots/excel/`.

## Timing boundaries

- Scoped capture ms = `ExcelAdapter.CaptureSnapshot` only. Workbook-backup source is read from capture metadata (`current-memory-savecopyas` vs `last-saved-file`), not hardcoded.
- SaveCopyAs ms = current memory to a unique file. Must not change FullName / Saved / workbook count. Not operation-scoped rollback.
- Checkpoint total (`--deferred-restore`) splits `copyMs` / `sha256Ms` / `ooxmlPreflightMs` / `identityMs` / `checkpointTotalMs`. `copyMs` is not the full checkpoint.
- Repeat 2 uses `savecopyas-then-scoped`; other capture repeats use `scoped-then-savecopyas`.
- Default path: one-cell A3 production apply/restore is baseline smoke only. Candidate native-checkpoint restore is unimplemented unless `--deferred-restore` is requested.
- `--deferred-restore`: after eligibility + checkpoint preflight, `ExcelAdapter.Apply` then deferred native extract/restore. That local experimental path is what `claims.candidateRestoreImplemented` reports.

Built exe (Release, `-warnaserror`): `bin/Release/net8.0-windows/DocBridge.ExcelSnapshotStrategyProbe.exe`
