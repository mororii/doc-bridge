# ExcelProductionWorkflowProbe

Independent production acceptance harness for the six finished Excel workbooks
plus pivot / protected-form / CSV entrypoints. Owns **new files in this folder
only**. Does not reference or compile `DocBridge.Core`.

```text
dotnet build -c Release
artifacts\bin\Release\net8.0-windows\DocBridge.ExcelProductionWorkflowProbe.exe --mode plan --scenario e1 --output <dir> --config <fixture-config.json>
```

Live requires `--transport mcp --mcp <already-built exe> --lease-ok` and `DOCBRIDGE_E2E=1`.
CLI live is refused. See `output/docbridge-production-20260910/verification/E1-CHECKPOINT.md`.
