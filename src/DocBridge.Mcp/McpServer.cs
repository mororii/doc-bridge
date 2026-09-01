using DocBridge.Core.Services;

namespace DocBridge.Mcp;

/// <summary>
/// MCP stdio 서버 루프. NDJSON (한 줄 = JSON-RPC 메시지 하나).
/// stdout에는 오직 프로토콜 메시지만 쓴다 (보안 원칙/명령서 §10: stdout 로그 금지).
/// 모든 로그는 stderr로만 보낸다.
///
/// 한 줄 처리 중 예외가 나도 루프를 죽이지 않는다 — 클라이언트(Claude/Codex/Kimi)는
/// 서버 프로세스 종료를 곧바로 연결 실패로 처리하기 때문이다.
/// </summary>
public sealed class McpServer
{
    private readonly JsonRpcDispatcher _dispatcher;
    private readonly TextReader _input;
    private readonly TextWriter _output;
    private readonly TextWriter _log;

    public McpServer(JsonRpcDispatcher dispatcher, TextReader input, TextWriter output, TextWriter? log = null)
    {
        _dispatcher = dispatcher;
        _input = input;
        _output = output;
        _log = log ?? TextWriter.Null;
    }

    public async Task RunAsync(CancellationToken ct = default)
    {
        await _log.WriteLineAsync("[doc-bridge] stdio server ready");
        while (!ct.IsCancellationRequested)
        {
            string? line;
            try
            {
                line = await _input.ReadLineAsync(ct);
            }
            catch (OperationCanceledException) { break; }
            catch (IOException) { break; }        // 입력 파이프 종료

            if (line is null) break;              // 클라이언트 종료 (EOF)
            // 일부 Windows PowerShell/클라이언트는 리디렉션된 stdin 첫 줄에 UTF-8 BOM을
            // 문자(U+FEFF)로 남긴다. JSON 앞의 BOM만 제거해 initialize parse error를 막는다.
            line = line.Trim().TrimStart('\uFEFF');
            if (line.Length == 0) continue;

            string payload;
            try
            {
                var response = _dispatcher.DispatchLine(line);
                if (response is null) continue;   // notification — 응답 없음
                payload = Json.ToCompact(response);
            }
            catch (Exception ex)
            {
                // 디스패처가 삼키지 못한 예외도 프로토콜 오류로 바꿔 연결을 유지한다.
                await _log.WriteLineAsync($"[doc-bridge] dispatch failure: {ex}");
                payload = """{"jsonrpc":"2.0","id":null,"error":{"code":-32603,"message":"internal error"}}""";
            }

            try
            {
                await _output.WriteLineAsync(payload);
                await _output.FlushAsync();
            }
            catch (IOException) { break; }        // stdout 파이프 종료
        }
    }
}
