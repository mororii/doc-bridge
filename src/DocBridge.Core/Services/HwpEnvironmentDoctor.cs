using Microsoft.Win32;
using System.Diagnostics;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// 한글 자동화 설치/TypeLib 상태를 COM 객체 생성 전에 검사한다. 잘못 등록된 구버전
/// HwpObject.tlb 때문에 빈 창이 반복 실행되는 문제를 사용자 작업 전에 진단한다.
/// </summary>
public static class HwpEnvironmentDoctor
{
    public const string TypeLibGuid = "{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}";
    public const string RecommendedHwp2024Version = "13.0.0.3870";
    private const string TypeLibSubKey = @"SOFTWARE\Classes\TypeLib\" + TypeLibGuid + @"\1.0\0\win32";
    private const string AppPathSubKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Hwp.exe";
    private static readonly object WorkingDirectoryGate = new();

    public static JsonObject Diagnose()
    {
        var installed = FindInstalledHwpExecutable();
        var automationWorkingDirectory = GetAutomationWorkingDirectory(installed);
        var registrations = ReadTypeLibRegistrations();
        var preferred = registrations.FirstOrDefault(item => item.Path is not null);
        var registrationState = EvaluateRegistration(installed, preferred?.Path);
        var uiFailure = HwpUiFailureDetector.Detect();
        var processWinDir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Process);
        var processSystemRoot = Environment.GetEnvironmentVariable("SystemRoot", EnvironmentVariableTarget.Process);
        var automationWindowsDirectory = GetAutomationWindowsDirectory();
        var automationEnvironmentRepairNeeded = NeedsProcessEnvironmentRepair(
            processWinDir, processSystemRoot, automationWindowsDirectory);
        var state = uiFailure is not null
            ? "HWP_UI_INITIALIZATION_FAILED"
            : automationWindowsDirectory is null
                ? "HWP_AUTOMATION_ENVIRONMENT_INVALID"
                : registrationState;
        var regArray = new JsonArray();
        foreach (var item in registrations)
            regArray.Add(new JsonObject
            {
                ["hive"] = item.Hive,
                ["view"] = item.View,
                ["path"] = item.Path,
                ["exists"] = item.Path is not null && File.Exists(item.Path),
            });

        var progIdRegistered = false;
        try { progIdRegistered = Type.GetTypeFromProgID("HWPFrame.HwpObject") is not null; } catch { }

        string? version = null;
        if (installed is not null)
        {
            try { version = FileVersionInfo.GetVersionInfo(installed).FileVersion; } catch { }
        }

        var updateRecommended = IsHwp2024VersionOlderThan(version, RecommendedHwp2024Version);
        var warnings = new JsonArray();
        if (updateRecommended)
            warnings.Add(
                $"한글 2024 {version ?? "unknown"}은 DocBridge 권장 패치 {RecommendedHwp2024Version}보다 오래되었습니다. " +
                "TourPopup/FontCache 초기화 안정성을 위해 한컴 자동 업데이트를 실행하세요.");
        if (automationEnvironmentRepairNeeded && automationWindowsDirectory is not null)
            warnings.Add(
                $"AI 클라이언트가 전달한 windir/SystemRoot가 비어 있거나 잘못되어 있습니다. " +
                $"DocBridge가 한글 시작 시 두 값을 {automationWindowsDirectory}(으)로 복구합니다.");

        var result = new JsonObject
        {
            ["ok"] = state == "CHECK_PASSED",
            ["app"] = "hwp",
            ["state"] = state,
            ["errorCode"] = state == "CHECK_PASSED" ? null : state,
            ["registrationState"] = registrationState,
            ["typeLibGuid"] = TypeLibGuid,
            ["progIdRegistered"] = progIdRegistered,
            ["installedExecutable"] = installed,
            ["installedVersion"] = version,
            ["automationWorkingDirectory"] = automationWorkingDirectory,
            ["workingDirectoryPolicy"] = "pin-installed-bin-during-com-activation",
            ["processWinDir"] = processWinDir,
            ["processSystemRoot"] = processSystemRoot,
            ["automationWindowsDirectory"] = automationWindowsDirectory,
            ["automationEnvironmentPolicy"] = "repair-windir-systemroot-and-pin-installed-bin-during-com-activation",
            ["automationEnvironmentRepairNeeded"] = automationEnvironmentRepairNeeded,
            ["recommendedHwp2024Version"] = RecommendedHwp2024Version,
            ["updateRecommended"] = updateRecommended,
            ["ownedAutomationBlocked"] = uiFailure is not null || automationWindowsDirectory is null,
            ["registrations"] = regArray,
            ["repairAvailable"] = installed is not null && state != "HWP_UI_INITIALIZATION_FAILED",
            ["restartRequiredAfterRepair"] = true,
            ["warnings"] = warnings,
            ["userAction"] = state switch
            {
                "CHECK_PASSED" when updateRecommended =>
                    $"자동화 환경과 등록은 정상입니다. 안정성을 위해 한글 2024를 {RecommendedHwp2024Version} 이상으로 업데이트하세요.",
                "CHECK_PASSED" => "조치가 필요하지 않습니다.",
                "HWP_UI_INITIALIZATION_FAILED" => HwpUiFailureDetector.UpdateAction,
                "HWP_AUTOMATION_ENVIRONMENT_INVALID" =>
                    "Windows 폴더를 확인할 수 없습니다. windir/SystemRoot 환경 변수와 Windows 설치 상태를 복구한 뒤 다시 실행하세요.",
                "HWP_NOT_INSTALLED" => "한컴 한글 2018 이상을 설치하세요.",
                "HWP_TYPELIB_NOT_REGISTERED" => "hwp_repair_typelib을 승인 실행한 뒤 한글과 AI 클라이언트를 다시 시작하세요.",
                "HWP_TYPELIB_VERSION_MISMATCH" => "설치된 한글 실행 파일로 TypeLib을 다시 등록한 뒤 한글과 AI 클라이언트를 다시 시작하세요.",
                _ => "한글 자동화 등록 상태를 확인하세요.",
            },
        };
        if (uiFailure is not null) result["uiFailure"] = HwpUiFailureDetector.ToJson(uiFailure);
        return result;
    }

    internal static bool IsHwp2024VersionOlderThan(string? version, string minimum)
    {
        if (!TryParseVersion(version, out var current) || !TryParseVersion(minimum, out var required))
            return false;
        return current.Major == 13 && current < required;
    }

    private static bool TryParseVersion(string? value, out Version version)
    {
        version = new Version();
        if (string.IsNullOrWhiteSpace(value)) return false;
        var numeric = value.Trim().Split(' ', '-', '+')[0];
        return Version.TryParse(numeric, out version!);
    }

    public static JsonObject Repair(string? explicitExecutable, bool elevate = true)
    {
        var before = Diagnose();
        var executable = string.IsNullOrWhiteSpace(explicitExecutable)
            ? Json.GetString(before, "installedExecutable")
            : Path.GetFullPath(explicitExecutable);
        if (string.IsNullOrWhiteSpace(executable) || !File.Exists(executable))
            return new JsonObject
            {
                ["ok"] = false,
                ["app"] = "hwp",
                ["errorCode"] = "HWP_EXECUTABLE_NOT_FOUND",
                ["errors"] = new JsonArray("TypeLib을 등록할 Hwp.exe를 찾지 못했습니다."),
                ["before"] = before,
            };

        try
        {
            var start = new ProcessStartInfo
            {
                FileName = executable,
                Arguments = "/RegServer",
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                Verb = elevate ? "runas" : "",
            };
            using var process = Process.Start(start)
                ?? throw new InvalidOperationException("Hwp.exe /RegServer 프로세스를 시작하지 못했습니다.");
            if (!process.WaitForExit(30000))
                throw new TimeoutException("Hwp.exe /RegServer가 30초 안에 끝나지 않았습니다.");

            JsonObject after = Diagnose();
            for (var attempt = 0; attempt < 10 && Json.GetString(after, "state") != "CHECK_PASSED"; attempt++)
            {
                Thread.Sleep(500);
                after = Diagnose();
            }
            var repaired = Json.GetString(after, "state") == "CHECK_PASSED";
            return new JsonObject
            {
                ["ok"] = repaired,
                ["app"] = "hwp",
                ["repairCommand"] = $"{executable} /RegServer",
                ["elevated"] = elevate,
                ["restartRequired"] = true,
                ["before"] = before,
                ["after"] = after,
                ["errors"] = repaired ? new JsonArray() : new JsonArray("TypeLib 재등록 후에도 검사가 통과하지 않았습니다."),
            };
        }
        catch (Exception ex)
        {
            return new JsonObject
            {
                ["ok"] = false,
                ["app"] = "hwp",
                ["errorCode"] = ex is System.ComponentModel.Win32Exception ? "HWP_TYPELIB_REPAIR_CANCELLED" : "HWP_TYPELIB_REPAIR_FAILED",
                ["errors"] = new JsonArray(ex.Message),
                ["before"] = before,
            };
        }
    }

    public static string EvaluateRegistration(string? installedExecutable, string? registeredTypeLib)
    {
        if (string.IsNullOrWhiteSpace(installedExecutable) || !File.Exists(installedExecutable))
            return "HWP_NOT_INSTALLED";
        if (string.IsNullOrWhiteSpace(registeredTypeLib) || !File.Exists(registeredTypeLib))
            return "HWP_TYPELIB_NOT_REGISTERED";

        var installedRoot = NormalizeDirectory(Path.GetDirectoryName(installedExecutable));
        var registeredRoot = NormalizeDirectory(Path.GetDirectoryName(registeredTypeLib));
        return string.Equals(installedRoot, registeredRoot, StringComparison.OrdinalIgnoreCase)
            ? "CHECK_PASSED"
            : "HWP_TYPELIB_VERSION_MISMATCH";
    }

    /// <summary>
    /// 한글 2024는 COM 서버로 시작될 때 호출자 작업 폴더를 상속한다. 설치 Bin이 아닌
    /// 폴더에서 시작하면 CultureFontManager가 사설 글꼴 경로를 잘못된 URI로 해석해
    /// MS.Internal.FontCache.Util 정적 초기화에서 종료될 수 있으므로 설치 Bin을 사용한다.
    /// </summary>
    public static string? GetAutomationWorkingDirectory(string? executable = null)
    {
        executable ??= FindInstalledHwpExecutable();
        if (string.IsNullOrWhiteSpace(executable) || !File.Exists(executable)) return null;
        var directory = Path.GetDirectoryName(Path.GetFullPath(executable));
        return !string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory)
            ? directory
            : null;
    }

    /// <summary>
    /// WPF FontCache.Util은 정적 초기화에서 process-level "windir" 값을 읽어
    /// 절대 Fonts URI를 만든다. 일부 AI/MCP 런처는 자식 환경을 축소해 이 값이
    /// 비어 있으므로, COM 서버 생성 전에 신뢰 가능한 Windows 경로를 복구한다.
    /// </summary>
    public static string? GetAutomationWindowsDirectory()
    {
        var candidates = new List<string?>
        {
            Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Process),
            Environment.GetEnvironmentVariable("SystemRoot", EnvironmentVariableTarget.Process),
        };
        if (OperatingSystem.IsWindows())
        {
            candidates.Add(Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine));
            candidates.Add(Environment.GetEnvironmentVariable("SystemRoot", EnvironmentVariableTarget.Machine));
            candidates.Add(Environment.GetFolderPath(Environment.SpecialFolder.Windows));
            try { candidates.Add(Directory.GetParent(Environment.SystemDirectory)?.FullName); } catch { }
        }
        return candidates.Select(NormalizeWindowsDirectory).FirstOrDefault(path => path is not null);
    }

    public static void ApplyAutomationEnvironment(ProcessStartInfo startInfo)
    {
        ArgumentNullException.ThrowIfNull(startInfo);
        var windowsDirectory = GetAutomationWindowsDirectory()
            ?? throw new InvalidOperationException(
                "windir/SystemRoot를 복구할 Windows 설치 폴더를 찾지 못했습니다.");
        startInfo.Environment["windir"] = windowsDirectory;
        startInfo.Environment["SystemRoot"] = windowsDirectory;
    }

    internal static bool NeedsProcessEnvironmentRepair(
        string? processWinDir, string? processSystemRoot, string? effectiveWindowsDirectory)
    {
        if (effectiveWindowsDirectory is null) return true;
        return !SameWindowsDirectory(processWinDir, effectiveWindowsDirectory) ||
               !SameWindowsDirectory(processSystemRoot, effectiveWindowsDirectory);
    }

    public static T RunWithAutomationWorkingDirectory<T>(Func<T> action)
    {
        ArgumentNullException.ThrowIfNull(action);
        var workingDirectory = GetAutomationWorkingDirectory();
        var windowsDirectory = GetAutomationWindowsDirectory();
        if (workingDirectory is null && windowsDirectory is null) return action();

        lock (WorkingDirectoryGate)
        {
            var previous = Environment.CurrentDirectory;
            var previousWinDir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Process);
            var previousSystemRoot = Environment.GetEnvironmentVariable("SystemRoot", EnvironmentVariableTarget.Process);
            try
            {
                if (workingDirectory is not null) Environment.CurrentDirectory = workingDirectory;
                if (windowsDirectory is not null)
                {
                    Environment.SetEnvironmentVariable("windir", windowsDirectory, EnvironmentVariableTarget.Process);
                    Environment.SetEnvironmentVariable("SystemRoot", windowsDirectory, EnvironmentVariableTarget.Process);
                }
                return action();
            }
            finally
            {
                Environment.CurrentDirectory = previous;
                Environment.SetEnvironmentVariable("windir", previousWinDir, EnvironmentVariableTarget.Process);
                Environment.SetEnvironmentVariable("SystemRoot", previousSystemRoot, EnvironmentVariableTarget.Process);
            }
        }
    }

    private static string? NormalizeWindowsDirectory(string? candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate)) return null;
        try
        {
            var full = Path.GetFullPath(candidate.Trim().Trim('"')).TrimEnd(Path.DirectorySeparatorChar);
            if (!Path.IsPathRooted(full) || !Directory.Exists(full)) return null;
            if (!Directory.Exists(Path.Combine(full, "Fonts")) ||
                !Directory.Exists(Path.Combine(full, "System32"))) return null;
            return full;
        }
        catch { return null; }
    }

    private static bool SameWindowsDirectory(string? candidate, string effective) =>
        string.Equals(NormalizeWindowsDirectory(candidate), effective, StringComparison.OrdinalIgnoreCase);

    private sealed record Registration(string Hive, string View, string? Path);

    private static List<Registration> ReadTypeLibRegistrations()
    {
        var result = new List<Registration>();
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
        {
            string? path = null;
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                using var key = baseKey.OpenSubKey(TypeLibSubKey, writable: false);
                path = key?.GetValue(null, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                if (!string.IsNullOrWhiteSpace(path)) path = Environment.ExpandEnvironmentVariables(path.Trim('"'));
            }
            catch { }
            result.Add(new Registration(hive.ToString(), view.ToString(), path));
        }
        return result;
    }

    public static string? FindInstalledHwpExecutable()
    {
        foreach (var process in Process.GetProcessesByName("Hwp"))
        {
            using (process)
            {
                try
                {
                    var path = process.MainModule?.FileName;
                    if (!string.IsNullOrWhiteSpace(path) && File.Exists(path)) return Path.GetFullPath(path);
                }
                catch { }
            }
        }

        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                using var key = baseKey.OpenSubKey(AppPathSubKey, writable: false);
                var path = key?.GetValue(null, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                if (!string.IsNullOrWhiteSpace(path))
                {
                    path = Environment.ExpandEnvironmentVariables(path.Trim('"'));
                    if (File.Exists(path)) return Path.GetFullPath(path);
                }
            }
            catch { }
        }

        var roots = new[]
        {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        };
        foreach (var root in roots.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            foreach (var vendor in new[] { "Hnc", "Hancom" })
            {
                var vendorRoot = Path.Combine(root, vendor);
                if (!Directory.Exists(vendorRoot)) continue;
                try
                {
                    var match = Directory.EnumerateFiles(vendorRoot, "Hwp.exe", SearchOption.AllDirectories)
                        .OrderByDescending(path => path, StringComparer.OrdinalIgnoreCase)
                        .FirstOrDefault();
                    if (match is not null) return Path.GetFullPath(match);
                }
                catch { }
            }
        }
        return null;
    }

    private static string NormalizeDirectory(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return "";
        return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}
