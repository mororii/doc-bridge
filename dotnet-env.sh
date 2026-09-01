# Git Bash 환경에 누락된 Windows 표준 환경변수를 env로 주입해 dotnet 실행
dn() {
  env OS=Windows_NT \
      ProgramFiles='C:\Program Files' \
      'ProgramFiles(x86)=C:\Program Files (x86)' \
      ProgramData='C:\ProgramData' \
      CommonProgramFiles='C:\Program Files\Common Files' \
      'CommonProgramFiles(x86)=C:\Program Files (x86)\Common Files' \
      ComSpec='C:\Windows\System32\cmd.exe' \
      DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_NOLOGO=1 \
      "/c/Program Files/dotnet/dotnet.exe" "$@"
}
