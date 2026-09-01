using Xunit;

// DOCBRIDGE_HOME 환경변수를 공유하는 파일 기반 서비스 테스트이므로
// 어셈블리 내 테스트 병렬 실행을 끈다 (프로세스 전역 env 경합 방지).
[assembly: CollectionBehavior(DisableTestParallelization = true)]
