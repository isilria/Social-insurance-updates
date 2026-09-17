# 사회보험 재원별 대사 보조 도우미 2.0.3 백업

2026-09-17 사용자 확인을 마친 0917d 수정본을 2.0.3으로 버전 전환한 백업입니다. 보험 정산·국민연금 양식·수기 재원 반영·직종 수정·파일 선택 아이콘 개선이 포함됩니다.

- [2.0.3 실행파일 및 전체 소스](https://github.com/isilria/Social-insurance-updates/releases/tag/v2.0.3)
- [변경 사항 및 사용 안내](RELEASE_2.0.3.md)
- [회귀 검증 결과: 177개 통과](TEST_RESULTS.txt)
- [파일 무결성 값](SHA256SUMS.txt)

backup/v2.0.3 브랜치와 v2.0.3 태그에 보관합니다. 자동 업데이트의 main/latest.ini와 기존 2.0.2 배포본은 변경하지 않았습니다.

GitHub 자동 생성 소스 ZIP에는 내장 서식·라이브러리가 포함되지 않습니다. 직접 빌드하려면 릴리스의 SocialInsurance_Ver2.0.3_Source.zip을 받아 src/build_Ver2.0.ps1을 실행하세요. test.ps1은 실제 Excel을 사용해 회귀 검증을 수행합니다. 선택적인 실자료 검증은 별도 입력 파일 경로를 지정해야 합니다.

Windows 및 .NET Framework 4.8이 필요하며, 제출서 생성에는 Microsoft Excel이 필요합니다. 실제 급여·보험 원본, 개인별 생성 결과, 사용자 설정과 인증정보는 포함하지 않습니다.
