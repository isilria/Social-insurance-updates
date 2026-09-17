# 사회보험 재원별 대사 보조 도우미

## 최신 정식 버전: 2.0.3

2026-09-17 사용자 검증 및 177개 회귀 검사 통과 후 자동 업데이트 배포를 활성화했습니다.

- [실행파일 다운로드](https://github.com/isilria/Social-insurance-updates/releases/download/v2.0.3/SocialInsurance_Reconciliation_Helper_Ver2.0.3.exe)
- [릴리스 및 변경 내역](https://github.com/isilria/Social-insurance-updates/releases/tag/v2.0.3)
- [전체 소스·서식·라이브러리](https://github.com/isilria/Social-insurance-updates/releases/download/v2.0.3/SocialInsurance_Ver2.0.3_Source.zip)
- [검증 기록](TEST_RESULTS.txt)
- [파일 SHA256](SHA256SUMS.txt)

보험 정산 중복 집계, 국민연금 보수월액 인식, 보험만 있는 대상자의 수기 재원 지정과 제출서 반영을 수정했습니다. 교육공무직 제출서의 보수월액·정산 칸 및 기존 보험료 계산 수식을 유지합니다. 대상자 상세에서 직종을 수정할 수 있으며 파일 선택 아이콘을 개선했습니다.

프로그램의 설정 → 업데이트 확인에서 2.0.3을 내려받을 수 있습니다. main/latest.ini가 자동 업데이트 기준입니다. 이전 배포본과 버전별 백업은 보존합니다.

직접 빌드할 때는 전체 소스 ZIP의 src/build_Ver2.0.ps1을 사용하세요. GitHub 자동 생성 Source code ZIP에는 서식과 라이브러리가 포함되지 않습니다. Windows 및 .NET Framework 4.8이 필요하며 제출서 생성에는 Microsoft Excel이 필요합니다.

직종 수정은 저장한 대사 결과에 적용됩니다. 새 대사 결과에 자동 승계하지 않습니다. 정산 분리 정보가 없는 구버전 결과는 원본으로 다시 대사해야 합니다. 실제 급여·보험 개인정보 자료와 개인 설정은 배포 파일에 포함하지 않습니다.
