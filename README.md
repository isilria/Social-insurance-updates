# 사회보험 재원별 대사 보조 도우미

Windows용 급여대장·사회보험 부과자료 대사 및 제출서 생성 프로그램입니다.

## 최신 정식 버전: 2.0.2

- [실행파일 다운로드](https://github.com/isilria/Social-insurance-updates/releases/download/v2.0.2/SocialInsurance_Reconciliation_Helper_Ver2.0.2.exe)
- [2.0.2 배포 및 변경 내역](https://github.com/isilria/Social-insurance-updates/releases/tag/v2.0.2)
- [전체 소스·서식·라이브러리·빌드자료](https://github.com/isilria/Social-insurance-updates/releases/download/v2.0.2/SocialInsurance_Ver2.0.2_Source.zip)

이번 배포에서는 PDF를 약 600dpi 무손실 압축으로 개선했습니다. 기존 A4 배치와 보정·확인·제출서 생성 기능을 유지합니다. 이전 PDF는 새 버전에서 다시 생성해야 개선된 품질이 적용됩니다.

## 소스 관리

`src/`에는 현재 배포본의 C# 소스와 빌드 스크립트를 보관합니다.
직접 빌드할 때는 위의 **전체 소스 압축파일**을 받으세요. GitHub가 자동으로 제공하는 Source code.zip에는 내장 서식과 라이브러리가 포함되지 않습니다.

- `InsurancePayrollValidator_Ver2.0.cs`: 기본 화면·입력·대사·출력
- `TestFeatures202.cs`: 2.0.2 화면 및 제출서 연결
- `ManualContributions202.cs`: 수기 보정·확인 처리
- `PrintQuality202.cs`: 고화질 PDF 이미지 생성
- `build_Ver2.0.ps1`: 실행파일 빌드
- `RELEASE_2.0.2.md`: 변경 내역과 검증 항목

Windows 10/11, .NET Framework 4.8 환경을 사용합니다. 제출서 생성에는 Microsoft Excel이 필요합니다.
실제 급여자료·보험 원본·개인 설정·인증정보는 배포 파일에 포함하지 않습니다.

## 자동 업데이트

`latest.ini`가 프로그램의 공개 업데이트 기준입니다. 새 릴리스의 실행파일을 게시하고 SHA256을 확인한 뒤 갱신합니다. 이전 릴리스는 Releases에서 확인할 수 있습니다.

Ver. 2.0.2 @ 살구아빠
