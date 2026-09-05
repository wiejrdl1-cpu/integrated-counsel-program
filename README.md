# 통합조사 상담 프로그램 v1

1팀(국기초)과 2팀(차상위, 기초연금 등) 상담내역 작성 기능을 하나로 통합한 Windows 데스크톱 프로그램입니다.

## v1 주요 변경사항

- 1팀 주거유형에 `민간(월세)` 추가
- 1팀과 2팀의 주거유형 목록과 금액 입력 규칙 통일
- 사용하지 않는 입력칸을 회색 음영으로 구분
- 1팀 부양의무자 소득인정액 항목 삭제
- 이전 자동저장 자료의 종전 항목은 오류 없이 무시하도록 호환성 유지

## 개발 실행

Python 3을 설치한 뒤 다음 파일을 실행합니다.

```text
RUN_APP.bat
```

## Windows 배포본 빌드

```text
BUILD_EXE.bat
```

빌드가 완료되면 다음 파일을 실행할 수 있습니다.

```text
dist\integrated_counsel_program_v1\integrated_counsel_program_v1.exe
```

## 개인정보 주의

이 저장소에는 실제 상담 내역, 주민등록번호, 성명, 주소, 비밀번호가 포함된 엑셀·JSON·로그·자동저장 파일을 올리지 마세요. 배포 전에는 소속기관의 보안 정책을 확인해야 합니다.

## 검증

```text
python -m unittest discover -s tests -v
```

세부 변경 이력은 `CHANGELOG.txt`를 확인하세요.
