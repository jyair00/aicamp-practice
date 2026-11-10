---
name: setup-environment
description: SK브로드밴드(SKB) AI Camp 수강생의 PC에서 Node.js·Python과 파일 실습용 라이브러리를 준비한다. "SKB 세팅해줘", "처음 세팅", "환경 준비", "Node랑 Python 설치" 요청에 사용한다. Windows 스토어 실행 별칭과 실제 Python을 구분하며, 한글 처리는 별도 korean-safe-windows 스킬과 연결한다.
---

# SKB 실습 환경 준비

너는 수강생 컴퓨터에서 명령을 실행할 수 있는 AI다. **컴퓨터 상태를 읽고, 필요한 명령을 직접 골라 하나씩 실행한다.** 설치용 `.bat`, `.ps1`, `.sh`, `.py` 파일이나 일괄 설치 프로그램을 만들지 않는다. 아래 명령은 필요한 경우에만 사용하는 예시이며, 현재 셸과 실제 경로에 맞춘다.

목표는 **Node.js·npm·npx, 실제 실행되는 Python, 문서 라이브러리 5개**가 이 실습 폴더에서 작동하는 것이다. 한글 경로·파일 인코딩·한글 검증의 정본은 [korean-safe-windows](../korean-safe-windows/SKILL.md)다. 한글 경로를 다루기 전에 그 스킬을 읽고, 설치 마지막에도 해당 스킬로 한글 검증을 한다. 계정 로그인, 사내 접속 설정, 다른 스킬 설치, 실습 과제 수행은 포함하지 않는다.

## 1. 어느 컴퓨터의 어느 폴더인지 확인

수강생에게 부드러운 존댓말로 "이 컴퓨터에 필요한 프로그램이 있는지 확인하고, 없는 것만 준비할게요."라고 알린다. 각 작업의 시작과 결과를 한 줄씩 설명한다. 설치 중에도 진행 상황을 알리고, 근거 없는 소요 시간을 약속하지 않는다.

- 세션 정보와 실제 명령 실행 환경으로 OS, CPU 종류, 셸, 현재 폴더, 쓰기 권한을 확인한다. **AI 도구의 이름이나 PowerShell의 유무만으로 OS를 추정하지 않는다.** Windows에서도 Git Bash·PowerShell·WSL이 다르고, PowerShell은 macOS에도 있을 수 있다.
- Cloud PC·WSL·컨테이너·원격 서버에서 실행 중이면 그 사실을 알린다. 수강생이 실습할 환경과 다르면 설치 전에 대상부터 확인한다. 브라우저 채팅만 가능하고 로컬 명령 도구가 없으면 설치했다고 말하지 말고, 강사가 안내한 로컬 실행 환경을 열도록 안내한다.
- 현재 선택된 실습 폴더를 기준으로 삼는다. 폴더 이름을 `SKB`, `실습`, 특정 번호로 고정하지 않는다. 다운로드 중복으로 붙은 `(1)`, 한글 사용자명, OneDrive 경로도 그대로 다룬다. ZIP 안을 열었거나 시스템 폴더·사용자 홈만 열린 경우에는 압축을 푼 실습 폴더를 선택하게 한다.
- 기존 `CLAUDE.md`, `AGENTS.md`, 환경 기록, 가상환경이 있으면 먼저 읽고 재사용 여부를 확인한다. 다시 실행해도 정상인 프로그램을 지우거나 같은 것을 무조건 재설치하지 않는다.

Windows 명령은 실제 PowerShell에서 실행하는 것을 우선한다. Git Bash 도구만 있으면 사용 가능한 `powershell.exe` 또는 `pwsh`에 **짧은 명령 하나**를 전달한다. Bash에 PowerShell 문법을 그대로 쓰지 않는다. PowerShell에서는 `where` 대신 `where.exe` 또는 `Get-Command`를 쓴다. macOS는 실제 Bash/zsh를 사용한다.

## 2. Windows에서 실제 Python 찾기 — 가장 중요

**설치 위치·명령 이름·버전 출력 하나만 보고 통과시키지 않는다.** 먼저 명령이 무엇을 가리키는지 조회한다.

```powershell
Get-Command py, pymanager, python, python3 -All -ErrorAction SilentlyContinue
```

Python Install Manager가 확인되면 `pymanager list`로 설치된 런타임을 본다. 기존 Python Launcher가 있으면 `py -0p`로 경로를 본다. 둘 다 없으면 조회된 실행 파일, 사용자용 설치 위치, 회사가 제공한 Python을 확인한다. **`py`가 없다는 이유만으로 Python이 없다고 판정하지 않는다.**

다음 셋을 구분한다.

- **스토어로 보내는 App Installer 별칭(stub)**: 이름만 있는 바로가기다. 실제 Python 코드를 실행하지 못하고 Store 안내나 미설치 메시지가 나온다. 실습에 쓸 Python으로 선택하지 않는다. 인자 없이 `python`을 실행해 스토어를 일부러 열지 않는다.
- **실제로 설치된 Microsoft Store Python**: 바로가기와 다른 정상 배포판이다. 코드 실행·가상환경·패키지·파일 접근 검사를 통과하면 사용할 수 있다. 경로에 `WindowsApps`가 있다는 이유만으로 가짜라고 부르거나 삭제하지 않는다.
- **공식 Python Install Manager / 기존 Launcher**: 실행기를 설치한 것과 Python 런타임을 설치한 것은 다르다. Manager는 python.org와 Store 양쪽에서 제공된다. `py`나 `pymanager`가 있다는 것만으로 런타임이 있다고 보지 않는다. 새 Manager는 런타임 없이 실행할 때 자동 설치할 수 있으므로 목록 조회부터 한다.

실행 후보를 찾았으면 **선택한 실행 파일**로 아래 Python 코드를 실행한다. Windows에서 전체 실행 경로를 호출할 때는 PowerShell 호출 연산자 `&`와 안전하게 인용한 경로를 쓴다.

```python
import sys, struct
print(sys.version)
print(sys.executable)
print(struct.calcsize('P') * 8)
```

종료 코드와 출력이 모두 정상인지 확인한다. 화면 출력이 비면 곧바로 stub이라고 단정하지 말고, 종료 코드·파일 출력으로 확인한다. 확인한 `sys.executable` 경로를 이후의 기준으로 삼는다. **항상 같은 Python으로 설치와 실행을 한다.** 기본 `py`가 다음에 다른 버전을 선택할 수 있으므로 이후 맨 `pip`·`python`·`py`에 의존하지 않는다.

공식 일반 설치본이나 회사 제공 Python이 정상이라면 그것을 우선한다. 기존 Python 3.11 이상은 지원 상태와 패키지 호환성을 확인해 사용한다. 3.14가 있다는 이유만으로 낮추지 않는다. 너무 오래된 버전이나 패키지 호환 실패가 확인된 경우에만 새 런타임을 추가한다.

별칭이 실습을 방해해도 먼저 정상 Python의 전체 경로로 진행한다. 설정 변경이 꼭 필요하면 **설정 → 앱 → 고급 앱 설정 → 앱 실행 별칭**에서 제공자가 **App Installer**인 `python.exe`·`python3.exe` 항목만 구분해서 안내한다. Python Install Manager의 정상 별칭을 전부 끄거나, `WindowsApps` 폴더를 삭제하거나, PATH에서 통째로 빼지 않는다.

## 3. 없는 Python만 설치

**Windows 신규 설치 기준은 일반 CPython 3.13이다.** 패치 버전과 CPU용 설치 파일은 공식 목록을 확인한다. 시험판, free-threaded 빌드, embeddable ZIP을 기본으로 고르지 않는다.

- 사용 가능한 공식 Install Manager가 이미 있으면 `pymanager install 3.13`으로 런타임을 추가할 수 있다. 기존 Launcher에 Manager 전용 명령을 보내지 않는다.
- winget이 있으면 먼저 해당 패키지와 설치 범위를 확인한다. 사용할 수 있을 때 다음과 같이 정확한 ID와 `winget` 소스를 지정한다.

```powershell
winget install --id Python.Python.3.13 --exact --source winget
```

- 관리자 권한이 없으면 그 설치 프로그램이 지원하는 사용자 설치를 선택한다. `--scope user`는 지원 여부를 확인한 뒤 사용한다. winget이 없거나 회사에서 막았으면 [Python Windows 다운로드](https://www.python.org/downloads/windows/)에서 3.13의 일반 설치 파일을 찾는다. 전통적인 `.exe` 설치 창에서는 `Add python.exe to PATH`를 체크한다. **최상단 다운로드 버튼이 늘 이 설치 파일이라고 가정하지 않는다.** Manager의 MSIX를 받았다면 Manager 설치 후 런타임까지 확인한다.

**macOS에서는** `python3`, 버전별 실행 파일, 기존 Homebrew의 경로를 확인한 뒤 같은 방식으로 실제 실행 경로를 고른다. 없다면 이미 Homebrew가 있을 때만 `brew install python@3.13`을 고려한다. 설치 위치는 `brew --prefix python@3.13`으로 확인한다. Homebrew가 없으면 [Python macOS 다운로드](https://www.python.org/downloads/macos/)의 공식 설치 파일을 안내한다. 이 세팅을 위해 Homebrew나 Xcode를 새로 설치하지 않는다.

설치 프로그램의 성공 메시지나 `already installed`는 완료 증거가 아니다. 새로 설치된 **실제 실행 경로로 2단계 검사를 다시 한다.** 이전 AI 프로세스의 PATH가 오래되었으면 전체 경로로 계속한다. 그래도 새 설치를 사용할 수 없을 때만 AI 앱/터미널을 완전히 다시 열도록 안내한다. 설치를 반복해서 해결하려 하지 않는다.

## 4. Node.js와 npm·npx 확인

`node --version`, `node -p "process.execPath"`, `npm --version`, `npx --version`을 각각 확인한다. Node만 있다고 npm·npx까지 정상이라고 판단하지 않는다. 여러 설치본이 있으면 명령별 경로도 확인하고, 선택한 Node와 함께 설치된 npm·npx를 사용한다.

- 새로 설치할 때는 **공식 LTS**를 선택한다. 2026-09-08 기준 Node 24·22가 LTS이며, 나중에 사용할 때는 [공식 지원표](https://nodejs.org/en/about/previous-releases)를 확인한다. 기존 지원 버전이 정상이고 실습 요구를 만족하면 유지한다.
- Windows는 `winget`의 `OpenJS.NodeJS.LTS` 패키지를 확인해 설치하거나 [Node.js 공식 다운로드](https://nodejs.org/en/download)에서 OS·CPU에 맞는 LTS 설치 파일을 받는다. 회사 설치 정책과 실제 권한에 맞는 방법을 고른다.
- macOS는 기존 Homebrew가 있으면 LTS formula(예: `node@24`)를 사용한다. 없으면 공식 `.pkg`를 안내한다. 버전 관리 도구가 이미 있으면 기존 방식을 존중한다.
- **PowerShell에서 `npm.ps1`·`npx.ps1`이 실행 정책에 막히면**, 같은 설치 폴더의 `npm.cmd`·`npx.cmd`를 호출한다. 실행 정책을 바꾸지 않는다. 전체 경로를 쓸 때도 `&`를 사용한다.
- npm·npx가 내부에서 Node를 못 찾으면, 선택한 Node의 폴더를 **해당 명령 프로세스의 PATH 앞에만** 추가하고 재검사한다. PATH를 통째로 덮어쓰거나 `setx PATH`를 쓰지 않는다.

세 가지 버전 검사와 실제 Node 코드 실행이 통과해야 Node 준비 완료다. 버전 확인 때문에 임의의 npm 패키지를 내려받아 실행하지 않는다.

## 5. 실습 전용 Python과 라이브러리 5개

실습 폴더 안의 **가상환경(이 폴더 전용 Python 공간)**을 사용한다. 정상적인 기존 `.venv`가 있으면 재사용한다. 없으면 2단계에서 확인한 Python의 전체 경로로 `-m venv .venv`를 실행한다. 기존 환경이 깨져 있으면 원인을 확인하고, 삭제·덮어쓰기 대신 별도 이름의 실습 환경을 만든다.

그다음부터는 가상환경의 Python을 직접 호출한다. Windows는 `.venv\Scripts\python.exe`, macOS는 `.venv/bin/python`이다. 폴더와 실행 파일의 실제 경로를 확인해서 사용한다. **활성화 스크립트를 실행할 필요가 없다.** 그래서 `Activate.ps1`을 위해 실행 정책을 바꿀 필요도 없다.

가상환경 Python으로 `sys.executable`, `sys.prefix`, `sys.base_prefix`를 확인한다. 실행 파일과 prefix가 선택한 가상환경을 가리키고 `sys.prefix != sys.base_prefix`인지 본다. `-m pip --version`의 경로도 같은 환경 안이어야 한다.

패키지는 LG판과 같은 다섯 개만 준비한다. `openpyxl`은 엑셀, `python-docx`는 Word, `python-pptx`는 PPT, `pdfplumber`는 PDF 읽기, `pandas`는 표 정리·통합용이다. 이미 설치되어 정상 import되는 것은 불필요하게 업그레이드하지 않는다.

Windows PowerShell에서의 설치 예시:

```powershell
& '.\.venv\Scripts\python.exe' -X utf8 -m pip install --only-binary=:all: openpyxl python-docx python-pptx pdfplumber pandas
```

macOS에서는 `./.venv/bin/python`으로 같은 인자를 사용한다. `--only-binary=:all:`은 미리 빌드된 배포 파일만 사용하게 한다. 호환 파일이 없으면 어떤 패키지·Python 버전·CPU 조합 때문인지 확인한다. 무조건 컴파일로 전환하거나 Visual C++·Xcode를 설치하지 않는다.

- pip만 없으면 **그 가상환경 Python**으로 `-m ensurepip --upgrade`를 먼저 시도한다. pip 누락을 Python 전체의 부재로 오해하지 않는다.
- `externally-managed-environment`가 나오면 시스템 Python을 잘못 호출했는지 확인한다. 가상환경으로 해결하며 `--break-system-packages`를 붙이지 않는다.
- 인증서나 사내망 오류는 Python 부재나 한글 문제와 구분한다. 현재 접속 결과로 판단한다. 회사가 제공한 인증서·프록시·패키지 저장소가 있으면 그 설정을 사용한다. `--trusted-host`, `strict-ssl=false`, `NODE_TLS_REJECT_UNAUTHORIZED=0`으로 검증을 끄지 않는다.
- 한글 임시 경로 오류가 확인되면 [korean-safe-windows의 설치 경로 대응](../korean-safe-windows/SKILL.md)을 적용한다. 한글 사용자명만 보고 원인으로 단정하지 않는다.

마지막에는 설치 성공 문구에 의존하지 말고, **가상환경 Python으로 모듈 다섯 개를 각각 import**한다. 설치 이름과 import 이름은 순서대로 `openpyxl`, `docx`, `pptx`, `pdfplumber`, `pandas`다. 하나가 실패하면 나머지도 확인하고 성공·실패를 따로 기록한다.

## 6. 한글 스킬로 최종 확인

[korean-safe-windows](../korean-safe-windows/SKILL.md)의 **한글 처리 검증**을 수행한다. 여기서는 그 절차를 복사하거나 다시 정의하지 않는다. 선택한 Python·Node와 실습 폴더를 기준으로 검증하게 하고, 텍스트 교차 읽기·CSV·XLSX 결과를 각각 받는다. 한글 스킬이 도구 부족으로 미완료를 반환하면 필요한 설치를 이 스킬에서 판단한다. 같은 설치와 검증을 무한 반복하지 않는다.

## 7. 다음 대화에 쓸 정보를 남기고 마무리

실습 폴더의 `.skb-setup-status.md`에 확인 시점·대상 환경·실습 폴더, Node/npm/npx의 버전과 실제 경로, 가상환경 Python 경로·버전, 모듈 5개 검사, 한글 검증, 변경한 설정, 남은 문제를 간단히 기록한다. 비밀번호·토큰·인증서 내용·프록시 자격증명은 적지 않는다. 재실행 때 이전 성공 기록만 믿지 말고 실제 상태를 다시 확인한다.

실습 루트의 `CLAUDE.md`는 브랜드와 실습 작업의 기준이다. 이 파일에 스킬 설명을 추가하거나 설치 로그로 덮어쓰지 않는다. PC별 실행 경로와 검사 결과는 `.skb-setup-status.md`만 갱신한다. 다른 위치에서 이 스킬을 사용하는 경우에도 기존 프로젝트 지침을 보존한다. 시스템·전역 AI 지침은 수정하지 않는다.

모든 필수 검사가 통과했으면 "Node.js·Python과 파일 실습 도구 준비가 끝났어요. 한글 경로에서 파일을 만들고 다시 읽는 것도 확인했어요. 새 대화에서 실습을 시작해주세요."라고 짧게 마무리한다.

일부가 막혔으면 **부분 완료**라고 말하고, 무엇이 준비되었고 어떤 실습이 아직 안 되는지와 다음 행동을 알린다. Node가 실패했거나 한글 검증을 안 했는데 전체 완료라고 말하지 않는다. 재시작 후 검사를 해야 한다면 완료로 처리하지 않는다.

## 직접 조작이나 회사 지원이 필요할 때

자동으로 해결할 수 있는 범위는 계속 진행한다. 원인이 같은 실패를 그대로 반복하지 않는다. 권한 팝업, 비밀번호, 설치 창 클릭, 앱 재시작, 회사 승인·망 허용처럼 직접 처리가 필요한 단계에서는 **지금 상태 → 할 일 하나와 클릭 순서 → 필요하면 강사 호출**을 안내하고 그 작업의 완료를 기다린다. 직접 조작 중에 연관된 설치를 몰래 이어가지 않는다.

사내망이 막혔다면 실제 실패한 호스트와 기능만 기록한다. Python 설치는 `python.org`·`www.python.org`, Python 패키지는 `pypi.org`·`files.pythonhosted.org`, Node는 `nodejs.org`, npm 패키지는 `registry.npmjs.org` 등을 사용한다. winget이나 리다이렉트에서 다른 호스트가 나올 수 있으므로 이 목록만으로 원인을 확정하지 않는다. 과거 SKB 메모를 현재 차단 상태라고 단정하지 않는다. 승인된 회사 배포 파일·오프라인 패키지가 있으면 그 경로로 계속할 수 있다.

## 공식 근거

- [Python Windows 설치·Manager·Launcher](https://docs.python.org/3/using/windows.html)
- [Microsoft: 스토어 바로가기와 앱 실행 별칭](https://learn.microsoft.com/en-us/windows/dev-environment/python)
- [Python: 가상환경은 활성화 없이 전체 경로로 사용 가능](https://docs.python.org/3/library/venv.html)
- [pip: 미리 빌드된 배포 파일만 설치](https://pip.pypa.io/en/stable/cli/pip_install/#cmdoption-only-binary) · [pip: 인증서 검증](https://pip.pypa.io/en/stable/topics/https-certificates/)
