# Claude Cowork 설치 키트 (Windows)

새 Windows PC에서 **Claude Cowork**을 바로 쓸 수 있도록 점검부터 설치까지 자동으로 해 주는 스크립트입니다.
직접 할 일은 **관리자 권한 허용**, (필요하면) **재시작**, 그리고 **로그인 1번**뿐입니다.

> **준비물**: Windows 10(2004 이상) 또는 11 · **관리자 권한** · 인터넷 · Claude **Pro 이상** 유료 계정
> WSL은 필요 없습니다.

---

## 1. 설치 — 아래 방법 중 하나만 하세요

### 방법 A. 한 줄 명령 (가장 쉬움 · 추천)

1. 키보드에서 `Win` + `X` → **터미널(관리자)** 또는 **Windows PowerShell(관리자)** 클릭 → [예]
2. 아래 한 줄을 복사해서 붙여넣고 `Enter`

   ```powershell
   irm https://raw.githubusercontent.com/tiger3c/claude-cowork-setup/main/bootstrap.ps1 | iex
   ```

3. 화면 안내대로 진행 → [2. 설치 중에 나오는 질문](#2-설치-중에-나오는-질문)

### 방법 B. ZIP 내려받기

1. 이 페이지 위쪽 초록색 **`<> Code`** 버튼 → **Download ZIP**
2. 받은 ZIP 파일에서 **마우스 오른쪽 → 모두 압축 풀기**
   (ZIP을 열어서 안에서 바로 실행하면 안 됩니다)
3. 풀린 폴더에서 **`setup.bat`** 더블클릭
   - 파란 창 "Windows의 PC 보호"가 뜨면 **추가 정보 → 실행**
   - 관리자 권한 창(UAC)에서 **[예]**

### 방법 C. USB로 가져가기 (설치할 PC의 인터넷이 느리거나 불안할 때)

**지금 쓰는 PC에서 미리:**

1. 방법 B처럼 ZIP을 받아 압축을 풉니다.
2. 브라우저로 아래 파일을 받아서, 풀린 폴더 안의 **`files`** 폴더에 넣습니다.

   | 파일 | 내려받기 | 필요 |
   |---|---|---|
   | Claude 오프라인 설치 파일 (수 GB, Cowork 실행 환경 포함) | [x64](https://claude.ai/api/desktop/win32/x64/offline/latest/redirect) · [ARM](https://claude.ai/api/desktop/win32/arm64/offline/latest/redirect) | 권장 |
   | Cowork 공식 점검 프로그램 | [x64](https://claude.ai/api/desktop/win32/x64/cowork-readiness-check/latest/redirect) · [ARM](https://claude.ai/api/desktop/win32/arm64/cowork-readiness-check/latest/redirect) | 선택 |

   > 대부분의 PC는 **x64** 입니다. ARM은 퀄컴 Snapdragon 노트북 같은 경우만 해당합니다.
   > 오프라인 링크가 404면 아직 최신판이 안 나온 것이니 조금 뒤에 다시 받으세요.

3. `check-only.bat` 을 실행해서 **"설치 파일 찾음"** 이 나오는지 확인합니다.
4. 폴더를 통째로 USB에 복사 → 새 PC에서 `setup.bat` 실행 (재시작할 때 USB를 뽑지 마세요)

> 설치 파일을 안 넣어도 됩니다. 없으면 스크립트가 공식 주소에서 일반 설치 파일(약 300MB)을 자동으로 받습니다.
> 브라우저 **다운로드 폴더**에 받아둔 Claude 설치 파일도 자동으로 찾아서 씁니다.

---

## 2. 설치 중에 나오는 질문

| 질문 | 추천 | 설명 |
|---|---|---|
| 하이퍼바이저를 켤까요? | **Y** | VMware/VirtualBox 등이 꺼둔 설정입니다. 오래된 VMware/VirtualBox를 쓰는 경우만 주의 |
| 예전 버전을 삭제할까요? | **Y** | 예전 .exe 버전에는 Cowork이 없습니다. 대화 기록은 계정에 있어서 안 지워집니다 |
| Claude Code도 같이 설치할까요? | 원하면 **Y** | 터미널에서 `claude` 로 쓰는 Claude. 가상화가 안 되는 PC에서는 Cowork 대신 쓰세요 |
| 지금 다시 시작할까요? | 작업 저장 후 **y** | 가상화 기능은 재시작해야 켜집니다 |

- 그냥 `Enter` 를 누르면 대문자 쪽이 선택됩니다 (`[Y/n]` → 예, `[y/N]` → 아니요).
- 한글 자판 상태에서 `y` 를 눌러 `ㅛ` 가 입력돼도 "예"로 인식합니다.
- **재시작하면** Windows에 로그인하자마자 설치 창이 **자동으로 다시 떠서** 마무리합니다 (UAC [예]).

---

## 3. 설치 후 사용

### Cowork
1. Claude 앱이 자동으로 열립니다 (안 열리면 시작 메뉴 → **Claude**)
2. **로그인** (이메일 또는 Google)
3. **Cowork** 탭 선택 → 작업할 폴더를 고르고 할 일을 입력
   - 오프라인 설치 파일을 안 쓴 경우, 처음 한 번은 실행 환경을 내려받느라 몇 분 걸립니다.

### Claude Code (설치한 경우)
1. **새** PowerShell 창을 엽니다 (설치 전에 열어둔 창에서는 `claude` 가 안 잡힐 수 있음)
2. 작업 폴더로 이동해서 `claude` 입력 → 처음 한 번 로그인

---

## 4. 스크립트가 하는 일

| 단계 | 내용 |
|---|---|
| 1. PC 점검 | Windows 버전, 메모리(8GB 이상 권장), 디스크(20GB 이상), BIOS 가상화, 가상 PC 여부, Home 에디션 |
| 2. 인터넷 | `claude.ai`, `downloads.claude.ai` 연결 확인 (회사 프록시 지원) |
| 3. 가상화 | **Virtual Machine Platform** 켜기 (실패하면 DISM으로 재시도), 꺼진 하이퍼바이저 켜기 |
| 4. Claude Desktop | `files` / 다운로드 폴더의 설치 파일을 쓰고(오프라인판 우선), 없으면 공식 주소에서 받기 → **Anthropic 서명·게시자 확인** → 모든 사용자용 설치 |
| 5. Claude Code | 공식 설치 스크립트 실행 + PATH 등록 (가상화가 안 되는 PC면 대안으로 권장) |
| 6. 마무리 | 재시작 필요 시 자동 재개 등록 → 공식 점검 프로그램(있으면) → Claude 실행 |

- 여러 번 실행해도 안전합니다. 이미 끝난 단계는 건너뜁니다.
- 로그는 `logs` 폴더(안 되면 `%TEMP%\ClaudeCoworkSetup`)에 남습니다.

---

## 5. 문제 해결

| 증상 | 해결 |
|---|---|
| `BIOS/UEFI 에서 가상화가 꺼져 있습니다` | 아래 [BIOS 가상화 켜기](#bios-가상화-켜기) |
| Claude 앱에서 "Virtualization is not available" | ① **시스템 종료가 아니라 [다시 시작]** ② `setup.bat` 다시 실행 ③ BIOS 가상화 확인 |
| `가상 PC(VM/VDI/클라우드 PC) 안에서 실행 중` | 그 PC 관리자에게 "중첩 가상화(nested virtualization)" 허용 요청. 안 되면 Claude Code 사용 |
| `Home 에디션입니다` 경고 후 Cowork이 안 됨 | Claude Code 사용 (`install-claude-code.bat`) |
| 창이 바로 닫히거나 "스크립트를 실행할 수 없으므로" | 회사 정책으로 막힌 경우 → 아래 [수동 설치](#수동-설치-스크립트가-막혔을-때) |
| `0x80073D02` | Claude 앱이 켜져 있음 → 작업 표시줄 트레이에서 완전히 종료 후 다시 실행 |
| `0x80073CF6` · `0x80073CFB` · `0x80073CF3` | 설정 → 앱 → 설치된 앱 → Claude 제거 후 다시 실행 |
| `0x800B0109` 등 인증서 오류 | PC 날짜·시간 확인, Windows 업데이트 후 다시 실행 |
| `0x80073CFF` | 정책상 앱 직접 설치가 막힌 PC → IT 담당자 문의 |
| `0x80073D28` | 관리자 권한 없음 → `setup.bat` 으로 실행하고 [예] |
| `claude` 명령을 못 찾음 | **새** 터미널 열기. 그래도 안 되면 `install-claude-code.bat` |
| `claude.ai 에 연결하지 못했습니다` | 인터넷/회사 방화벽 확인. 안 되면 방법 C(USB)로 설치 파일 준비 |
| 그 밖의 오류 | `logs` 폴더의 로그 파일 내용을 claude.ai 에 붙여넣고 물어보기 |

### BIOS 가상화 켜기

1. **설정 → 시스템 → 복구 → 고급 시작 [지금 다시 시작]** → 문제 해결 → 고급 옵션 → **UEFI 펌웨어 설정** → 다시 시작
   (또는 전원을 켤 때 `F2` / `Del` / `F10` / `F1` 연타 — 제조사마다 다름)
2. 아래 항목을 **Enabled** 로 바꿉니다. 보통 *Advanced*, *CPU Configuration*, *Security* 메뉴 안에 있습니다.
   - Intel: `Intel Virtualization Technology` (또는 `VT-x`, `VMX`)
   - AMD: `SVM Mode` (또는 `AMD-V`)
3. 저장(`F10`) 후 재부팅 → `setup.bat` 다시 실행

> 회사 PC는 BIOS에 비밀번호가 걸려 있을 수 있습니다. 작업 관리자 → 성능 → CPU 의 **가상화: 사용** 으로 켜졌는지 확인할 수 있습니다.

### 수동 설치 (스크립트가 막혔을 때)

**관리자** PowerShell에서 한 줄씩 실행합니다.

```powershell
# 1) 가상화 기능 켜기 → 끝나면 PC [다시 시작]
Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All -NoRestart

# 2) 설치 파일 받기 (오프라인판을 브라우저로 받아뒀다면 생략하고 그 파일 경로를 쓰세요)
Invoke-WebRequest -UseBasicParsing https://claude.ai/api/desktop/win32/x64/msix/latest/redirect -OutFile "$env:TEMP\Claude.msix"

# 3) 설치 (모든 사용자용 + Cowork 서비스 등록 → 지금 계정에 등록)
Add-AppxProvisionedPackage -Online -PackagePath "$env:TEMP\Claude.msix" -SkipLicense
Add-AppxPackage -Path "$env:TEMP\Claude.msix"
```

Claude Code만 따로 (관리자 아닌 일반 PowerShell):

```powershell
irm https://claude.ai/install.ps1 | iex
# PATH 에 없으면 추가 (새 터미널부터 claude 명령 사용 가능)
$p = [Environment]::GetEnvironmentVariable('Path','User'); if ($p -notlike '*\.local\bin*') { [Environment]::SetEnvironmentVariable('Path', "$p;$env:USERPROFILE\.local\bin", 'User') }
```

---

## 6. 자주 묻는 질문

**Q. 로그인까지 미리 해두고 가져갈 수 없나요?**
안 됩니다. 로그인 정보는 PC와 Windows 계정에 묶여서 암호화 저장되기 때문에 다른 PC로 옮길 수 없습니다. 로그인은 1분이면 끝납니다.
(로그인 화면 자체를 없애려면 Anthropic API 키로 미리 설정하는 방식이 있지만, 구독과 별도로 사용량만큼 API 요금이 나갑니다.)

**Q. WSL이 필요한가요?**
아니요. Cowork은 Windows 기본 기능(Virtual Machine Platform)만 씁니다. 단, BIOS 가상화가 꺼져 있으면 WSL도 Cowork도 둘 다 안 됩니다.

**Q. 가상화가 안 되는 PC에서는요?**
Cowork은 쓸 수 없습니다. 대신 **Claude Code**(터미널)와 Claude 앱의 **일반 채팅**은 됩니다. 스크립트가 자동으로 Claude Code 설치를 권합니다.

**Q. 지우려면?**
- Claude Desktop: 설정 → 앱 → 설치된 앱 → **Claude** 제거
- Claude Code: `%USERPROFILE%\.local\bin\claude.exe` 와 `%USERPROFILE%\.local\share\claude` 폴더 삭제

---

## 파일 구성

| 파일 | 용도 |
|---|---|
| `setup.bat` | **설치 — 이것만 실행하면 됩니다** |
| `check-only.bat` | 아무것도 바꾸지 않고 점검만 (관리자 권한 불필요) |
| `install-claude-code.bat` | Claude Code만 설치 (관리자 권한 불필요) |
| `setup.ps1` | 설치 스크립트 본체 |
| `bootstrap.ps1` | 한 줄 설치(방법 A)용 |
| `files/` | 미리 받은 설치 파일을 넣는 곳 (선택) |

## 참고 (공식 문서)

- [Deploy Claude Desktop for Windows](https://support.claude.com/en/articles/12622703-deploy-claude-desktop-for-windows)
- [Claude Desktop 설치 / 오프라인 설치](https://claude.com/docs/third-party/claude-desktop/installation)
- [Windows 배포 체크리스트](https://claude.com/docs/government/deploy-desktop/windows-checklist)
- [Claude Code 설치](https://code.claude.com/docs/en/setup)
