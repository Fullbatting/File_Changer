# src/launcher/get-node.ps1
# PC에 쓸 수 있는 Node.js가 없을 때 공식 배포판(설치 없이 쓰는 zip)을 프로그램 폴더 안에 내려받는다.
# - 버전과 SHA-256 검사값을 여기에 고정해 두고, 내려받은 파일이 공식 파일과 똑같을 때만 쓴다.
# - 관리자 권한이 필요 없고 PC 설정(PATH 등)을 바꾸지 않는다. 지우려면 runtime 폴더를 지우면 된다.
# - 실패하면 스스로 오류 창을 띄우고 종료 코드 3으로 끝낸다(다른 코드면 PowerShell이 스크립트를 돌리지 못한 것).
# 이 파일은 UTF-8(BOM 포함)로 저장해야 Windows PowerShell 5.1이 한글을 바르게 읽는다.

param([Parameter(Mandatory = $true)][string]$Dest)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # 진행 막대를 그리면 내려받기가 몇 배 느려진다

$Version = 'v24.22.1'
$Sha256 = @{
  'x64'   = 'd2046d28c2c3dcd211ac5145c0bf30e2047ef9dfb306693c8c6e2eccd8b95fa5'
  'arm64' = 'a0c807319fede24c4a0d8f84564c82a619bd313b16b10c007c383c2609e99171'
}
$Title = '사회사업실 도우미 - 오류'

function Show-Error([string]$message) {
  [Console]::Error.WriteLine("[오류] $message")
  try {
    Add-Type -AssemblyName PresentationFramework
    [void][System.Windows.MessageBox]::Show($message, $Title, 'OK', 'Error')
  } catch {}
}

$stage = "$Dest.download"
$zip = Join-Path ([IO.Path]::GetTempPath()) ('swh-node-' + [guid]::NewGuid().ToString('N') + '.zip')
try {
  $arch = $null
  if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { $arch = 'arm64' }
  elseif ([Environment]::Is64BitOperatingSystem) { $arch = 'x64' }
  if (-not $arch) { throw '32비트 Windows에서는 이 프로그램을 쓸 수 없습니다. 64비트 Windows 10/11이 필요합니다.' }

  $name = "node-$Version-win-$arch"
  $url = "https://nodejs.org/dist/$Version/$name.zip"

  # 오래된 Windows의 PowerShell 5.1은 TLS 1.2를 기본으로 켜지 않는다. 회사 프록시는 Windows 로그인 정보로 통과
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
  if ([Net.WebRequest]::DefaultWebProxy) { [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials }

  Write-Host "Node.js $Version 을(를) 내려받습니다(약 40MB). 인터넷 연결이 필요하며 1~2분 걸릴 수 있습니다."
  try {
    Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
  } catch {
    throw ("Node.js를 내려받지 못했습니다. 인터넷 연결 상태를 확인한 뒤 run.bat을 다시 실행하세요." + [Environment]::NewLine + "(회사망이면 nodejs.org 접속이 허용되어야 합니다: " + $_.Exception.Message + ")" + [Environment]::NewLine + [Environment]::NewLine + "인터넷을 쓸 수 없는 PC라면, 인터넷이 되는 PC에서 run.bat으로 프로그램을 한 번 연 뒤 그 폴더를 통째로 이 PC에 옮겨 오세요.")
  }

  $hash = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
  if ($hash -ne $Sha256[$arch]) {
    throw '내려받은 Node.js 파일이 공식 파일과 다릅니다(검사값 불일치). 안전을 위해 쓰지 않았습니다. 잠시 뒤 run.bat을 다시 실행하세요.'
  }

  Write-Host 'Node.js 압축을 풉니다...'
  if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [IO.Compression.ZipFile]::ExtractToDirectory($zip, $stage)

  # 압축 안의 node-버전-win-x64 폴더를 runtime\node로 옮긴다(같은 드라이브 안이라 바로 옮겨진다)
  if (Test-Path -LiteralPath $Dest) { Remove-Item -LiteralPath $Dest -Recurse -Force }
  Move-Item -LiteralPath (Join-Path $stage $name) -Destination $Dest
  if (-not (Test-Path -LiteralPath (Join-Path $Dest 'node.exe'))) { throw 'Node.js 압축을 풀었지만 node.exe가 없습니다.' }
  Write-Host "Node.js를 준비했습니다: $Dest"
  exit 0
} catch {
  Show-Error $_.Exception.Message
  exit 3
} finally {
  Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
}
