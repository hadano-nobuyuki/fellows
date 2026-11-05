# ============================================================
#  研修環境 一括セットアップ
#   setup.bat から管理者として実行される
#   インストール済みのツールは自動でスキップする
#   -DbOnly を付けると(setup_db.bat)、データベース作成だけを行う
# ============================================================
param([switch]$DbOnly)

# ---------- 設定(資材の名前や場所が変わったらここを直す) ----------
$ToolsDir     = $PSScriptRoot                          # インストーラーを置いたフォルダ(このファイルと同じ場所)
$PleiadesRoot = 'C:\pleiades'                          # Eclipse(Pleiades)の展開先
$PgPassword   = '1234'                                 # postgres のパスワード(アプリ生成時に使う値)
$PgPort       = 5432
$SqlDir       = Join-Path $ToolsDir '商品関連テーブル作成'   # テーブル定義書(*.sql)の場所(SJIS)
$DbName       = 'myappdb'
$DbUser       = 'myapp'                                # myappdb の所有者
$DbUserPass   = '1234'

$PleiadesExe   = 'pleiades-*-java-win-64bit-jre_*.exe'
$VSCodeExe     = 'VSCodeSetup-x64-*.exe'              # 全ユーザー用(システム)インストーラー
$PythonExe     = 'python-3.*-amd64.exe'
$PostgresExe   = 'postgresql-*-windows-x64.exe'
$OdbcMsi       = 'psqlodbc_x64.msi'

# インストーラーが無いときの自動ダウンロード元(GitHub Releases / VS Code 公式)
$ReleaseBase = 'https://github.com/hadano-nobuyuki/fellows-phase1/releases/download/v1.0/'
$DownloadMap = @{
    $PleiadesExe = $ReleaseBase + 'pleiades-2026-09-java-win-64bit-jre_20261016.exe'
    $PythonExe   = $ReleaseBase + 'python-3.12.10-amd64.exe'
    $PostgresExe = $ReleaseBase + 'postgresql-18.6-5-windows-x64.exe'
    $OdbcMsi     = $ReleaseBase + 'psqlodbc_x64.msi'
    $VSCodeExe   = 'https://update.code.visualstudio.com/latest/win32-x64/stable'
}
$MaterialDir = Join-Path $ToolsDir '資材'                   # ダウンロードした資材の保存先
$DownloadDir = $MaterialDir
try {
    New-Item -ItemType Directory -Force -Path $MaterialDir | Out-Null
    $probe = Join-Path $MaterialDir '.write_test'
    Set-Content -LiteralPath $probe -Value 'x' -ErrorAction Stop
    Remove-Item -LiteralPath $probe -Force
} catch {
    # 読み取り専用の場所(サンドボックスなど)のときは一時フォルダに保存する
    $DownloadDir = Join-Path $env:TEMP 'setup_downloads'
}

# ---------- 共通処理 ----------
$ErrorActionPreference = 'Stop'

# ---------- インストール対象のユーザーを調べる ----------
# 別の管理者アカウントのパスワードを入力して実行された場合でも、
# 「実際にログインしているユーザー」の領域(VS Code の拡張機能、ログなど)に設定する
$TargetUser    = $env:USERNAME
$TargetProfile = $env:USERPROFILE
try {
    $cu = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
    if ($cu) {
        $sid = (New-Object Security.Principal.NTAccount($cu)).Translate([Security.Principal.SecurityIdentifier]).Value
        $pp  = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction Stop).ProfileImagePath
        $pp  = [Environment]::ExpandEnvironmentVariables($pp)
        if ($pp -and (Test-Path -LiteralPath $pp)) { $TargetUser = $cu; $TargetProfile = $pp }
    }
} catch { }
$SameProfile  = ($TargetProfile -ieq $env:USERPROFILE)
$TargetDesktop = Join-Path $TargetProfile 'Desktop'
if (-not (Test-Path -LiteralPath $TargetDesktop)) { $TargetDesktop = [Environment]::GetFolderPath('Desktop') }
$PublicDesktop = [Environment]::GetFolderPath('CommonDesktopDirectory')   # 全ユーザー共通のデスクトップ
$LogFile = Join-Path $TargetDesktop 'setup_log.txt'
Start-Transcript -Path $LogFile -Append | Out-Null
$Results = [ordered]@{}

function Write-Step($msg) { Write-Host ''; Write-Host "==== $msg ====" -ForegroundColor Cyan }
function Write-Info($msg) { Write-Host "  $msg" }
function Write-Warn($msg) { Write-Host "  [注意] $msg" -ForegroundColor Yellow }

function Get-Download($pattern) {
    $url = $DownloadMap[$pattern]
    if (-not $url) { return $null }
    New-Item -ItemType Directory -Force -Path $DownloadDir | Out-Null
    $name = if ($pattern -eq $VSCodeExe) { 'VSCodeSetup-x64-latest.exe' } else { [IO.Path]::GetFileName($url) }
    $dest = Join-Path $DownloadDir $name
    if (Test-Path -LiteralPath $dest) { return $dest }
    Write-Info "ダウンロード中: $name(ファイルが大きいと数分〜十数分かかります)"
    $tmp = $dest + '.part'
    try {
        $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
        if ($curl) {
            & $curl.Source -L -f --retry 3 -o $tmp $url
            if ($LASTEXITCODE -ne 0) { throw "curl 終了コード $LASTEXITCODE" }
        } else {
            Start-BitsTransfer -Source $url -Destination $tmp
        }
        Move-Item -LiteralPath $tmp -Destination $dest -Force
    } catch {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        throw "ダウンロードに失敗しました: $url ($($_.Exception.Message))"
    }
    return $dest
}

function Find-Installer($pattern) {
    $f = $ToolsDir, $MaterialDir, $DownloadDir | Select-Object -Unique | Where-Object { Test-Path -LiteralPath $_ } |
         ForEach-Object { Get-ChildItem -LiteralPath $_ -Filter $pattern -File -ErrorAction SilentlyContinue } |
         Sort-Object Name -Descending | Select-Object -First 1
    if ($f) { return $f.FullName }
    $dl = Get-Download $pattern
    if ($dl) { return $dl }
    throw "インストーラーが見つかりません: $pattern($ToolsDir)"
}

function Invoke-Installer($file, $arguments) {
    Write-Info "実行中: $(Split-Path $file -Leaf)(数分かかることがあります)"
    $p = Start-Process -FilePath $file -ArgumentList $arguments -Wait -PassThru
    if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) { throw "終了コード $($p.ExitCode)" }
}

function Update-SessionPath {
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
}

function New-Shortcut($name, $target) {
    $lnk = Join-Path $PublicDesktop "$name.lnk"      # 全ユーザー共通のデスクトップに作る
    $ws = New-Object -ComObject WScript.Shell
    $s = $ws.CreateShortcut($lnk)
    $s.TargetPath = $target
    $s.WorkingDirectory = Split-Path $target
    $s.Save()
}

# フォルダ配下からファイルを探す
# (存在しないフォルダに -Recurse を付けると親フォルダ全体を検索してしまうため、先に存在を確認する)
function Find-File($root, $name) {
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return $null }
    Get-ChildItem -LiteralPath $root -Filter $name -Recurse -Depth 4 -File -ErrorAction SilentlyContinue | Select-Object -First 1
}

function Get-InstalledApp($namePattern) {
    $keys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    Get-ItemProperty $keys -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like $namePattern } | Select-Object -First 1
}

# 各ツールを順に処理する。失敗しても次のツールへ進む
function Invoke-Tool($name, [scriptblock]$isInstalled, [scriptblock]$install) {
    Write-Step $name
    try {
        $found = & $isInstalled
        if ($found) {
            Write-Info "インストール済みのためスキップ($found)"
            $Results[$name] = 'スキップ(インストール済み)'
            return $false
        }
        & $install
        $Results[$name] = 'インストール完了'
        return $true
    } catch {
        Write-Host "  [失敗] $_" -ForegroundColor Red
        $Results[$name] = "失敗: $_"
        return $false
    }
}

Write-Host '研修環境のセットアップを開始します。完了までウィンドウを閉じないでください。' -ForegroundColor Green
Write-Info "資材フォルダ: $ToolsDir"
Write-Info "ログ: $LogFile"
Write-Info "対象ユーザー: $TargetUser"

$pgInstalledNow = $false
if (-not $DbOnly) {

# ダウンロードした資材の「インターネットから取得」の印を外し、実行時の確認画面を出さないようにする
Get-ChildItem -Path $ToolsDir -File -Recurse -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue

# ---------- 1. Eclipse(Pleiades) ----------
Invoke-Tool 'Eclipse(Pleiades)' {
    $exe = Find-File $PleiadesRoot 'eclipse.exe'
    if ($exe) { $exe.FullName }
} {
    $sfx = Find-Installer $PleiadesExe
    $dest = Join-Path $PleiadesRoot ((Split-Path $sfx -Leaf) -replace '^pleiades-(\d{4}-\d{2}).*$', '$1')
    Write-Info "展開中: $dest"
    # 自己解凍形式(7-Zip)の無人展開オプション
    Start-Process -FilePath $sfx -ArgumentList "-o`"$dest`"", '-y' -Wait | Out-Null
    $exe = Find-File $PleiadesRoot 'eclipse.exe'
    if (-not $exe) {
        Write-Warn "自動展開できませんでした。表示される画面で展開先に $PleiadesRoot を指定して展開してください。"
        Start-Process -FilePath $sfx -Wait | Out-Null
        $exe = Find-File $PleiadesRoot 'eclipse.exe'
        if (-not $exe) { throw 'eclipse.exe が見つかりません' }
    }
    New-Shortcut 'Eclipse' $exe.FullName
} | Out-Null

# ---------- 2. Python ----------
Invoke-Tool 'Python' {
    # 3.12 以上が(全ユーザー版、または対象ユーザーのユーザー版で)入っていればスキップ
    $py = Get-ChildItem 'HKLM:\SOFTWARE\Python\PythonCore', 'HKLM:\SOFTWARE\WOW6432Node\Python\PythonCore' -ErrorAction SilentlyContinue |
          Where-Object { $_.PSChildName -match '^3\.(\d+)' -and [int]$Matches[1] -ge 12 } | Select-Object -First 1
    if ($py) { "Python $($py.PSChildName)" }
    else {
        $u = Get-ChildItem (Join-Path $TargetProfile 'AppData\Local\Programs\Python') -Directory -Filter 'Python3*' -ErrorAction SilentlyContinue |
             Where-Object { [int]($_.Name -replace '\D', '') -ge 312 -and (Test-Path (Join-Path $_.FullName 'python.exe')) } | Select-Object -First 1
        if ($u) { "$($u.FullName)(ユーザー版)" }
    }
} {
    $exe = Find-Installer $PythonExe
    Invoke-Installer $exe '/quiet InstallAllUsers=1 PrependPath=1 Include_launcher=1 Include_test=0'
} | Out-Null

# ---------- 3. VS Code ----------
$codeCandidates = "$env:ProgramFiles\Microsoft VS Code\bin\code.cmd",
                  (Join-Path $TargetProfile 'AppData\Local\Programs\Microsoft VS Code\bin\code.cmd')
Invoke-Tool 'VS Code' {
    $c = $codeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($c) { $c }
} {
    $exe = Find-Installer $VSCodeExe
    # 手順書の「追加タスク」をすべて ON にした状態で、全ユーザー版としてインストール
    Invoke-Installer $exe '/VERYSILENT /NORESTART /MERGETASKS=!runcode,desktopicon,addcontextmenufiles,addcontextmenufolders,associatewithfiles,addtopath'
} | Out-Null

# 拡張機能と日本語表示は、VS Code があれば不足分だけ追加する
$code = $codeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($code) {
    Write-Step 'VS Code 拡張機能(Python・日本語化)'
    $ErrorActionPreference = 'Continue'   # code.cmd の警告出力で止まらないように
    try {
        Update-SessionPath
        # 拡張機能は、実際にログインしているユーザーの領域に入れる
        $extArgs = @()
        if (-not $SameProfile) {
            $extDir = Join-Path $TargetProfile '.vscode\extensions'
            New-Item -ItemType Directory -Path $extDir -Force | Out-Null
            $extArgs = @('--extensions-dir', $extDir)
        }
        $installed = & $code @extArgs --list-extensions 2>$null
        foreach ($ext in 'ms-python.python', 'MS-CEINTL.vscode-language-pack-ja') {
            if ($installed -contains $ext) { Write-Info "$ext はインストール済み" }
            else { Write-Info "$ext を追加中"; & $code @extArgs --install-extension $ext --force | Out-Null }
        }
        $argv = Join-Path $TargetProfile '.vscode\argv.json'
        if (-not (Test-Path $argv)) {
            New-Item -ItemType Directory -Path (Split-Path $argv) -Force | Out-Null
            # BOM 付きだと VS Code が読み込めないため、BOM なしの UTF-8 で書く
            [IO.File]::WriteAllText($argv, '{ "locale": "ja" }', (New-Object System.Text.UTF8Encoding($false)))
        } elseif (-not (Select-String -Path $argv -Pattern '"locale"' -Quiet)) {
            Write-Warn '表示言語は VS Code 起動後の「Change Language and Restart」で日本語にしてください'
        }
        $Results['VS Code 拡張機能'] = '確認済み'
    } catch {
        Write-Host "  [失敗] $_  (拡張機能はネット接続が必要です)" -ForegroundColor Red
        $Results['VS Code 拡張機能'] = "失敗: $_"
    }
    $ErrorActionPreference = 'Stop'
}

# ---------- 4. PostgreSQL ----------
$pgInstalledNow = Invoke-Tool 'PostgreSQL' {
    $svc = Get-Service -Name 'postgresql*' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($svc) { "サービス $($svc.Name)" }
} {
    $exe = Find-Installer $PostgresExe
    Invoke-Installer $exe "--mode unattended --unattendedmodeui minimal --superpassword $PgPassword --serverport $PgPort --locale C --disable-components stackbuilder"
}

}   # if (-not $DbOnly)

# bin フォルダを Path に追加(未登録の場合のみ)
$pgBin = Get-ChildItem 'C:\Program Files\PostgreSQL' -Directory -ErrorAction SilentlyContinue |
         Sort-Object { [int]($_.Name -replace '\D', '') } -Descending |
         ForEach-Object { Join-Path $_.FullName 'bin' } | Where-Object { Test-Path (Join-Path $_ 'psql.exe') } | Select-Object -First 1
if ($pgBin) {
    $machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    if (($machinePath -split ';') -notcontains $pgBin) {
        [Environment]::SetEnvironmentVariable('Path', "$machinePath;$pgBin", 'Machine')
        Write-Info "Path に追加しました: $pgBin"
    }
    Update-SessionPath
}

# ---------- 5. psqlODBC ----------
if (-not $DbOnly) {
Invoke-Tool 'psqlODBC' {
    if (Test-Path 'HKLM:\SOFTWARE\ODBC\ODBCINST.INI\PostgreSQL Unicode(x64)') { 'PostgreSQL Unicode(x64)' }
} {
    $msi = Find-Installer $OdbcMsi
    Invoke-Installer 'msiexec.exe' "/i `"$msi`" /qn /norestart"
} | Out-Null

# Excel が 32 ビット版だと x64 の ODBC ドライバは使えない
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($c2r -and $c2r.Platform -eq 'x86') {
    Write-Warn 'Excel が 32 ビット版です。psqlodbc_x86.msi を別途インストールしてください。'
}
}   # if (-not $DbOnly)

# ---------- 6. テーブル作成 ----------
Write-Step 'テーブル作成'
if ($DbOnly -and $pgBin) {
    # 既存の PostgreSQL では postgres のパスワードやポートが違うことがあるので入力してもらう
    Write-Host "データベース $DbName を作成します。" -ForegroundColor Green
    Write-Warn "既に $DbName がある場合、商品・注文などのテーブルは作り直され、データは初期状態に戻ります。"
    if ((Read-Host '続けますか? (Y/N)') -notmatch '^[Yy]') { Stop-Transcript | Out-Null; exit }
    $sec = Read-Host 'postgres ユーザーのパスワード' -AsSecureString
    $PgPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
    $port = Read-Host "ポート番号(そのまま Enter で $PgPort)"
    if ($port) { $PgPort = [int]$port }
}
if (-not ($DbOnly -or $pgInstalledNow)) {
    Write-Info 'PostgreSQL が既にあったため(または今回のインストールに失敗したため)スキップします。'
    Write-Info '同じフォルダの setup_db.bat を実行して作成してください。'
    $Results['テーブル作成'] = 'スキップ(setup_db.bat で作成)'
} elseif (-not $pgBin) {
    Write-Host '  [失敗] psql.exe が見つかりません' -ForegroundColor Red
    $Results['テーブル作成'] = '失敗: psql.exe なし'
} else {
    try {
        # 商品関連テーブル作成 に無ければ、このファイルと同じフォルダも探す
        $sql = $SqlDir, $ToolsDir | Where-Object { Test-Path -LiteralPath $_ } |
               ForEach-Object { Get-ChildItem -LiteralPath $_ -Filter '*.sql' -File } | Select-Object -First 1
        if (-not $sql) { throw "テーブル定義書(.sql)がありません: $SqlDir" }
        # 日本語のパスを psql に渡さないよう、一時フォルダにコピーして実行する
        $tmpSql = Join-Path $env:TEMP 'myappdb_tables.sql'
        Copy-Item $sql.FullName $tmpSql -Force
        $psql = Join-Path $pgBin 'psql.exe'
        $env:PGCLIENTENCODING = 'SJIS'   # テーブル定義書(.sql)は Shift_JIS で保存されている
        $env:PGOPTIONS = '-c client_min_messages=warning'   # 「テーブルが無いのでスキップ」などの NOTICE を表示しない
        $ErrorActionPreference = 'Continue'   # psql の NOTICE 出力で止まらないように

        # (1) ユーザー myapp を作成(無い場合のみ)
        $env:PGPASSWORD = $PgPassword
        Write-Info "ユーザー $DbUser を作成"
        & $psql -h localhost -p $PgPort -U postgres -d postgres -v ON_ERROR_STOP=1 -q -c "DO `$`$BEGIN IF NOT EXISTS (SELECT FROM pg_roles WHERE rolname = '$DbUser') THEN CREATE ROLE $DbUser LOGIN PASSWORD '$DbUserPass'; END IF; END`$`$;"
        if ($LASTEXITCODE -ne 0) { throw "ユーザー $DbUser の作成でエラー" }

        # (2) データベース myappdb を作成(無い場合のみ)
        $exists = & $psql -h localhost -p $PgPort -U postgres -d postgres -tAc "SELECT 1 FROM pg_database WHERE datname = '$DbName'"
        if ($exists -ne '1') {
            Write-Info "データベース $DbName を作成"
            & $psql -h localhost -p $PgPort -U postgres -d postgres -v ON_ERROR_STOP=1 -q -c "CREATE DATABASE $DbName OWNER $DbUser ENCODING 'UTF8' TEMPLATE template0;"
            if ($LASTEXITCODE -ne 0) { throw "データベース $DbName の作成でエラー" }
        }

        # (3) テーブル作成とデータ投入(SET ROLE で myapp として実行し、テーブルの所有者を myapp にする)
        Write-Info "テーブル作成: $($sql.Name)"
        & $psql -h localhost -p $PgPort -U postgres -d $DbName -v ON_ERROR_STOP=1 -q -c "SET ROLE $DbUser" -f $tmpSql
        if ($LASTEXITCODE -ne 0) { throw "$($sql.Name) でエラー(終了コード $LASTEXITCODE)" }
        Remove-Item $tmpSql -ErrorAction SilentlyContinue
        $Results['テーブル作成'] = '作成完了'
    } catch {
        Write-Host "  [失敗] $_" -ForegroundColor Red
        Write-Info 'パスワードを確認して setup_db.bat を再実行するか、手順書「_テーブル生成.txt」に従って手動で作成してください。'
        $Results['テーブル作成'] = "失敗: $_"
    } finally {
        Remove-Item Env:\PGPASSWORD -ErrorAction SilentlyContinue
        $ErrorActionPreference = 'Stop'
    }
}

# ---------- 結果 ----------
Write-Step '結果'
foreach ($k in $Results.Keys) {
    $color = if ($Results[$k] -like '失敗*') { 'Red' } elseif ($Results[$k] -like 'スキップ*') { 'Yellow' } else { 'Green' }
    Write-Host ('  {0,-20} {1}' -f $k, $Results[$k]) -ForegroundColor $color
}
Write-Host ''
Write-Host 'Path の変更を反映するため、開いているコマンドプロンプトや VS Code は一度閉じてから使ってください。' -ForegroundColor Green
Stop-Transcript | Out-Null
