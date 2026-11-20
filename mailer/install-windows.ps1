# МУРР — установка отправщика писем на Windows (классический Outlook). Права администратора НЕ нужны.
# Запуск (обычный PowerShell): команду с вашими данными копируйте в CRM: Письма → «Настроить отправку».
$MurrDir = Join-Path $env:LOCALAPPDATA 'MURR-mailer'
$MurrStartup = [Environment]::GetFolderPath('Startup'); if (-not $MurrStartup) { $MurrStartup = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup' }
$MurrLnk = Join-Path $MurrStartup 'МУРР отправщик.lnk'

function Murr-Post($Url, $Token, $Act) {
    $js = $null; try { Add-Type -AssemblyName System.Web.Extensions; $js = New-Object System.Web.Script.Serialization.JavaScriptSerializer; $js.MaxJsonLength = [int]::MaxValue } catch {}
    $body = if ($js) { $js.Serialize(@{ k = $Token; act = $Act }) } else { (@{ k = $Token; act = $Act } | ConvertTo-Json -Compress) }
    $r = Invoke-WebRequest -Uri $Url -Method Post -Body ([Text.Encoding]::UTF8.GetBytes($body)) -ContentType 'text/plain; charset=utf-8' -UseBasicParsing -TimeoutSec 60
    $t = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
    if ($js) { return $js.DeserializeObject($t) } else { return ($t | ConvertFrom-Json -AsHashtable) }
}
function Murr-Ok($m) { Write-Host ('  [ОК] ' + $m) -ForegroundColor Green }
function Murr-Bad($m) { Write-Host ('  [!!] ' + $m) -ForegroundColor Red }
function Murr-Info($m) { Write-Host ('  ' + $m) -ForegroundColor Gray }

# Проверка: PowerShell, Outlook, связь с таблицей, адрес отправителя
function Test-MurrEnv($Url, $Token) {
    $good = $true; Write-Host ''; Write-Host 'Проверяю компьютер…' -ForegroundColor Cyan
    if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') { Murr-Bad ('Корпоративная политика ограничивает PowerShell (режим ' + $ExecutionContext.SessionState.LanguageMode + '). Этот способ здесь не сработает: напишите администратору МУРР.'); return $false }
    Murr-Ok 'PowerShell работает без ограничений'
    $accounts = @()
    try {
        $ol = New-Object -ComObject Outlook.Application
        foreach ($a in $ol.Session.Accounts) { $accounts += [string]$a.SmtpAddress }
        Murr-Ok ('Классический Outlook найден (версия ' + $ol.Version + ')')
        if ($accounts.Count) { Murr-Info ('Учётные записи Outlook: ' + ($accounts -join ', ')) }
    } catch {
        $m = $_.Exception.Message
        if ($m -match '(?i)80040154|class not registered|не зарегистрирован|Недопустимая строка класса') {
            Murr-Bad 'Классический Outlook не найден. Скорее всего, включён «Новый Outlook».'
            Murr-Info 'Что сделать: откройте Outlook и выключите переключатель «Новый Outlook» в правом верхнем углу, дождитесь перезапуска и повторите установку.'
        } else { Murr-Bad ('Не удалось подключиться к Outlook: ' + $m); Murr-Info 'Откройте Outlook, дождитесь загрузки почты и повторите установку.' }
        $good = $false
    }
    try {
        $me = $null; try { $me = Murr-Post $Url $Token 'me' } catch { throw }
        if ($me['ok']) {
            $u = $me['me']; Murr-Ok ('Связь с МУРР есть. Вы вошли как: ' + $u['name'])
            $em = [string]$u['email']
            if (-not $em) { Murr-Bad 'В профиле МУРР не указана рабочая почта (⚙ → Профиль). Без неё нельзя отправлять «от себя».'; $good = $false }
            else {
                Murr-Info ('Письма «от меня» будут уходить с адреса: ' + $em)
                if ($accounts.Count -and -not ($accounts | Where-Object { $_ -ieq $em })) { Murr-Info 'В Outlook нет учётной записи с этим адресом: письма уйдут «от имени» — нужно право Send on Behalf. Если адрес указан с ошибкой, поправьте его в ⚙ → Профиль.' }
            }
            if ($me['shared']) { Murr-Info ('Общая почта отдела: ' + $me['shared'] + ' (нужно право «Отправить как» на этот ящик)') }
        } elseif ($me['error'] -eq 'auth') { Murr-Bad 'Код доступа не принят. Скопируйте команду из МУРР заново; если не поможет — попросите администратора выдать вам новый код.'; $good = $false }
        else { Murr-Info 'Сервер таблицы ответил, но не знает команды «me»: обновите Code.gs (администратор). Установка продолжится.' }
    } catch { Murr-Bad ('Нет связи с таблицей: ' + $_.Exception.Message); $good = $false }
    return $good
}

function Stop-Murr {
    try { Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like '*MURR-mailer*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } } catch {}
}
function Murr-Args { $p = (Join-Path $MurrDir 'outbox.ps1').Replace("'", "''"); return ('-NoProfile -WindowStyle Hidden -Command "iex ([IO.File]::ReadAllText(''' + $p + ''',[Text.Encoding]::UTF8))"') }

function Install-Murr {
    param([string]$Url, [string]$Token, [string]$Base)
    $ErrorActionPreference = 'Stop'
    if ($Url -notmatch '^https://script\.google\.com/.+/exec' -or $Token.Length -lt 16 -or -not $Base) { Murr-Bad 'Команда скопирована не полностью. Скопируйте её заново из МУРР.'; return }
    if (-not (Test-MurrEnv $Url $Token)) { Write-Host ''; Murr-Bad 'Установка остановлена: исправьте пункты выше и запустите команду ещё раз.'; return }
    Write-Host ''; Write-Host 'Устанавливаю…' -ForegroundColor Cyan
    Stop-Murr
    if (-not (Test-Path $MurrDir)) { [void](New-Item -ItemType Directory -Path $MurrDir) }
    $w = New-Object Net.WebClient; $w.Encoding = [Text.Encoding]::UTF8
    $script = $w.DownloadString($Base + 'outbox.ps1')
    [IO.File]::WriteAllText((Join-Path $MurrDir 'outbox.ps1'), $script, (New-Object Text.UTF8Encoding($true)))
    $cfg = '{"url":"' + $Url + '","token":"' + $Token + '","max_per_minute":20}'
    [IO.File]::WriteAllText((Join-Path $MurrDir 'config.json'), $cfg, (New-Object Text.UTF8Encoding($false)))
    Murr-Ok ('Файлы установлены в ' + $MurrDir)
    $ws = New-Object -ComObject WScript.Shell; $l = $ws.CreateShortcut($MurrLnk)
    $l.TargetPath = (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'); $l.Arguments = (Murr-Args); $l.WindowStyle = 7; $l.Description = 'МУРР: отправка писем'; $l.Save()
    Murr-Ok 'Автозапуск при входе в Windows настроен (ярлык в «Автозагрузке»)'
    Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') -ArgumentList (Murr-Args) -WindowStyle Hidden
    Write-Host 'Запускаю и жду отклика…' -ForegroundColor Cyan
    $live = $false
    for ($i = 0; $i -lt 12 -and -not $live; $i++) {
        Start-Sleep -Seconds 3
        try { $s = Murr-Post $Url $Token 'status'; $hb = [double]$s['hb']; if ($hb -gt 0 -and (([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()) - $hb) -lt 60000) { $live = $true } } catch {}
    }
    Write-Host ''
    if ($live) { Write-Host 'ГОТОВО. Отправщик работает. Это окно можно закрыть; в МУРР индикатор станет зелёным.' -ForegroundColor Green }
    else { Murr-Bad 'Отправщик запущен, но пока не ответил. Подождите минуту и обновите страницу «Письма» в МУРР. Если не появится — откройте журнал:'; Murr-Info (Join-Path $MurrDir 'mailer.log') }
}

function Uninstall-Murr {
    Stop-Murr
    if (Test-Path $MurrLnk) { Remove-Item $MurrLnk -Force }
    if (Test-Path $MurrDir) { Remove-Item $MurrDir -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host 'Отправщик остановлен и удалён: ярлык автозапуска и папка настроек убраны.' -ForegroundColor Green
}
