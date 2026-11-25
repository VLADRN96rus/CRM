# МУРР — отправщик писем для Windows. Берёт письма из очереди и отправляет их через ваш установленный КЛАССИЧЕСКИЙ Outlook.
# Права администратора не нужны, пароль нигде не хранится (используется уже выполненный вход в Outlook).
# Запускается сам при входе в Windows (ярлык в «Автозагрузке»). Настройки: %LOCALAPPDATA%\MURR-mailer\config.json, журнал: mailer.log
$ErrorActionPreference = 'Stop'
$Dir = Join-Path $env:LOCALAPPDATA 'MURR-mailer'
$CfgPath = Join-Path $Dir 'config.json'; $LedgerPath = Join-Path $Dir 'sent.json'; $LogPath = Join-Path $Dir 'mailer.log'
$script:State = 'ожидаю письма'; $script:Started = $false; $script:Fast = [DateTime]::MinValue; $script:WaitSec = 0; $script:OL = $null; $script:Cfg = $null; $script:Defer = New-Object System.Collections.ArrayList; $script:Adapt = 1.0
$BATCH = 20; $IDLE = 20; $BUSY = 3; $MAX_TRIES = 3; $MAX_THROTTLE = 15

try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}   # старые Windows без TLS 1.2 по умолчанию
$script:JS = $null
try { Add-Type -AssemblyName System.Web.Extensions; $script:JS = New-Object System.Web.Script.Serialization.JavaScriptSerializer; $script:JS.MaxJsonLength = [int]::MaxValue } catch { $script:JS = $null }
function ToJson($o) { if ($script:JS) { return $script:JS.Serialize($o) } else { return ($o | ConvertTo-Json -Depth 10 -Compress) } }
function FromJson($t) { if ($script:JS) { return $script:JS.DeserializeObject($t) } else { return ($t | ConvertFrom-Json -AsHashtable) } }

function Log($m) {
    $line = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $m
    try { Write-Host $line } catch {}
    try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch {}
}
function Load-Cfg { return (FromJson ([IO.File]::ReadAllText($CfgPath, [Text.Encoding]::UTF8))) }

function Api($act, $data) {
    $body = @{ k = [string]$script:Cfg['token']; act = $act }
    if ($data) { foreach ($k in $data.Keys) { $body[$k] = $data[$k] } }
    $bytes = [Text.Encoding]::UTF8.GetBytes((ToJson $body)); $last = $null
    for ($i = 1; $i -le 3; $i++) {
        try {
            $r = Invoke-WebRequest -Uri ([string]$script:Cfg['url']) -Method Post -Body $bytes -ContentType 'text/plain; charset=utf-8' -UseBasicParsing -TimeoutSec 90
            $j = FromJson ([Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()))
            if (-not $j['ok']) {
                if ($j['error'] -eq 'auth') { throw 'AUTH' }
                if ($j['error'] -eq 'busy') { throw 'BUSY' }
                throw ('API: ' + $j['error'])
            }
            return $j
        } catch {
            $last = $_; $m = $_.Exception.Message
            if ($m -eq 'AUTH') { throw }
            if ($m -eq 'BUSY') { Start-Sleep -Seconds (2 + (Get-Random -Maximum 3)) } else { Log ('запрос «' + $act + '» не прошёл (попытка ' + $i + ' из 3): ' + $m); Start-Sleep -Seconds 2 }
        }
    }
    throw $last
}

function Load-Ledger { try { $o = FromJson ([IO.File]::ReadAllText($LedgerPath, [Text.Encoding]::UTF8)); if ($o) { return $o } } catch {}; return @{} }
function Save-Ledger($l) {
    if ($l.Count -gt 3000) { $n = @{}; foreach ($e in ($l.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 3000)) { $n[$e.Key] = $e.Value }; $l = $n }
    [IO.File]::WriteAllText($LedgerPath, (ToJson $l), (New-Object Text.UTF8Encoding($false)))
}

function Get-Outlook {
    if ($script:OL) { try { [void]$script:OL.Version; return $script:OL } catch { $script:OL = $null } }
    $script:OL = New-Object -ComObject Outlook.Application
    return $script:OL
}
function Set-Sender($ol, $m, $from) {   # свой адрес или общий ящик: учётная запись Outlook, а если её нет — «от имени» (нужно право Send As / Send on Behalf)
    foreach ($a in $ol.Session.Accounts) { if (([string]$a.SmtpAddress) -ieq $from) { $m.SendUsingAccount = $a; return } }
    $m.SentOnBehalfOfName = $from
}
function Send-Mail($job) {
    $ol = Get-Outlook; $m = $ol.CreateItem(0); $tmp = $null
    try {
        $m.To = [string]$job['to']
        $cc = ([string]$job['cc']).Trim()
        if ($cc) { $m.CC = (($cc -split '[;,\s]+') | Where-Object { $_ }) -join '; ' }
        $m.Subject = [string]$job['subject']
        $m.BodyFormat = 2
        $m.HTMLBody = [string]$job['html']
        $from = [string]$job['fromAddr']
        if ($from) { Set-Sender $ol $m $from }
        $files = $job['files']
        if ($files) {
            $tmp = Join-Path ([IO.Path]::GetTempPath()) ('murr-' + [guid]::NewGuid().ToString('N')); [void](New-Item -ItemType Directory -Path $tmp)
            foreach ($f in $files) {
                $n = ([string]$f['name']) -replace '[\\/:*?"<>|]', '_'; if (-not $n) { $n = 'file' }
                $p = Join-Path $tmp $n; [IO.File]::WriteAllBytes($p, [Convert]::FromBase64String([string]$f['b64'])); [void]$m.Attachments.Add($p)
            }
        }
        $m.Send()
    } finally {
        if ($tmp) { Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m) } catch {}
    }
}

function Is-Throttle($t) { return ($t -match '(?i)throttl|rate limit|too many|limit exceeded|quota|try again later|submission rate|configured limit|превыш|лимит') }
function Is-NoOutlook($t) { return ($t -match '(?i)80040154|class not registered|не зарегистрирован|Недопустимая строка класса') }
function Add-Result($out, $job, $ok, $err) { [void]$out.Add(@{ id = [string]$job['id']; to = [string]$job['to']; subject = [string]$job['subject']; ok = [bool]$ok; err = ([string]$err).Substring(0, [Math]::Min(300, ([string]$err).Length)); from = [string]$job['fromAddr'] }) }
function Flush-Results($out) {
    while ($out.Count -gt 0) {
        $n = [Math]::Min(50, $out.Count); $part = @($out.GetRange(0, $n))
        try { [void](Api 'outbox_done' @{ items = $part }); $out.RemoveRange(0, $n) } catch { if ($_.Exception.Message -eq 'AUTH') { throw }; Log ('результаты пока не переданы (повторю): ' + $_.Exception.Message); return }
    }
}
# ES_CONTINUOUS (+ES_SYSTEM_REQUIRED): пока идёт отправка, компьютер не засыпает
function Keep-Awake($on) { try { if (-not ('Murr.Pwr' -as [type])) { Add-Type -Namespace Murr -Name Pwr -MemberDefinition '[System.Runtime.InteropServices.DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint f);' }; [void][Murr.Pwr]::SetThreadExecutionState([uint32]$(if ($on) { 2147483649 } else { 2147483648 })) } catch {} }

# $true — письмо обработано; $false — временный сбой (остальные письма пачки вернутся в очередь)
function Handle-Job($job, $ledger, $tries, $out) {
    $id = [string]$job['id']; $t0 = Get-Date; $script:WaitSec = 0
    if ($ledger.ContainsKey($id)) { Log ('уже отправлено ранее: ' + $id); Add-Result $out $job $true ''; return $true }
    try { Send-Mail $job }
    catch {
        $msg = $_.Exception.Message
        if (Is-NoOutlook $msg) { $script:OL = $null; $script:State = 'не найден классический Outlook (возможно, включён «Новый Outlook»)'; Log $script:State; $script:WaitSec = 60; return $false }
        if (Is-Throttle $msg) {
            $k = 't' + $id; $tries[$k] = 1 + [int]$tries[$k]
            $script:State = 'пауза: сервер ограничивает скорость (' + $tries[$k] + ' из ' + $MAX_THROTTLE + ')'; Log $script:State
            if ($tries[$k] -ge $MAX_THROTTLE) { Add-Result $out $job $false ('Сервер долго ограничивает отправку: ' + $msg); return $true }
            $script:Adapt = [Math]::Min(4.0, $script:Adapt * 1.3); [void]$script:Defer.Add($job)   # письмо откладываем, остальные идут дальше
            $script:State = 'лимит сервера: отложено писем ' + $script:Defer.Count + ', добью в конце'; Start-Sleep -Seconds 30; return $true
        }
        $script:OL = $null; $tries[$id] = 1 + [int]$tries[$id]
        $script:State = 'сбой отправки через Outlook (' + $tries[$id] + ' из ' + $MAX_TRIES + ')'; Log ($script:State + ' ' + [string]$job['to'] + ': ' + $msg)
        if ($tries[$id] -ge $MAX_TRIES) { Add-Result $out $job $false ('Не удалось отправить: ' + $msg); return $true }
        $script:WaitSec = 5; return $false
    }
    $ledger[$id] = [double]([DateTimeOffset]::UtcNow.ToUnixTimeSeconds()); Save-Ledger $ledger   # сначала запоминаем, потом сообщаем: дубль исключён
    $script:State = 'отправляю'; Log ('отправлено ' + [string]$job['fromAddr'] + ' -> ' + [string]$job['to']); Add-Result $out $job $true ''
    $mpm = 0; try { $mpm = [int]$script:Cfg['max_per_minute'] } catch {}
    $script:Adapt = [Math]::Max(1.0, $script:Adapt * 0.98)
    if ($mpm -gt 0) { $left = (60.0 / $mpm) * $script:Adapt - ((Get-Date) - $t0).TotalSeconds; if ($left -gt 0) { Start-Sleep -Milliseconds ([int]($left * 1000)) } }
    return $true
}

# один проход: забрать письма, отправить, передать результаты. Возвращает число писем
function Run-Once($ledger, $tries) {
    $script:Cfg = Load-Cfg
    if (-not $script:Started) { [void](Api 'outbox_requeue' $null); $script:Started = $true }   # после перезапуска «зависшие» письма возвращаются в очередь
    $r = Api 'outbox_list' @{ n = $BATCH; msg = $script:State }
    $jobs = @($r['jobs'] | Where-Object { $_ })
    if ($jobs.Count -eq 0 -and $script:Defer.Count -gt 0) { $jobs = @($script:Defer); $script:Defer.Clear(); $script:State = 'добиваю отложенные письма (' + $jobs.Count + ')' }
    $out = New-Object System.Collections.ArrayList
    if ($jobs.Count -gt 0) { Keep-Awake $true }
    for ($i = 0; $i -lt $jobs.Count; $i++) {
        if (-not (Handle-Job $jobs[$i] $ledger $tries $out)) {
            for ($j = $i; $j -lt $jobs.Count; $j++) { [void](Api 'outbox_requeue' @{ id = [string]$jobs[$j]['id'] }) }
            Flush-Results $out
            if ($script:WaitSec -gt 0) { Start-Sleep -Seconds $script:WaitSec }
            return $jobs.Count
        }
    }
    Flush-Results $out
    if ($jobs.Count -gt 0) { $script:Fast = (Get-Date).AddSeconds(120) } else { Keep-Awake $false; $script:State = 'ожидаю письма' }
    return $jobs.Count
}

function Main {
    if (-not (Test-Path $Dir)) { [void](New-Item -ItemType Directory -Path $Dir) }
    if ((Test-Path $LogPath) -and ((Get-Item $LogPath).Length -gt 1MB)) { Move-Item $LogPath ($LogPath + '.old') -Force }
    $created = $false; $mutex = New-Object System.Threading.Mutex($true, 'Local\MURR-mailer', [ref]$created)
    if (-not $created) { return }   # уже запущен
    Log 'отправщик запущен'
    if (-not (Test-Path $CfgPath)) { Log 'нет config.json — запустите установку заново'; return }
    $ledger = Load-Ledger; $tries = @{}; $fails = 0; $iter = 0
    while ($true) {
        $iter++; if ($env:MURR_TEST_LOOPS -and $iter -gt [int]$env:MURR_TEST_LOOPS) { return }
        try { $n = Run-Once $ledger $tries; $fails = 0; if ($n -ge $BATCH) { continue } }
        catch {
            $msg = $_.Exception.Message
            if ($msg -eq 'AUTH') { $script:State = 'код доступа не принят — обратитесь к администратору'; Log $script:State; Start-Sleep -Seconds 300; continue }
            $fails++; $w = [int][Math]::Min(60, 10 * [Math]::Pow(2, [Math]::Min($fails - 1, 3)))
            $script:State = 'нет связи с таблицей'; Log ('нет связи с таблицей (попытка ' + $fails + ', повтор через ' + $w + ' с): ' + $msg); Start-Sleep -Seconds $w; continue
        }
        if ((Get-Date) -lt $script:Fast) { Start-Sleep -Seconds $BUSY } else { Start-Sleep -Seconds $IDLE }
    }
}
if (-not (Test-Path variable:MURR_NO_MAIN)) { Main }
