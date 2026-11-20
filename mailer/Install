#!/bin/bash
# Установка отправщика МУРР на Mac. Обычно запускается командой из МУРР (Письма → «Настроить отправку»):
#   curl -fsSL '<сайт>/mailer/install.sh' | bash -s -- '<адрес /exec>' '<код доступа>' '<сайт>/mailer/'
# Права администратора не нужны. Пароль от почты хранится в Связке ключей.
set -e
URL="$1"; TOKEN="$2"; BASE="$3"; TTY="${MURR_TTY:-/dev/tty}"
D="$HOME/.crm-mailer"; mkdir -p "$D" "$HOME/Library/LaunchAgents"
if [ -n "$BASE" ]; then curl -fsSL "${BASE}outbox.py" -o "$D/outbox.py"; else cp "$(cd "$(dirname "$0")" && pwd)/outbox.py" "$D/outbox.py"; fi
if [ ! -f "$D/config.json" ]; then
  exec 3< "$TTY"   # ввод открываем один раз: с терминала (curl | bash) или из файла (проверка)
  [ -z "$URL" ] && read -u 3 -r -p "Адрес веб-приложения (https://script.google.com/.../exec): " URL 
  [ -z "$TOKEN" ] && read -u 3 -r -p "Ваш код доступа в МУРР: " TOKEN 
  read -u 3 -r -p "Логин для почтового сервера (например, kontur/ivanov): " LOGIN 
  read -u 3 -r -s -p "Пароль от почты (хранится в Связке ключей): " PW ; echo
  read -u 3 -r -p "Как подписывать письма (Имя Фамилия): " FNAME 
  read -u 3 -r -p "Ваша рабочая почта (name@skbkontur.ru): " FMAIL 
  security add-generic-password -U -a "$LOGIN" -s crm-mailer -w "$PW"
  python3 - "$URL" "$TOKEN" "$LOGIN" "$FNAME" "$FMAIL" "$D/config.json" <<'PY'
import json, sys
url, token, login, fname, fmail, path = sys.argv[1:7]
base = {"host": "smtp.skbkontur.ru", "port": 587, "user": login, "bcc_self": False}
cfg = {"url": url, "token": token, "senders": {
  "me":     dict(base, from_email=fmail, from_name=fname, imap={"host": "imap.skbkontur.ru", "port": 993, "folder": "Sent Items"}),
  "shared": dict(base, from_email="ornp@skbkontur.ru", from_name="ОРНП")}}
json.dump(cfg, open(path, "w"), ensure_ascii=False, indent=2)
PY
  chmod 600 "$D/config.json"
fi
P="$HOME/Library/LaunchAgents/ru.crm.mailer.plist"
cat > "$P" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
<key>Label</key><string>ru.crm.mailer</string>
<key>ProgramArguments</key><array><string>/usr/bin/python3</string><string>-u</string><string>$D/outbox.py</string></array>
<key>RunAtLoad</key><true/><key>KeepAlive</key><true/>
<key>StandardOutPath</key><string>$D/mailer.log</string><key>StandardErrorPath</key><string>$D/mailer.log</string>
</dict></plist>
EOF
launchctl bootout "gui/$(id -u)" "$P" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$P"
echo "Готово. Отправщик работает в фоне и стартует вместе с Mac. Журнал: $D/mailer.log"
