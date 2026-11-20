#!/usr/bin/env python3
# CRM-отправщик для Mac: сам забирает письма из очереди (Google Таблица) и отправляет через Exchange SMTP.
# Запускается автоматически через launchd (см. install.sh). Настройки: ~/.crm-mailer/config.json
import json, os, re, ssl, sys, time, base64, imaplib, smtplib, subprocess, threading, queue, random, socket, urllib.request, urllib.error
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate, make_msgid

DIR = os.path.expanduser('~/.crm-mailer')
CFG, LEDGER = os.path.join(DIR, 'config.json'), os.path.join(DIR, 'sent.json')
IDLE, BUSY, MAX_TRIES, MAX_THROTTLE, PER_CONN, BATCH = 20, 3, 3, 15, 40, 100  # пауза без работы / при работе (с); попыток при сбоях; ожиданий при ограничении сервера; писем на одно соединение

SENT_BUF = []            # письма пачки, ушедшие получателям; копии в «Отправленные» делаются позже
SENT_Q = queue.Queue()   # очередь копий: наполняется только после передачи результатов таблице
STATE, OUT = ['ожидаю письма'], []   # что делает отправщик (видно в CRM) и результаты, ещё не переданные таблице

def log(*a): print(time.strftime('%Y-%m-%d %H:%M:%S'), *a, flush=True)

class Perm(Exception): pass
class Busy(Exception): pass   # таблица занята другим запросом — не ошибка, просто повторим чуть позже

def load_cfg(): return json.load(open(CFG, encoding='utf-8'))

def api(cfg, act, d=None):
    """Запрос к таблице. Короткие сбои (таблица занята, обрыв сети) переживаем здесь же — 3 попытки, —
    чтобы CRM не показывала «отправщик не отвечает» из-за секундной заминки."""
    body = {'k': cfg['token'], 'act': act}; body.update(d or {}); last = None
    for attempt in range(3):
        t0 = time.time()
        try:
            req = urllib.request.Request(cfg['url'], data=json.dumps(body).encode('utf-8'), headers={'Content-Type': 'text/plain;charset=utf-8'})
            with urllib.request.urlopen(req, timeout=30) as r: j = json.loads(r.read().decode('utf-8'))
            if time.time() - t0 > 10: log('запрос «%s» к таблице шёл %.0f с' % (act, time.time() - t0))
            if not j.get('ok'):
                if j.get('error') == 'busy': raise Busy('таблица занята')
                raise RuntimeError(j.get('error') or 'ошибка API')
            return j
        except Busy as e: last = e; time.sleep(1.5 + random.random() * 2)
        except (urllib.error.URLError, socket.timeout, TimeoutError, ConnectionError, ValueError) as e:
            last = e; log('запрос «%s» не прошёл (попытка %d из 3): %s' % (act, attempt + 1, e)); time.sleep(2)
    raise last

def password(p):
    if p.get('password'): return p['password']
    r = subprocess.run(['security', 'find-generic-password', '-s', 'crm-mailer', '-a', p['user'], '-w'], capture_output=True, text=True)
    if r.returncode: raise Perm('Пароль не найден ни в config.json, ни в Связке ключей')
    return r.stdout.strip()

def enc_name(n):
    """Имя файла в виде RFC 2047 (=?utf-8?b?…?=): так кириллицу правильно показывают Exchange/Outlook, OWA и Почта на Mac."""
    try: n.encode('ascii'); return n
    except UnicodeEncodeError: pass
    parts, cur = [], b''
    for ch in n:
        b = ch.encode('utf-8')
        if len(cur) + len(b) > 30: parts.append(cur); cur = b''
        cur += b
    parts.append(cur)
    return ' '.join('=?utf-8?b?' + base64.b64encode(x).decode() + '?=' for x in parts)

def build(p, job):
    alt = MIMEMultipart('alternative')
    alt.attach(MIMEText(job.get('text') or ' ', 'plain', 'utf-8'))
    alt.attach(MIMEText(job.get('html') or '', 'html', 'utf-8'))
    files = job.get('files') or []
    if files:
        m = MIMEMultipart('mixed'); m.attach(alt)
        for f in files:
            mt, _, st = (f.get('type') or '').partition('/')
            if not st: mt, st = 'application', 'octet-stream'
            part = MIMEBase(mt, st); part.set_payload(base64.b64decode(f['b64'])); encoders.encode_base64(part)
            n = enc_name(f['name'])
            part.set_param('name', n); part.add_header('Content-Disposition', 'attachment', filename=n)
            m.attach(part)
    else: m = alt
    m['From'] = formataddr((p.get('from_name', ''), p['from_email']), charset='utf-8')
    if p.get('sender_email'): m['Sender'] = p['sender_email']   # для «Отправить от имени» (Send on Behalf)
    m['To'] = job['to']
    cc = [a for a in re.split(r'[;,\s]+', job.get('cc') or '') if a]
    if cc: m['Cc'] = ', '.join(cc)
    m['Subject'] = Header(job.get('subject', ''), 'utf-8')
    m['Date'] = formatdate(localtime=True)
    m['Message-ID'] = make_msgid(domain=p['from_email'].split('@')[1])
    rcpt = [job['to']] + cc + ([p['from_email']] if p.get('bcc_self') else [])
    return m, rcpt

bad_auth, conns = set(), {}

def drop(key):
    c = conns.pop(key, None)
    if c:
        try: c[0].quit()
        except Exception: pass

def close_all():
    for k in list(conns): drop(k)

def smtp_for(p):
    """Одно соединение на много писем (меньше входов на сервер), раз в PER_CONN писем — новое."""
    key = p['user'] + '|' + p['host']
    c = conns.get(key)
    if c:
        try:
            if c[1] < PER_CONN: c[0].noop(); return key, c
        except Exception: pass
        drop(key)
    ctx = ssl.create_default_context()
    if p.get('verify_tls') is False: ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
    s = smtplib.SMTP(p['host'], int(p.get('port', 587)), timeout=90)
    try: s.ehlo(); s.starttls(context=ctx); s.ehlo(); s.login(p['user'], password(p))
    except Exception:
        try: s.close()
        except Exception: pass
        raise
    conns[key] = [s, 0]
    return key, conns[key]

def send(p, job):
    key = p['user'] + '|' + p['host']
    if key in bad_auth: raise Perm('Отправка остановлена после ошибки входа — проверьте пароль и перезапустите отправщик')
    msg, rcpt = build(p, job)
    try:
        key, c = smtp_for(p)
        c[0].send_message(msg, from_addr=p.get('sender_email') or p['from_email'], to_addrs=rcpt)
        c[1] += 1
        return msg
    except smtplib.SMTPAuthenticationError as e:
        bad_auth.add(key); drop(key)   # не долбим сервер неверным паролем: во избежание блокировки учётки
        raise Perm('Ошибка входа в почту: %s' % e)
    except (smtplib.SMTPRecipientsRefused, smtplib.SMTPSenderRefused) as e:
        raise Perm('Сервер отклонил адрес: %s' % e)
    except smtplib.SMTPResponseException as e:
        if 500 <= e.smtp_code < 600: raise Perm('Сервер отклонил письмо: %s %s' % (e.smtp_code, e.smtp_error))
        drop(key); raise
    except Exception:
        drop(key); raise

def throttled(e):
    t = str(getattr(e, 'smtp_error', e)).lower()
    return getattr(e, 'smtp_code', 0) in (421, 450, 451, 452) or any(w in t for w in ('throttl', 'rate limit', 'too many', 'limit exceeded', 'try again later'))

imaps = {}

def imap_conn(p, fresh=False):
    im = p['imap']; key = im['host'] + '|' + (im.get('user') or p['user'])
    if not fresh and key in imaps:
        try: imaps[key].noop(); return imaps[key]
        except Exception: imaps.pop(key, None)
    c = imaplib.IMAP4_SSL(im['host'], int(im.get('port', 993)))
    c.login(im.get('user') or p['user'], password(p)); imaps[key] = c; return c

def save_sent(p, msg):
    """Кладёт копию в «Отправленные» (если в config.json у отправителя задан блок imap). Соединение переиспользуется; сбой не отменяет отправку."""
    if not p.get('imap'): return
    for attempt in (0, 1):
        try:
            c = imap_conn(p, fresh=attempt == 1)
            c.append(p['imap'].get('folder', 'Sent Items'), '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes()); return
        except Exception as e:
            imaps.clear()
            if attempt: log('не удалось сохранить в «Отправленные»:', e)

def folders(name):
    c = imap_conn(load_cfg()['senders'][name], fresh=True)
    for f in c.list()[1]: print(f.decode('utf-8', 'replace'))

def load_ledger():
    try:
        with open(LEDGER) as f: return json.load(f)
    except Exception: return {}

def save_ledger(l):
    l = dict(sorted(l.items(), key=lambda kv: kv[1])[-3000:])
    with open(LEDGER, 'w') as f: json.dump(l, f)

def done(cfg, job, p, ok, err=''):
    OUT.append({'id': job['id'], 'to': job.get('to'), 'subject': job.get('subject'), 'ok': ok, 'err': err[:300], 'from': (p or {}).get('from_email') or job.get('from')})

def flush(cfg):
    """Результаты уходят в таблицу пачкой — одним запросом, а не по одному на письмо."""
    while OUT:
        part = OUT[:50]
        try: api(cfg, 'outbox_done', {'items': part}); del OUT[:len(part)]
        except Exception as e: log('результаты пока не переданы (повторю):', e); return

def handle(cfg, job, ledger, tries):
    """True — письмо обработано; False — временный сбой (сеть/сервер), прекращаем пачку."""
    t0 = time.time()
    jid, p = job['id'], cfg['senders'].get(job.get('from') or 'me')
    if p and job.get('fromAddr'): p = dict(p, from_email=job['fromAddr'])   # адрес отправителя задаёт сервер: у каждого сотрудника свой, общий — один на отдел
    if jid in ledger: log('уже отправлено ранее:', jid); done(cfg, job, p, True); return True
    if not p: done(cfg, job, None, False, 'В config.json нет отправителя «%s»' % job.get('from')); return True
    try:
        sent_msg = send(p, job)
    except Perm as e:
        STATE[0] = 'ошибка: %s' % str(e)[:120]; log('ОШИБКА', job.get('to'), e); done(cfg, job, p, False, str(e)); return True
    except Exception as e:
        if throttled(e):   # сервер просит не спешить: ждём и пробуем снова, это не ошибка письма
            k = ('t', jid); tries[k] = tries.get(k, 0) + 1
            STATE[0] = 'пауза: почтовый сервер ограничивает скорость (%d из %d)' % (tries[k], MAX_THROTTLE)
            log('сервер ограничивает скорость (%d/%d), жду минуту:' % (tries[k], MAX_THROTTLE), e)
            if tries[k] >= MAX_THROTTLE: done(cfg, job, p, False, 'Сервер долго ограничивает отправку: %s' % e); return True
            flush(cfg); time.sleep(60); return False
        tries[jid] = tries.get(jid, 0) + 1
        STATE[0] = 'сбой связи с почтовым сервером (%d из %d)' % (tries[jid], MAX_TRIES)
        log('сбой (%d/%d)' % (tries[jid], MAX_TRIES), job.get('to'), e)
        if tries[jid] >= MAX_TRIES: done(cfg, job, p, False, 'Не удалось отправить: %s' % e); return True
        return False
    ledger[jid] = time.time(); save_ledger(ledger)      # сначала запоминаем, потом сообщаем — чтобы не отправить дважды
    STATE[0] = 'отправляю'
    SENT_BUF.append((p, sent_msg))   # копию сохраним после отправки всей пачки и отчёта таблице
    log('отправлено', p['from_email'], '->', job['to']); done(cfg, job, p, True)
    if p.get('max_per_minute'):   # целевая скорость: ждём только остаток интервала (время самой отправки уже учтено)
        time.sleep(max(0.0, 60.0 / max(1, int(p['max_per_minute'])) - (time.time() - t0)))
    elif p.get('delay'): time.sleep(float(p['delay']))   # без ключей max_per_minute/delay пауз между письмами нет
    return True

def release_copies():
    while SENT_BUF: SENT_Q.put(SENT_BUF.pop(0))

def sent_worker():
    while True:
        p, msg = SENT_Q.get()
        try: save_sent(p, msg)
        except Exception as e: log('«Отправленные»:', e)

def main():
    os.makedirs(DIR, exist_ok=True)
    threading.Thread(target=sent_worker, daemon=True).start()
    cfg, ledger, tries, fast, awake, fails = load_cfg(), load_ledger(), {}, 0, 0, 0
    log('отправщик запущен')
    while True:
        try:
            try: cfg = load_cfg()   # настройки перечитываются сами, перезапуск после правки config.json не нужен
            except Exception as e: log('config.json:', e)
            if not tries.get('_started'):
                api(cfg, 'outbox_requeue'); tries['_started'] = 1   # после перезапуска возвращаем «зависшие» письма в очередь
            jobs = api(cfg, 'outbox_list', {'n': BATCH, 'msg': STATE[0]})['jobs']
            if jobs and time.time() - awake > 100:   # пока идёт отправка, не даём Mac уснуть
                awake = time.time()
                try: subprocess.Popen(['caffeinate', '-i', '-t', '240'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception: pass
            for i, job in enumerate(jobs):
                if not handle(cfg, job, ledger, tries):
                    for j in jobs[i:]: api(cfg, 'outbox_requeue', {'id': j['id']})
                    break
            flush(cfg)            # 2) результаты — в таблицу
            release_copies()      # 3) и только теперь копии в «Отправленные» (фоном)
            if jobs: fast = time.time() + 120
            else: close_all(); STATE[0] = 'ожидаю письма'
            fails = 0
            if len(jobs) >= BATCH: continue   # очередь ещё большая — берём следующую пачку без паузы
        except Busy:
            fails += 1; wait = min(15, 2 * fails); log('таблица занята, повторю через %d с' % wait); time.sleep(wait); continue
        except Exception as e:
            # при сбоях не засыпаем сервер запросами: пауза растёт 15 → 30 → 60 → … до 5 минут (зависшие запросы Google сами завершаются за ≤6 минут)
            fails += 1; wait = min(60, 10 * 2 ** min(fails - 1, 3)); log('нет связи с таблицей (попытка %d, повтор через %d с): %s' % (fails, wait, e)); STATE[0] = 'нет связи с таблицей'; time.sleep(wait); continue
        time.sleep(BUSY if time.time() < fast else IDLE)

if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--folders': folders(sys.argv[2])
    else: main()
