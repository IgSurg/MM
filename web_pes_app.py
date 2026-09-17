"""Локальный интерфейс: python web_pes_app.py, http://127.0.0.1:8001."""
import argparse
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
from mm_pes import PU_BORDER, RSM_FILE, PU_FILE
import secrets
import socket
import subprocess
import sys
import threading
from urllib.parse import parse_qs, quote, unquote, urlparse
import uuid

ROOT = Path(__file__).resolve().parent
ROLES = {'rsm': ('rsm', {'.xlsx'}), 'pu': ('PU', {'.xlsx'}),
         'target': ('data_pes', {'.xlsx'})}
MAX_UPLOAD = 200 * 1024 * 1024


class App:
    def __init__(self, root=ROOT):
        self.root = Path(root)
        self.lock = threading.Lock()
        self.token = secrets.token_urlsafe(32)
        self.selected = {'rsm': self.root / RSM_FILE.relative_to(ROOT), 'pu': self.root / PU_FILE.relative_to(ROOT),
                         'target': self.root / 'data_pes/20260917_160818.xlsx'}
        self.job = {'status': 'idle', 'lines': [], 'result': None, 'error': None}
        manifest = self.root / 'web_pes_inputs.json'
        if manifest.exists():
            try:
                saved = json.loads(manifest.read_text(encoding='utf-8'))
                for role, value in saved.items():
                    if role == 'target':
                        path = (self.root / value).resolve()
                        if path.parent == (self.root / ROLES[role][0]).resolve() and path.is_file():
                            self.selected[role] = path
            except (ValueError, OSError, TypeError):
                pass

    def snapshot(self):
        with self.lock:
            files = {role: {'name': p.name, 'ready': p.is_file(),
                            'size': p.stat().st_size if p.is_file() else 0,
                            'directory': ROLES[role][0]}
                     for role, p in self.selected.items()}
            results = []
            for p in sorted((self.root / 'output_pes').glob('*.xlsx'),
                            key=lambda p: p.stat().st_mtime, reverse=True):
                if not p.name.startswith('~$'):
                    results.append({'name': p.name, 'size': p.stat().st_size,
                                    'modified': datetime.fromtimestamp(p.stat().st_mtime).isoformat(timespec='seconds'),
                                    'url': '/download/' + quote(p.name)})
            return {'files': files, 'job': dict(self.job, lines=list(self.job['lines'])),
                    'results': results, 'token': self.token}

    def start(self, border):
        with self.lock:
            if self.job['status'] == 'running':
                raise ValueError('Обработка уже выполняется.')
            if not all(p.is_file() for p in self.selected.values()):
                raise ValueError('Загрузите все три входных файла.')
            run_id = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            inputs = dict(self.selected)
            name = f"MM_{inputs['target'].stem}_{run_id}-with-CE.xlsx"
            command = [sys.executable, '-u', str(ROOT / 'mm_pes.py'),
                       '--rsm', str(inputs['rsm']), '--pu', str(inputs['pu']),
                       '--target', str(inputs['target']), '--border', str(border),
                       '--output-dir', str(self.root / 'output_pes'), '--log-dir', str(self.root / 'log_pes'),
                       '--run-id', run_id]
            self.job = {'status': 'running', 'lines': [], 'result': None, 'error': None}
            threading.Thread(target=self._worker, args=(command, name), daemon=True).start()

    def _worker(self, command, name):
        try:
            with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                  encoding='utf-8', errors='replace', cwd=self.root,
                                  creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)) as process:
                for line in process.stdout:
                    with self.lock:
                        self.job['lines'].append(line.rstrip())
                        self.job['lines'] = self.job['lines'][-100:]
                code = process.wait()
            with self.lock:
                if code == 0 and (self.root / 'output_pes' / name).is_file():
                    self.job.update(status='done', result='/download/' + quote(name))
                else:
                    self.job.update(status='error', error='Обработка не завершена. Подробности — в журнале ниже.')
        except Exception as exc:
            with self.lock:
                self.job.update(status='error', error=str(exc))


class Handler(BaseHTTPRequestHandler):
    @property
    def app(self):
        return self.server.app

    def log_message(self, *args):
        pass

    def reply(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def valid_host(self):
        value = self.headers.get('Host', '')
        try:
            parsed = urlparse('//' + value)
            if parsed.username or parsed.password or parsed.path or parsed.query or parsed.fragment:
                return False
            port = parsed.port or 80
            hosts = self.server.allowed_hosts | {self.connection.getsockname()[0].lower()}
            return port == self.server.server_port and parsed.hostname in hosts
        except ValueError:
            return False

    def do_GET(self):
        if not self.valid_host():
            return self.reply({'error': 'Недопустимый адрес сервера.'}, 403)
        route = urlparse(self.path).path
        if route == '/api/status':
            return self.reply(self.app.snapshot())
        if route == '/':
            return self.send_file(ROOT / 'web/pes.html', 'text/html; charset=utf-8')
        if route.startswith('/download/'):
            name = unquote(route[len('/download/'):])
            if '/' in name or '\\' in name or Path(name).suffix.lower() != '.xlsx' or name.startswith('~$'):
                return self.reply({'error': 'Файл недоступен.'}, 404)
            path = (self.app.root / 'output_pes' / name).resolve()
            if path.parent != (self.app.root / 'output_pes').resolve() or not path.is_file():
                return self.reply({'error': 'Файл не найден.'}, 404)
            return self.send_file(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', name)
        return self.reply({'error': 'Страница не найдена.'}, 404)

    def send_file(self, path, content_type, filename=None):
        try:
            stream = path.open('rb')
        except OSError:
            return self.reply({'error': 'Не удалось открыть файл.'}, 404)
        with stream:
            self.send_response(200)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(path.stat().st_size))
            self.send_header('X-Content-Type-Options', 'nosniff')
            if filename:
                self.send_header('Content-Disposition', "attachment; filename*=UTF-8''" + quote(filename))
            self.end_headers()
            while chunk := stream.read(1024 * 1024):
                self.wfile.write(chunk)

    def do_POST(self):
        if not self.valid_host() or self.headers.get('X-App-Token') != self.app.token:
            return self.reply({'error': 'Обновите страницу и повторите действие.'}, 403)
        try:
            size = int(self.headers.get('Content-Length', '0'))
            if not 0 < size <= MAX_UPLOAD:
                return self.reply({'error': 'Файл пуст или превышает 200 МБ.'}, 413)
            parsed = urlparse(self.path)
            if parsed.path == '/api/upload':
                return self.upload(parse_qs(parsed.query), size)
            if parsed.path == '/api/run':
                if size > 1024:
                    return self.reply({'error': 'Слишком большой запрос.'}, 413)
                border = float(json.loads(self.rfile.read(size)).get('border', PU_BORDER))
                if not 0 <= border <= 100:
                    raise ValueError('Порог должен быть от 0 до 100%.')
                self.app.start(border)
                return self.reply({'ok': True}, 202)
            return self.reply({'error': 'Неизвестное действие.'}, 404)
        except (ValueError, TypeError, AttributeError) as exc:
            return self.reply({'error': str(exc)}, 400)
        except OSError as exc:
            return self.reply({'error': f'Ошибка работы с файлом: {exc}'}, 500)

    def upload(self, query, size):
        role = query.get('role', [''])[0]
        name = query.get('name', [''])[0]
        if role not in ROLES:
            raise ValueError('Неизвестный тип файла.')
        if role != 'target':
            raise ValueError('Справочники используются напрямую из rsm и PU. Загружать можно только выгрузку PES.')
        suffix = Path(name).suffix.lower()
        if suffix not in ROLES[role][1]:
            raise ValueError('Неподдерживаемый формат файла.')
        if not name or any(c in name for c in '/\\:*?"<>|') or any(ord(c) < 32 for c in name) or len(name) > 150:
            raise ValueError('Недопустимое имя файла.')
        with self.app.lock:
            if self.app.job['status'] == 'running':
                return self.reply({'error': 'Дождитесь завершения обработки.'}, 409)
            folder = self.app.root / ROLES[role][0]
            folder.mkdir(parents=True, exist_ok=True)
            path = folder / (uuid.uuid4().hex[:8] + '_' + name)
            temporary = path.with_suffix(path.suffix + '.part')
            try:
                with temporary.open('xb') as stream:
                    remaining = size
                    while remaining:
                        chunk = self.rfile.read(min(remaining, 1024 * 1024))
                        if not chunk:
                            raise ValueError('Загрузка прервана. Выберите файл повторно.')
                        stream.write(chunk)
                        remaining -= len(chunk)
                temporary.replace(path)
                updated = dict(self.app.selected, **{role: path})
                manifest = self.app.root / 'web_pes_inputs.json'
                draft = manifest.with_suffix('.tmp')
                draft.write_text(json.dumps({k: str(v.relative_to(self.app.root)) for k, v in updated.items()},
                                            ensure_ascii=False), encoding='utf-8')
                draft.replace(manifest)
                self.app.selected = updated
                self.app.job = {'status': 'idle', 'lines': [], 'result': None, 'error': None}
            finally:
                temporary.unlink(missing_ok=True)
        return self.reply({'ok': True, 'name': path.name})


def local_hosts():
    hosts = {'127.0.0.1', 'localhost', socket.gethostname().lower(), socket.getfqdn().lower()}
    for name in tuple(hosts):
        try:
            hosts.update(info[4][0] for info in socket.getaddrinfo(name, None, socket.AF_INET))
        except socket.gaierror:
            pass
    return hosts


def make_server(root=ROOT, port=8001, host='127.0.0.1'):
    server = ThreadingHTTPServer((host, port), Handler)
    server.allowed_hosts = local_hosts() | {host.lower()}
    server.app = App(root)
    return server


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Локальная страница обработки PES')
    parser.add_argument('--port', type=int, default=8001)
    parser.add_argument('--host', default='0.0.0.0',
                        help='Адрес прослушивания: 0.0.0.0 для локальной сети, 127.0.0.1 только для этого ПК')
    args = parser.parse_args()
    server = make_server(port=args.port, host=args.host)
    print(f'Open http://127.0.0.1:{server.server_port}', flush=True)
    if args.host != '127.0.0.1':
        print(f'LAN: http://{socket.gethostname()}:{server.server_port}', flush=True)
        for address in sorted(server.allowed_hosts):
            if address not in {'0.0.0.0', '127.0.0.1'}:
                try:
                    socket.inet_aton(address)
                except OSError:
                    continue
                print(f'LAN: http://{address}:{server.server_port}', flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
