import io
import json
from pathlib import Path
import tempfile
import threading
import time
import unittest
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen
import socket

import pandas as pd

from mm_pes import OUTPUT_COLUMN_ORDER
from web_pes_app import App, make_server


class WebTests(unittest.TestCase):
    scheme = 'http'

    def create_server(self, root):
        return make_server(root, port=0, host='0.0.0.0')

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.server = self.create_server(self.root)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.base = f'{self.scheme}://127.0.0.1:{self.server.server_port}'
        self.token = self.server.app.token

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()
        self.temp.cleanup()

    def request(self, path, body=None, token=True):
        headers = {'X-App-Token': self.token} if token else {}
        return urlopen(Request(self.base + path, data=body, headers=headers), timeout=10)

    def status(self):
        with self.request('/api/status') as response:
            return json.load(response)

    def upload(self, role, frame, suffix='.xlsx'):
        if role != 'target':
            path = self.server.app.selected[role]
            path.parent.mkdir(parents=True, exist_ok=True)
            frame.to_excel(path, index=False)
            return
        buffer = io.BytesIO()
        if suffix == '.xlsx':
            frame.to_excel(buffer, index=False)
        else:
            buffer.write(frame.to_csv(index=False).encode())
        with self.request('/api/upload?role=' + role + '&name=' + quote('Тест' + suffix), buffer.getvalue()) as response:
            self.assertEqual(response.status, 200)

    def wait_for_job(self):
        deadline = time.monotonic() + 30
        while time.monotonic() < deadline:
            state = self.status()
            if state['job']['status'] != 'running':
                return state
            time.sleep(.1)
        self.fail('Processing did not finish')

    def test_upload_process_download_and_validation_error(self):
        self.upload('rsm', pd.DataFrame({'Сервер Имя': ['a', 'b', 'c'],
                                       'Сервер IP address': ['10.0.0.1', '10.0.0.2', '10.0.0.3'],
                                       'Сервер КЭ': ['CE1', 'CE2', 'CE3']}))
        self.upload('pu', pd.DataFrame({'Сервер КЭ': ['CE1', 'CE2', 'CE3'],
                                      'Производственная утилизация': [.45, .46, .2]}))
        self.upload('target', pd.DataFrame({'node_name': ['a', 'a', 'b', 'c'],
                                          'node_state': ['["maintenance"]', '["maintenance"]', '["disconnected"]', '["active"]'],
                                          'max_cpu': [10, 20, 30, 40],
                                          'cpu_2sigma_usage': [1, 2, 3, 4]}))
        state = self.status()
        self.assertTrue(all(f['ready'] for f in state['files'].values()))
        for role, directory in [('rsm', 'rsm'), ('pu', 'PU'), ('target', 'data_pes')]:
            self.assertTrue((self.root / directory / state['files'][role]['name']).is_file())
        self.assertEqual(App(self.root).selected, self.server.app.selected)
        with self.request('/api/run', json.dumps({'border': 45}).encode()) as response:
            self.assertEqual(response.status, 202)
        with self.assertRaises(HTTPError) as raised:
            self.request('/api/run', b'{"border":45}')
        self.assertEqual(raised.exception.code, 400)
        state = self.wait_for_job()
        self.assertEqual(state['job']['status'], 'done', state['job'])
        with self.request(state['job']['result']) as response:
            self.assertIn('attachment', response.headers['Content-Disposition'])
            result = pd.read_excel(io.BytesIO(response.read()))
        self.assertEqual(result.columns.tolist(), OUTPUT_COLUMN_ORDER)
        self.assertEqual(result['КЭ'].tolist(), ['CE1'])
        before = pd.read_excel(next((self.root / 'log_pes').glob('*.xlsx')))
        self.assertEqual(len(before), 2)
        self.assertTrue(next((self.root / 'output_pes').glob('*.xlsx')).name.startswith('MM_'))
        self.assertFalse((self.root / 'web_inputs.json').exists())
        self.assertFalse((self.root / 'output').exists())
        self.assertFalse((self.root / 'log').exists())
        self.upload('target', pd.DataFrame({'wrong_column': [1]}), '.xlsx')
        with self.request('/api/run', b'{"border":45}'):
            pass
        state = self.wait_for_job()
        self.assertEqual(state['job']['status'], 'error')
        self.assertIsNone(state['job']['result'])

    def test_rejects_invalid_requests(self):
        for path, body, token, code in [
            ('/api/upload?role=rsm&name=a.xlsx', b'x', False, 403),
            ('/api/upload?role=rsm&name=a.xlsx', b'x', True, 400),
            ('/api/upload?role=pu&name=a.xlsx', b'x', True, 400),
            ('/api/upload?role=rsm&name=../a.xlsx', b'x', True, 400),
            ('/api/upload?role=rsm&name=a.exe', b'x', True, 400),
            ('/api/upload?role=rsm&name=a.xlsx', b'', True, 413),
            ('/api/run', b'{"border":101}', True, 400),
            ('/api/run', b'{"border":45}', True, 400),
            ('/download/..%2Frequirements.txt', None, True, 404),
        ]:
            with self.subTest(path=path), self.assertRaises(HTTPError) as raised:
                self.request(path, body, token)
            self.assertEqual(raised.exception.code, code)

    def test_page_and_empty_results(self):
        with self.request('/') as response:
            page = response.read().decode('utf-8')
        self.assertIn('Входные файлы', page)
        self.assertIn('Готовые отчёты', page)
        self.assertEqual(self.status()['results'], [])

    def test_stale_reference_paths_are_ignored(self):
        folder = self.root / 'data_pes'
        folder.mkdir()
        for name in ['old_rsm.xlsx', 'old_pu.xlsx']:
            (folder / name).write_bytes(b'old')
        (self.root / 'web_pes_inputs.json').write_text(json.dumps({
            'rsm': 'data_pes/old_rsm.xlsx', 'pu': 'data_pes/old_pu.xlsx'}), encoding='utf-8')
        app = App(self.root)
        self.assertEqual(app.selected['rsm'].parent, self.root / 'rsm')
        self.assertEqual(app.selected['pu'].parent, self.root / 'PU')

    def test_lan_hostname_and_rejected_host(self):
        self.assertEqual(self.server.server_address[0], '0.0.0.0')
        host = f'{socket.gethostname()}:{self.server.server_port}'
        with urlopen(Request(self.base + '/api/status', headers={'Host': host}), timeout=10) as response:
            self.assertEqual(response.status, 200)
        for host in ['unrelated.invalid', f'localhost:1', f'localhost:{self.server.server_port}/bad']:
            with self.subTest(host=host), self.assertRaises(HTTPError) as raised:
                urlopen(Request(self.base + '/api/status', headers={'Host': host}), timeout=10)
            self.assertEqual(raised.exception.code, 403)


if __name__ == '__main__':
    unittest.main()
