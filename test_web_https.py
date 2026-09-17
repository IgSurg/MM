from functools import partial
import ssl
from unittest.mock import patch
from urllib.error import URLError
from urllib.request import urlopen

from create_https_cert import create_certificate
import test_web_app
from web_S_app import make_server


class HTTPSWebTests(test_web_app.WebTests):
    scheme = 'https'

    def create_server(self, root):
        cert, key = create_certificate(root / 'certs')
        self.context = ssl.create_default_context(cafile=str(cert))
        self.enterContext(patch('test_web_app.urlopen', partial(urlopen, context=self.context)))
        return make_server(root, port=0, host='0.0.0.0', certfile=cert, keyfile=key)

    def test_untrusted_certificate_rejected(self):
        with self.assertRaises(URLError):
            urlopen(self.base, timeout=5)

    def test_missing_certificate_fails(self):
        with self.assertRaises(FileNotFoundError):
            make_server(self.root, port=0, certfile=self.root / 'missing.crt', keyfile=self.root / 'missing.key')
