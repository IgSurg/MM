"""Создание самоподписанного сертификата для локального HTTPS через OpenSSL."""
import argparse
from ipaddress import ip_address
from pathlib import Path
import shutil
import socket
import subprocess
import sys

from web_S_app import ROOT, local_hosts


def find_openssl():
    executable = shutil.which('openssl')
    if executable:
        return executable
    bundled = Path(sys.base_prefix) / 'Library/bin/openssl.exe'
    if bundled.is_file():
        return str(bundled)
    raise FileNotFoundError('OpenSSL not found. Add it to PATH or provide an existing certificate and key.')


def create_certificate(directory, names=(), openssl=None):
    directory = Path(directory)
    cert, key = directory / 'server.crt', directory / 'server.key'
    if cert.exists() or key.exists():
        raise FileExistsError('Certificate or key already exists; choose a different --directory.')
    entries = []
    for name in sorted(local_hosts() | set(names)):
        try:
            entries.append('IP:' + str(ip_address(name)))
        except ValueError:
            if not name or any(c not in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-' for c in name):
                raise ValueError('Invalid DNS name: ' + name)
            entries.append('DNS:' + name)
    executable = openssl or find_openssl()
    directory.mkdir(parents=True, exist_ok=True)
    subprocess.run([executable, 'req', '-config', str(ROOT / 'https_openssl.cnf'),
                    '-x509', '-newkey', 'rsa:2048', '-nodes',
                    '-keyout', str(key), '-out', str(cert), '-days', '365',
                    '-subj', '/CN=' + socket.gethostname(),
                    '-addext', 'subjectAltName=' + ','.join(entries),
                    '-addext', 'basicConstraints=critical,CA:FALSE',
                    '-addext', 'extendedKeyUsage=serverAuth'],
                   check=True, capture_output=True)
    return cert, key


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Create a local self-signed HTTPS certificate')
    parser.add_argument('--directory', type=Path, default=ROOT / 'certs')
    parser.add_argument('--name', action='append', default=[], help='Additional DNS name or IP (repeatable)')
    parser.add_argument('--openssl', help='Path to openssl executable')
    args = parser.parse_args()
    try:
        cert, key = create_certificate(args.directory, args.name, args.openssl)
    except (OSError, ValueError, subprocess.CalledProcessError) as exc:
        details = exc.stderr.decode(errors='replace') if isinstance(exc, subprocess.CalledProcessError) else str(exc)
        parser.exit(1, f'Certificate creation failed: {details}\n')
    print(f'Certificate: {cert}\nPrivate key: {key}\nValid for 365 days. Self-signed: clients must trust the certificate.')
