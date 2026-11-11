"""Package an immutable source candidate and test it without touching Google Cloud."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from datetime import datetime, timezone
from zipfile import ZipFile, ZIP_DEFLATED


ROOT = Path(__file__).resolve().parents[1]
ROOT_FILES = {'app.py', 'Dockerfile', 'cloudbuild.json', 'requirements.txt', 'runtime.txt', 'README.md',
              '.dockerignore', '.gcloudignore'}
SOURCE_DIRS = {'services', 'repositories', 'templates', 'static', 'tests'}
EXCLUDED = {'services/local_runtime.py', 'templates/local_outbox.html'}


def git(*args):
    return subprocess.check_output(['git', *args], cwd=ROOT).decode('utf-8')


def main():
    candidates = set(git('ls-files', '-z').split('\0'))
    candidates.update(git('ls-files', '--others', '--exclude-standard', '-z').split('\0'))
    names = sorted(name for name in candidates if name and name not in EXCLUDED and (
        name in ROOT_FILES or name.split('/')[0] in SOURCE_DIRS
        or name == 'docs/PREPARACION_PRODUCCION.md'))
    required = {'app.py', 'Dockerfile', 'requirements.txt', 'services/workspace.py',
                'services/operations.py', 'services/madrinas.py', 'static/workspace.css',
                'static/workspace.js', 'templates/_pagination.html'}
    if required - set(names):
        raise RuntimeError('Faltan archivos: ' + ', '.join(sorted(required - set(names))))
    for name in names:
        path = Path(name)
        if path.is_absolute() or '..' in path.parts or any(part.startswith('.') and part not in ROOT_FILES for part in path.parts):
            raise RuntimeError('Ruta inesperada: ' + name)
        if path.suffix.lower() in {'.env', '.db', '.dump', '.pem', '.key', '.pyc'}:
            raise RuntimeError('Archivo privado o generado: ' + name)
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
    folder = ROOT / '.local' / 'releases' / ('sig-' + stamp)
    folder.mkdir(parents=True, exist_ok=False)
    archive = folder / 'source.zip'
    files = {}
    with ZipFile(archive, 'w', compression=ZIP_DEFLATED) as package:
        for name in names:
            data = (ROOT / name).read_bytes()
            files[name] = {'sha256': hashlib.sha256(data).hexdigest(), 'bytes': len(data)}
            package.writestr(name, data)
    manifest = {'created_utc': stamp, 'base_commit': git('rev-parse', 'HEAD').strip(),
                'archive_sha256': hashlib.sha256(archive.read_bytes()).hexdigest(),
                'files': files, 'verification': 'pending', 'published': False}
    manifest_path = folder / 'MANIFEST.json'
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    source = folder / 'source'
    with ZipFile(archive) as package:
        # Every archive entry comes from the validated relative allowlist above.
        package.extractall(source)
    env = {**os.environ, 'INIT_DB': 'false', 'SECRET_KEY': 'release-test-only-secret-key', 'PYTHONPATH': ''}
    result = subprocess.run([sys.executable, '-m', 'unittest', 'discover', '-s', 'tests', '-q'],
                            cwd=source, env=env, capture_output=True, text=True, encoding='utf-8', errors='replace')
    (folder / 'tests.log').write_text(result.stdout + result.stderr, encoding='utf-8')
    manifest['verification'] = 'passed' if result.returncode == 0 else 'failed'
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'folder': str(folder), 'files': len(files),
                      'archive_sha256': manifest['archive_sha256'], 'tests': manifest['verification']}, indent=2))
    print(result.stdout + result.stderr)
    return result.returncode


if __name__ == '__main__':
    raise SystemExit(main())
