"""Offline workspace backup/restore. Never include server API credentials."""
from contextlib import contextmanager
import fcntl
import hashlib
import json
from pathlib import Path
import shutil
import sqlite3
import tempfile
import zipfile


@contextmanager
def data_lock(root: Path):
    root.mkdir(parents=True, exist_ok=True)
    with (root / '.service.lock').open('a+') as handle:
        try:
            fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise RuntimeError('此数据目录已有服务运行，请勿重复启动；备份或恢复前请先停止服务') from None
        try:
            yield
        finally:
            fcntl.flock(handle, fcntl.LOCK_UN)


def backup(root: Path, target: Path):
    root, target = root.resolve(), target.resolve()
    if root in target.parents:
        raise ValueError('备份文件必须放在data目录外')
    if target.exists():
        raise ValueError('备份文件已存在，请使用新文件名')
    with data_lock(root):
        if not (root / 'zhai.sqlite3').is_file():
            raise ValueError('没有找到zhai数据，请核对数据目录')
        with sqlite3.connect(root / 'zhai.sqlite3') as db:
            if db.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
                raise ValueError('数据库校验失败，未生成备份')
        files = [p for p in sorted(root.rglob('*')) if p.name != '.service.lock']
        if any(p.is_symlink() for p in files):
            raise ValueError('数据目录包含符号链接，需人工检查后再备份')
        target.parent.mkdir(parents=True, exist_ok=True)
        manifest = {}
        try:
            with zipfile.ZipFile(target, 'x', zipfile.ZIP_DEFLATED) as archive:
                for p in files:
                    if p.is_file():
                        relative = p.relative_to(root).as_posix()
                        archive.write(p, 'data/' + relative)
                        manifest[relative] = hashlib.sha256(p.read_bytes()).hexdigest()
                archive.writestr('backup.json', json.dumps({'format': 1, 'source': str(root), 'files': manifest}))
            target.chmod(0o600)
        except Exception:
            target.unlink(missing_ok=True)
            raise


def restore(source: Path, root: Path):
    root = root.resolve()
    with data_lock(root):
        if any(p.name != '.service.lock' for p in root.iterdir()):
            raise ValueError('只允许恢复到空数据目录；原有会话不会被覆盖')
        with tempfile.TemporaryDirectory(prefix='zhai-restore-', dir=root.parent) as temp:
            stage = Path(temp) / 'data'
            stage.mkdir()
            with zipfile.ZipFile(source) as archive:
                meta = json.loads(archive.read('backup.json'))
                if meta.get('format') != 1 or not isinstance(meta.get('files'), dict):
                    raise ValueError('不支持的备份格式')
                names = archive.namelist()
                expected = {'backup.json'} | {'data/' + p for p in meta['files']}
                if len(names) != len(set(names)) or set(names) != expected:
                    raise ValueError('备份内容不符合清单')
                for name, digest in meta['files'].items():
                    path = Path(name)
                    if path.is_absolute() or '..' in path.parts or not path.parts or '.service.lock' in path.parts:
                        raise ValueError('备份路径无效')
                    payload = archive.read('data/' + name)
                    if hashlib.sha256(payload).hexdigest() != digest:
                        raise ValueError('备份文件校验失败')
                    target = stage / path
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(payload)
            old_root = str(meta['source']).rstrip('/') + '/'
            def relocate(value):
                if isinstance(value, str) and value.startswith(old_root):
                    return str(root / value[len(old_root):])
                if isinstance(value, list):
                    return [relocate(v) for v in value]
                if isinstance(value, dict):
                    return {k: relocate(v) for k, v in value.items()}
                return value
            with sqlite3.connect(stage / 'zhai.sqlite3') as db:
                if db.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
                    raise ValueError('备份数据库校验失败')
                for identifier, path in db.execute('SELECT id, server_path FROM attachments').fetchall():
                    if path and not path.startswith(old_root):
                        raise ValueError('备份引用了工作区外的文件，不能自动迁移')
                    if path and '..' in Path(path[len(old_root):]).parts:
                        raise ValueError('备份数据库路径无效')
                    if path and not (stage / path[len(old_root):]).is_file():
                        raise ValueError('备份缺少材料或成果文件')
                    db.execute('UPDATE attachments SET server_path=? WHERE id=?', (relocate(path), identifier))
                for identifier, payload in db.execute('SELECT id, payload FROM skills').fetchall():
                    db.execute('UPDATE skills SET payload=? WHERE id=?', (json.dumps(relocate(json.loads(payload)), ensure_ascii=False), identifier))
                db.execute("UPDATE sessions SET status='idle' WHERE status='running'")
            for path in stage.iterdir():
                shutil.move(str(path), str(root / path.name))
