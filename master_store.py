"""Validated, immutable snapshots and an atomic active pointer.
Set BLUECAB_DATA_DIR to a persistent disk on the production host.
The bundled source is used only until a first upload/sync is activated.
A broken active pointer fails closed; it never silently uses the bundled table.
"""
from __future__ import annotations
import json, os, shutil, tempfile, threading
from datetime import datetime
from pathlib import Path
from bluecab_engine import MasterData, DEFAULT_MASTER, MasterError, TZ

LOCK=threading.RLock()


def root():
    return Path(os.environ.get('BLUECAB_DATA_DIR',str(Path(__file__).resolve().parent/'instance'/'master')))


def active():
    pointer=root()/'current.json'
    if not pointer.exists():
        master=MasterData(DEFAULT_MASTER)
        info=json.loads((DEFAULT_MASTER.parent/'MASTER_META.json').read_text(encoding='utf-8'))
        info.update(master.summary()); info['mode']='Pakete dahil Drive kopyas\u0131'
        return master,info
    try:
        info=json.loads(pointer.read_text(encoding='utf-8'))
        digest=info['sha256']
        if len(digest)!=64 or any(c not in '0123456789abcdef' for c in digest):raise ValueError('checksum')
        master=MasterData(root()/'versions'/digest/'MASTER.xlsx')
        if master.sha256!=digest:raise ValueError('checksum')
        info.update(master.summary()); info['mode']='Y\u00fcklenmi\u015f ortak MASTER'
        return master,info
    except Exception as exc:
        raise MasterError('Etkin MASTER do\u011frulanamad\u0131. Eski tabloya d\u00f6n\u00fclmedi; i\u015flem durduruldu.') from exc


def activate(path,source,expected_version):
    candidate=MasterData(path)
    with LOCK:
        current,_=active()
        if expected_version!=current.version:
            raise MasterError('MASTER bu s\u0131rada de\u011fi\u015fti. Sayfay\u0131 yenileyip yeniden deneyin.')
        # Accidental delivery-list/partial-district uploads must not replace the company MASTER.
        missing=set(current.by_key)-set(candidate.by_key)
        if missing:
            examples=', '.join(f'{d}/{n}' for d,n in sorted(missing)[:5])
            raise MasterError(f'MASTER eksik: {len(missing)} mevcut mahalle siliniyor ({examples}). Tam dosyay\u0131 y\u00fckleyin.')
        target=root()/'versions'/candidate.sha256
        target.mkdir(parents=True,exist_ok=True)
        if not (target/'MASTER.xlsx').exists():
            temp=target/'incoming.xlsx'; shutil.copy2(path,temp); os.replace(temp,target/'MASTER.xlsx')
        info={**candidate.summary(),'source':source,'snapshot_at':datetime.now(TZ).isoformat(timespec='seconds')}
        (target/'metadata.json').write_text(json.dumps(info,ensure_ascii=False,indent=2),encoding='utf-8')
        fd,temp_name=tempfile.mkstemp(dir=root(),prefix='pointer_',suffix='.json')
        with os.fdopen(fd,'w',encoding='utf-8') as f:
            json.dump(info,f,ensure_ascii=False,indent=2); f.flush(); os.fsync(f.fileno())
        os.replace(temp_name,root()/'current.json')
        return info
