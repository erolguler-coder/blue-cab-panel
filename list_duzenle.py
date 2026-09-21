"""Both panel actions deliberately share the exact same parser and assignment engine."""
from pathlib import Path
from bluecab_engine import MasterData, DEFAULT_MASTER, iter_source_records, assign_record, ACTIVE, CLOSED
from outputs import standard_workbook


def normalize_workbook(src_path, dst_path, master_path=None):
    master=master_path if isinstance(master_path,MasterData) else MasterData(master_path or DEFAULT_MASTER)
    records=[assign_record(rec,master) for rec in iter_source_records(src_path)]
    standard_workbook(records,Path(dst_path),master)
    return {'total':len(records),'assigned':sum(r['status']==ACTIVE for r in records),
            'closed':sum(r['status']==CLOSED for r in records),
            'unresolved':sum(r['review_needed'] for r in records),'master_version':master.version}
