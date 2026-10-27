from __future__ import annotations

from collections import defaultdict

from .constants import (
    LO_PO_PREFIXES,
    PRODUCTION_PREFIXES,
    SHEET_FIXED,
    SHEET_LO_PO,
    SHEET_PRODUCTION,
)
from .parser import Worker, id_prefix


def sheet_for_worker(worker: Worker) -> str:
    prefix = id_prefix(worker.emp_id)
    if prefix in PRODUCTION_PREFIXES:
        return SHEET_PRODUCTION
    if prefix in LO_PO_PREFIXES:
        return SHEET_LO_PO
    return SHEET_FIXED


def group_workers(workers: list[Worker]) -> dict[str, list[Worker]]:
    grouped: dict[str, list[Worker]] = defaultdict(list)
    for worker in workers:
        grouped[sheet_for_worker(worker)].append(worker)
    for sheet_workers in grouped.values():
        sheet_workers.sort(key=lambda w: (w.subsection.lower(), w.emp_id))
    return grouped
