"""Synthetic, local, reproducible batch check. No real workbook data."""
import json
import resource
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT / "tests")]
from test_gate import fixture, REL, DOC
from sheet_handoff_gate import audit_batch, DEFAULT_POLICY

with tempfile.TemporaryDirectory() as tmp:
    paths = []
    for i in range(20):
        path = Path(tmp) / f"{i}.xlsx"
        # 256 small, synthetic relationship parts per file; 5,120 checks per batch.
        rel = f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{DOC}/hyperlink" Target="https://example.invalid/" TargetMode="External"/></Relationships>'
        extra = {f"xl/worksheets/_rels/sheet-{j}.xml.rels": rel for j in range(256)}
        fixture(path, extra=extra)
        paths.append(path)
    times = []
    for _ in range(3):
        start = time.perf_counter()
        result = audit_batch(paths, DEFAULT_POLICY)
        times.append(time.perf_counter() - start)
        assert result["summary"] == {"pass": 0, "fail": 20, "uninspected": 0}
        assert all(r["counts"]["external_relationships"] == 256 for r in result["files"])
    print(json.dumps({"files": len(paths), "bytes": sum(p.stat().st_size for p in paths),
                      "relationships": 5120, "seconds": sorted(times),
                      "peak_rss_kib_linux": resource.getrusage(resource.RUSAGE_SELF).ru_maxrss}))
