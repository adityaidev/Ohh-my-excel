import re
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

VBA_SUB_RE = re.compile(r"Sub\s+(\w+)")
VBA_FUNC_RE = re.compile(r"Function\s+(\w+)")
VBA_COMMENT_RE = re.compile(r"^\s*'")
VBA_DIM_RE = re.compile(r"Dim\s+(\w+)\s+As\s+(\w+)")
VBA_SET_RE = re.compile(r"Set\s+(\w+)\s*=")
VBA_IF_RE = re.compile(r"\bIf\b|\bElseIf\b|\bElse\b|\bEnd If\b")
VBA_FOR_RE = re.compile(r"\bFor\b|\bFor Each\b|\bNext\b")
VBA_DO_RE = re.compile(r"\bDo\b|\bLoop\b|\bWhile\b|\bUntil\b")
VBA_CALL_RE = re.compile(r"\.(\w+)\(")
VBA_RANGE_RE = re.compile(r"Range\(['\"](\w+:\w+)['\"]\)")
VBA_CELL_RE = re.compile(r"Cells?\((\d+),\s*(\d+)\)")
VBA_WS_RE = re.compile(r'Worksheets?\(["\'](\w+)["\']\)')


class VBAAnalyzer:
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.macros: list[dict] = []
        self._extract()

    def _extract(self):
        try:
            wb = load_workbook(self.file_path, keep_vba=True)
            for vba_project in (wb.vba_project or []):
                pass
            if hasattr(wb, "vbaProject") and wb.vbaProject:
                modules = dir(wb.vbaProject)
                for mod_name in modules:
                    if mod_name.startswith("_"):
                        continue
                    try:
                        code = getattr(wb.vbaProject, mod_name)
                        if isinstance(code, str):
                            analysis = self._analyze_module(mod_name, code)
                            if analysis:
                                self.macros.append(analysis)
                    except Exception:
                        pass
        except Exception:
            pass

    def _analyze_module(self, name: str, code: str) -> Optional[dict]:
        if not code.strip():
            return None
        subs = VBA_SUB_RE.findall(code)
        funcs = VBA_FUNC_RE.findall(code)
        lines = code.split("\n")
        comment_count = sum(1 for line in lines if VBA_COMMENT_RE.match(line))
        dims = VBA_DIM_RE.findall(code)
        range_refs = VBA_RANGE_RE.findall(code)
        cell_refs = VBA_CELL_RE.findall(code)
        ws_refs = VBA_WS_RE.findall(code)
        has_conditionals = bool(VBA_IF_RE.search(code))
        has_loops = bool(VBA_FOR_RE.search(code)) or bool(VBA_DO_RE.search(code))
        calls = VBA_CALL_RE.findall(code)
        summary_parts = []
        if subs:
            summary_parts.append(f"Subroutines: {', '.join(subs)}")
        if funcs:
            summary_parts.append(f"Functions: {', '.join(funcs)}")
        if dims:
            summary_parts.append(f"Variables declared: {len(dims)}")
        if range_refs:
            summary_parts.append(f"Range references: {', '.join(range_refs[:5])}")
        if ws_refs:
            summary_parts.append(f"Sheet references: {', '.join(ws_refs)}")
        if has_conditionals:
            summary_parts.append("Contains conditional logic")
        if has_loops:
            summary_parts.append("Contains loops")
        return {
            "module": name,
            "subroutines": subs,
            "functions": funcs,
            "line_count": len(lines),
            "comment_count": comment_count,
            "variables": [{"name": n, "type": t} for n, t in dims],
            "range_references": range_refs,
            "cell_references": [f"Row{r},Col{c}" for r, c in cell_refs],
            "sheet_references": ws_refs,
            "has_conditionals": has_conditionals,
            "has_loops": has_loops,
            "method_calls": list(set(calls)),
            "summary": "; ".join(summary_parts) or f"{name}: {len(lines)} lines of VBA code",
        }

    def to_dict(self) -> dict:
        return {
            "file": str(self.file_path),
            "macro_count": len(self.macros),
            "macros": self.macros,
        }


def analyze_vba(file_path: str) -> dict:
    analyzer = VBAAnalyzer(Path(file_path))
    return analyzer.to_dict()


def explain_vba(file_path: str, module_name: str = "") -> dict:
    analyzer = VBAAnalyzer(Path(file_path))
    if module_name:
        for macro in analyzer.macros:
            if macro["module"] == module_name:
                return {"explanation": macro["summary"], "details": macro}
        return {"error": f"Module '{module_name}' not found"}
    if analyzer.macros:
        combined = "; ".join(m["summary"] for m in analyzer.macros)
        return {"explanation": combined, "macro_count": len(analyzer.macros)}
    return {"explanation": "No VBA macros found", "macro_count": 0}
