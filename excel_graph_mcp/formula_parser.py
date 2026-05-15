import re
from typing import Optional

from formulas import Parser as FormulaParser

CELL_REF_RE = re.compile(
    r"(?:(?P<sheet>(?:'[^']+')|[^!]+)!)?"  
    r"(?P<range>\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)"
)
FUNC_RE = re.compile(r"([A-Za-z_]\w*)\(")
STRUCTURED_REF_RE = re.compile(r"(\w+)\[([^\]]+)\]")


class FormulaAnalysis:
    def __init__(self, formula_text: str):
        self.formula_text = formula_text
        self.functions_used: list[str] = []
        self.cell_references: list[dict] = []
        self.range_references: list[dict] = []
        self.structured_references: list[dict] = []
        self.nesting_depth: int = 0
        self.is_ambiguous: bool = False
        self._parse()

    def _parse(self):
        if not self.formula_text or not self.formula_text.startswith("="):
            return
        text = self.formula_text[1:]
        self.functions_used = list(set(FUNC_RE.findall(text)))
        self.structured_references = [
            {"table": m.group(1), "column": m.group(2)}
            for m in STRUCTURED_REF_RE.finditer(text)
        ]
        for m in CELL_REF_RE.finditer(text):
            ref = m.group("range")
            sheet = m.group("sheet")
            info = {"ref": ref, "sheet": sheet.strip("'") if sheet else None}
            if ":" in ref:
                info["type"] = "range"
                start, end = ref.split(":")
                info["start"] = start
                info["end"] = end
                self.range_references.append(info)
            else:
                info["type"] = "cell"
                self.cell_references.append(info)
        self._compute_nesting(text)
        self._check_ambiguous(text)

    def _compute_nesting(self, text: str):
        depth = 0
        max_depth = 0
        for ch in text:
            if ch == "(":
                depth += 1
                max_depth = max(max_depth, depth)
            elif ch == ")":
                depth -= 1
        self.nesting_depth = max_depth

    def _check_ambiguous(self, text: str):
        ambiguous_funcs = {"INDIRECT", "OFFSET", "HYPERLINK", "ADDRESS", "INDIRECT"}
        if any(f.upper() in ambiguous_funcs for f in self.functions_used):
            self.is_ambiguous = True

    @property
    def cell_count(self) -> int:
        return len(self.cell_references) + len(self.range_references)

    @property
    def all_references(self) -> list[str]:
        return [
            r["ref"] for r in self.cell_references + self.range_references
        ]


def parse_formula(formula_text: str) -> Optional[FormulaAnalysis]:
    if not formula_text or not formula_text.startswith("="):
        return None
    try:
        FormulaParser().ast(formula_text)
    except Exception:
        pass
    return FormulaAnalysis(formula_text)
