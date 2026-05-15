import json
import shutil
import subprocess

HYPERFORMULA_SCRIPT = """
const { HyperFormula } = require('hyperformula');
const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
const sheetName = hf.addSheet('Sheet1');
const input = JSON.parse(require('fs').readFileSync('/dev/stdin', 'utf8'));
const results = {};
for (const [cellRef, formula] of Object.entries(input.formulas || {})) {
    const col = cellRef.charCodeAt(0) - 65;
    const row = parseInt(cellRef.slice(1), 10) - 1;
    hf.setCellContents({ sheet: 0, row, col }, [[formula]]);
}
for (const cellRef of input.evaluate || []) {
    const col = cellRef.charCodeAt(0) - 65;
    const row = parseInt(cellRef.slice(1), 10) - 1;
    try {
        results[cellRef] = hf.getCellValue({ sheet: 0, row, col });
    } catch (e) {
        results[cellRef] = { error: e.message };
    }
}
console.log(JSON.stringify(results));
"""


class HyperFormulaBridge:
    def __init__(self):
        self._available = shutil.which("node") is not None

    def is_available(self) -> bool:
        return self._available

    def evaluate(self, formulas: dict[str, str], evaluate: list[str]) -> dict:
        if not self._available:
            return {"error": "Node.js not found. Install Node.js for HyperFormula support, or use formulas Python lib."}
        try:
            result = subprocess.run(
                ["node", "-e", HYPERFORMULA_SCRIPT],
                input=json.dumps({"formulas": formulas, "evaluate": evaluate}),
                capture_output=True, text=True, timeout=10,
            )
            if result.returncode != 0:
                return {"error": result.stderr.strip()}
            return json.loads(result.stdout)
        except subprocess.TimeoutExpired:
            return {"error": "HyperFormula evaluation timed out"}
        except Exception as e:
            return {"error": str(e)}

    def parse_formula(self, formula_text: str) -> dict:
        if not self._available:
            return {"engine": "python", "note": "HyperFormula unavailable, use formulas lib"}
        try:
            result = subprocess.run(
                ["node", "-e", f"""
const {{ HyperFormula }} = require('hyperformula');
const hf = HyperFormula.buildEmpty({{ licenseKey: 'gpl-v3' }});
const sheetName = hf.addSheet('Sheet1');
const ast = hf._parser.parseWithCaching('={formula_text}', 0);
console.log(JSON.stringify(ast));
"""],
                capture_output=True, text=True, timeout=5,
            )
            if result.returncode == 0:
                return {"engine": "hyperformula", "ast": json.loads(result.stdout)}
        except Exception:
            pass
        return {"engine": "python", "note": "Falling back to formulas library"}


bridge = HyperFormulaBridge()


def evaluate_formulas(formulas: dict[str, str], cells: list[str]) -> dict:
    return bridge.evaluate(formulas, cells)


def parse_with_hyperformula(formula_text: str) -> dict:
    return bridge.parse_formula(formula_text)
