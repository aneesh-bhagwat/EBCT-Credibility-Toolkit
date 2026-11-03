# Excel VBA Empirical Bayes Credibility Theory (EBCT) Toolkit

A robust, modular, and production-ready VBA library designed to compute **Empirical Bayes Credibility Theory (EBCT) Model 1 and Model 2** quantities directly within Microsoft Excel. 

Built to bypass traditional VBA type-constraint errors by leveraging flexible `Variant` containers and clean matrix-handling functions, this toolkit is ideal for actuarial analysis, credibility weighting, and portfolio risk evaluation.

---

## Features

- **EBCT Model 1 Engine (`EBCT`)**: Calculates overall means ($m$), process variance ($s^2$), between-variance ($v$), and individual credibility weights ($Z$) with clean string-dispatch aliases.
- **EBCT Model 2 Engine (`EBCT2`)**: Manages weighted hierarchical data structures, computing collective volume parameters ($P_{bar}, P_{star}$), expected values of hypothetical means ($E[m(\theta)]$ and $Var(m(\theta))$), and individual row credibility vectors ($Z_i$).
- **Flexible Dispatching**: Uses case-insensitive aliases (`"M"`, `"MEAN"`, `"VAR"`, `"Z"`, `"ALL"`, etc.) so workbook formulas remain clean and intuitive.
- **Robust Error Handling**: Safely intercepts division-by-zero scenarios and returns standard Excel error codes (`#VALUE!`) rather than crashing execution.

---

## Codebase Modules

| Module Name | Description |
| :--- | :--- |
| `modHelpers.bas` | Core math utilities (`mean`, `var_s`, `rowsums`, `MatrixDivision`) and `RangeToDouble` parsing. |
| `modEBCT2Helpers.bas` | Specialized Model 2 calculations (`CalculatePbar`, `CalculatePstar`, `CalculateEstheta`, etc.). |
| `modEBCTMain.bas` | Main orchestration functions exposing `EBCT()` and `EBCT2()` to Excel cells. |

---

## Installation & Setup

1. Download the latest `.xlsm` release or clone this repository.
2. Open your target Excel workbook and press `ALT + F11` to open the Visual Basic for Applications (VBA) editor.
3. Go to **File > Import File...** and select the `.bas` files from the `src/` directory:
   - `modHelpers.bas`
   - `modEBCT2Helpers.bas`
   - `modEBCTMain.bas`
4. Ensure your workbook is saved as an **Excel Macro-Enabled Workbook (`.xlsm`)**.

---

## Usage Examples

### EBCT Model 1
Once imported, you can use the function directly inside your worksheets:
```excel
=EBCT(A1:J10, "MEAN")
=EBCT(A1:J10, "Z")
