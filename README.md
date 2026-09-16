# Accrual File Transformer

This Streamlit app validates one accrual-data workbook, identifies each company's ERP system, and creates a separate AX CSV file for every supported AX company. D365 companies are reported and skipped until the D365 transformation is defined.

## Run locally

From this folder:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

Upload an `.xlsx` workbook through the browser. The app keeps uploaded and generated content in memory. Download each generated AX CSV separately or download all successful CSV files as one ZIP file.

CSV files use UTF-8 with a byte-order mark for Excel compatibility. The default separator is a semicolon, and the interface also offers comma-separated output.

## Reference files

The app reads fixed files from `reference_files/`:

- `COMPANY MAPPING.xlsx` assigns each Company ID to AX or D365.
- `AX TEMPLATE.xlsx` supplies the AX column names and order.
- `ACCOUNT MAPPING.xlsx` is bundled for future D365 work and is not used by the AX transformation.

The supplied AX template contains the worksheet `Tabelle1` but no headers or data structure. For this reference file, the transformer uses the 19 required AX columns in the requested order. If the file is later replaced with a populated template containing all required headers, the generated CSV follows that header order. CSV files do not contain Excel styles, worksheet names, formulas, widths, or row heights.

## Processing behavior

- Input headers are located within the first 50 rows of every sheet.
- Header matching ignores surrounding whitespace and capitalization.
- Blank source rows are ignored.
- Company IDs are normalized without stripping meaningful leading zeroes.
- Duplicate identical company mappings are accepted; conflicting mappings fail only that company.
- AX companies with invalid, missing, or zero posting values are not exported.
- Positive values populate `Soll`; negative values populate `Haben` as positive amounts.
- `Rückbuchungsdatum` is written as `DD.MM.YYYY` for the first day of the current month in the Europe/Berlin timezone.
- D365 companies and unknown ERP systems never block valid AX outputs.

## Tests

Install the development dependencies, then run:

```powershell
python -m pytest -q
```

The tests cover amount signs, counteraccounts, posting text, company splitting, index resets, mixed ERP input, invalid mappings, blank rows, invalid values, German dates, UTF-8 encoding, CSV quoting, separator selection, template-driven header order, header discovery, and ZIP contents.
