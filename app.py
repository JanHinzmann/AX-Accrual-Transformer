from __future__ import annotations

from pathlib import Path

import streamlit as st

from transformer import WorkbookValidationError, process_accrual_workbook


APP_DIR = Path(__file__).resolve().parent
REFERENCE_DIR = APP_DIR / "reference_files"
COMPANY_MAPPING = REFERENCE_DIR / "COMPANY MAPPING.xlsx"
AX_TEMPLATE = REFERENCE_DIR / "AX TEMPLATE.xlsx"


def _render_outcomes(result) -> None:
    generated_count = sum(item.status == "Generated" for item in result.outcomes)
    skipped_count = sum(item.status == "Skipped" for item in result.outcomes)
    failed_count = sum(item.status == "Failed" for item in result.outcomes)

    first, second, third, fourth = st.columns(4)
    first.metric("Source rows", result.source_row_count)
    second.metric("Generated files", generated_count)
    third.metric("Skipped companies", skipped_count)
    fourth.metric("Failed companies", failed_count)

    st.caption(
        f"Detected source sheet: {result.source_sheet} · Header row: {result.source_header_row}"
    )
    outcome_rows = [
        {
            "Company ID": item.company_id,
            "ERP system": item.erp_system,
            "Source rows": item.source_row_count,
            "Outcome": item.status,
            "Details": item.message,
        }
        for item in result.outcomes
    ]
    st.dataframe(outcome_rows, hide_index=True, use_container_width=True)

    if result.issues:
        with st.expander(f"Row validation issues ({len(result.issues)})", expanded=True):
            st.dataframe(
                [
                    {
                        "Company ID": issue.company_id,
                        "Source row": issue.source_row,
                        "Field": issue.field,
                        "Issue": issue.message,
                    }
                    for issue in result.issues
                ],
                hide_index=True,
                use_container_width=True,
            )


def _render_downloads(result, original_name: str) -> None:
    if not result.files:
        st.warning("No downloadable AX workbooks were generated.")
        return

    st.subheader("Downloads")
    source_stem = Path(original_name).stem or "accrual_data"
    st.download_button(
        "Download all generated files (.zip)",
        data=result.as_zip(),
        file_name=f"AX_transformed_{source_stem}.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )

    for generated in result.files:
        with st.container(border=True):
            left, right = st.columns([3, 1])
            left.write(f"**Company {generated.company_id}**")
            left.caption(generated.filename)
            right.download_button(
                "Download",
                data=generated.content,
                file_name=generated.filename,
                mime="text/csv; charset=utf-8",
                key=f"download-{generated.company_id}-{generated.filename}",
                use_container_width=True,
            )


def main() -> None:
    st.set_page_config(page_title="Accrual File Transformer", page_icon="📄", layout="wide")
    st.title("Accrual File Transformer")
    st.write(
        "Upload one accrual workbook. AX companies are split into separate CSV posting files. "
        "D365 companies are listed but are not transformed yet."
    )

    missing_references = [
        path.name for path in (COMPANY_MAPPING, AX_TEMPLATE) if not path.is_file()
    ]
    if missing_references:
        st.error(
            "The application reference files are incomplete. Missing: "
            + ", ".join(missing_references)
        )
        st.stop()

    delimiter_label = st.selectbox(
        "CSV separator",
        options=("Semicolon (;)", "Comma (,)"),
        index=0,
        help="Semicolon is recommended for German Excel installations.",
    )
    csv_delimiter = ";" if delimiter_label.startswith("Semicolon") else ","

    uploaded = st.file_uploader("Accrual data workbook", type=["xlsx"])
    if uploaded is None:
        st.info("Select an .xlsx file to begin.")
        return

    try:
        with st.spinner("Validating, splitting, and transforming the workbook..."):
            result = process_accrual_workbook(
                uploaded.getvalue(),
                original_filename=uploaded.name,
                company_mapping_path=COMPANY_MAPPING,
                ax_template_path=AX_TEMPLATE,
                csv_delimiter=csv_delimiter,
            )
    except WorkbookValidationError as exc:
        st.error(str(exc))
        return
    except Exception:
        st.error(
            "The workbook could not be processed because of an unexpected error. "
            "Check that the file is a valid, unlocked .xlsx workbook and try again."
        )
        return

    st.success("Processing finished.")
    _render_outcomes(result)
    _render_downloads(result, uploaded.name)


if __name__ == "__main__":
    main()
