"""xslx-a11y CLI entry point."""
import argparse
from pathlib import Path
import sys
from typing import List, Optional

from xslx_a11y.audit import audit_file, audit_result_to_json
from xslx_a11y.remediate import remediate_file
from xslx_a11y.reports.html import render_html
from xslx_a11y.reports.md import render_md
from xslx_a11y.reports.pdf import render_pdf
from xslx_a11y.reports.theme import available_themes
from xslx_a11y.triage import run_interactive_triage


def process_single_file(
    input_path: Path,
    args: argparse.Namespace,
    out_dir: Path,
) -> bool:
    """Processes a single .xlsx file: triage, audit, remediation, and reporting."""
    stem = input_path.stem
    target_path = input_path

    if args.triage:
        triaged_xlsx = Path(args.out_xlsx) if args.out_xlsx else out_dir / f"{stem}-triaged.xlsx"
        run_interactive_triage(target_path, triaged_xlsx)
        target_path = triaged_xlsx
        stem = target_path.stem

    # 1. Initial audit
    audit_before = audit_file(target_path)
    audit_after = None

    # 2. Remediate if requested
    if args.fix:
        fixed_xlsx = Path(args.out_xlsx) if args.out_xlsx else out_dir / f"{stem}-remediated.xlsx"
        if fixed_xlsx.resolve() == target_path.resolve():
            print(
                "Error: --out-xlsx cannot match input document. xslx-a11y strictly guarantees "
                "that original files remain untouched and immutable.",
                file=sys.stderr,
            )
            sys.exit(2)

        rem_res = remediate_file(target_path, fixed_xlsx)
        print(f"[Integrity Verified] Original file preserved unchanged (SHA-256: {rem_res.get('original_sha256')})")
        print(f"Remediation saved to: {fixed_xlsx}")
        for fix in rem_res.get("remediations_applied", []):
            print(f" - {fix}")
        audit_after = audit_file(fixed_xlsx)

    # 3. Render reports
    formats = [f.strip().lower() for f in args.format.split(",")]
    md_text = render_md(audit_before, after_result=audit_after, source_path=str(input_path))

    if "md" in formats:
        md_file = out_dir / f"{stem}-a11y-report.md"
        md_file.write_text(md_text, encoding="utf-8")
        print(f"Markdown report: {md_file}")

    if "json" in formats:
        json_file = out_dir / f"{stem}-audit.json"
        json_file.write_text(audit_result_to_json(audit_before), encoding="utf-8")
        print(f"JSON audit: {json_file}")

    if "html" in formats:
        html_doc = render_html(md_text, theme=args.theme)
        html_file = out_dir / f"{stem}-a11y-report.html"
        html_file.write_text(html_doc, encoding="utf-8")
        print(f"Accessible HTML report: {html_file}")

    if "pdf" in formats:
        html_doc = render_html(md_text, theme=args.theme)
        pdf_file = out_dir / f"{stem}-a11y-report.pdf"
        render_pdf(html_doc, out_path=pdf_file)
        print(f"Accessible PDF report: {pdf_file}")

    final_summary = (audit_after or audit_before)["summary"]
    return bool(final_summary["pass"])


def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser(
        prog="xslx-a11y",
        description="Audit and remediate Microsoft Excel .xlsx workbooks against WCAG 2.1 AA standards.",
    )
    parser.add_argument("file", nargs="?", default=None, help="Path to .xlsx file or folder")
    parser.add_argument("--gui", action="store_true", help="Launch graphical user interface")
    parser.add_argument(
        "--format",
        default="md",
        help="Report formats (comma-separated): md, html, pdf, json (default: md)",
    )
    parser.add_argument(
        "--theme",
        default="light",
        help=f"SMACSS theme for HTML/PDF reports: {[t['name'] for t in available_themes()]}",
    )
    parser.add_argument("--output-dir", default=".", help="Directory to save generated reports")
    parser.add_argument("--fix", action="store_true", help="Perform deterministic remediation")
    parser.add_argument("--triage", action="store_true", help="Launch interactive human triage session")
    parser.add_argument("--out-xlsx", help="Output path for remediated .xlsx file")
    parser.add_argument("--batch", action="store_true", help="Process all .xlsx files in specified directory")

    args = parser.parse_args(argv)

    if args.gui:
        try:
            from xslx_a11y.gui.app import main as gui_main
            gui_main()
            sys.exit(0)
        except (ImportError, ModuleNotFoundError) as e:
            print(f"Error: GUI dependencies not installed. Run 'pip install xslx-a11y[gui]'. ({e})", file=sys.stderr)
            sys.exit(2)

    if not args.file:
        parser.print_help(sys.stderr)
        sys.exit(2)

    input_path = Path(args.file)
    if not input_path.exists():
        print(f"Error: File or directory '{input_path}' not found.", file=sys.stderr)
        sys.exit(2)

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if input_path.is_dir() or args.batch:
        files = sorted(input_path.glob("*.xlsx")) if input_path.is_dir() else [input_path]
        if not files:
            print(f"No .xlsx files found in {input_path}", file=sys.stderr)
            sys.exit(2)
        all_passed = True
        for f in files:
            print(f"\nProcessing: {f.name}...")
            passed = process_single_file(f, args, out_dir)
            if not passed:
                all_passed = False
        sys.exit(0 if all_passed else 1)
    else:
        passed = process_single_file(input_path, args, out_dir)
        sys.exit(0 if passed else 1)


if __name__ == "__main__":
    main()
