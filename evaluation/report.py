"""Markdown rendering of evaluation results. The report contains aggregate
numbers only (no invoice contents), so it is safe to publish."""

from __future__ import annotations

from .metrics import FIELD_LABELS, FieldScore


def _pct(x: float | None) -> str:
    return "–" if x is None else f"{100 * x:.1f}%"


def _secs(summary: dict | None, key: str) -> str:
    return "–" if not summary else f"{summary[key]:.2f}"


def accuracy_table(scores: dict[str, FieldScore], perfect: int, scored_docs: int) -> list[str]:
    lines = [
        "| Field | Invoices with a reviewed value | Correct | Wrong | Missing | Accuracy | Precision |",
        "|---|---:|---:|---:|---:|---:|---:|",
    ]
    for field, s in scores.items():
        if not s.total:
            continue
        lines.append(f"| {FIELD_LABELS[field]} | {s.total} | {s.correct} | {s.wrong} | {s.missing} "
                     f"| **{_pct(s.accuracy)}** | {_pct(s.precision)} |")
    if scored_docs:
        lines += ["", f"Invoices with **every** scored field correct: "
                      f"**{_pct(perfect / scored_docs)}** ({perfect} of {scored_docs})."]
    return lines


def coverage_table(fill: dict[str, int], n_docs: int) -> list[str]:
    lines = ["| Field | Filled | Fill rate |", "|---|---:|---:|"]
    for field, n in fill.items():
        lines.append(f"| {FIELD_LABELS[field]} | {n} / {n_docs} | {_pct(n / n_docs if n_docs else None)} |")
    return lines


def timing_table(rows: list[tuple[str, dict | None]]) -> list[str]:
    lines = ["| Stage | Runs | Median (s) | Mean (s) | p90 (s) | Max (s) |", "|---|---:|---:|---:|---:|---:|"]
    for label, t in rows:
        if t:
            lines.append(f"| {label} | {t['n']} | {_secs(t, 'median')} | {_secs(t, 'mean')} "
                         f"| {_secs(t, 'p90')} | {_secs(t, 'max')} |")
    return lines


def render(summary: dict) -> str:
    out = ["# Extraction evaluation", ""]
    out += [f"- Documents processed: **{summary['documents']}** "
            f"({summary['text_layer_docs']} with a text layer, {summary['ocr_docs']} via OCR)"]
    if summary.get("errors"):
        out.append(f"- Documents that failed to open: {summary['errors']}")
    if summary.get("ground_truth"):
        gt = summary["ground_truth"]
        out.append(f"- Ground truth: {gt['description']}; {gt['matched']} matched to a PDF")
    out.append(f"- Machine: {summary['machine']}")
    out.append("")

    for setup in summary["setups"]:
        out += [f"## {setup['name']}", ""]
        if "scores" in setup:
            out += accuracy_table(setup["scores"], setup["perfect"], setup["scored_docs"])
        else:
            out += ["No ground truth given: fill rate shows how often a field was extracted "
                    "at all, **not** whether it was correct.", ""]
            out += coverage_table(setup["fill"], summary["documents"])
        if setup.get("llm_calls") is not None:
            out += ["", f"LLM fallback was called on {setup['llm_calls']} of {summary['documents']} documents."]
        out.append("")

    out += ["## Speed per document", ""]
    out += timing_table(summary["timings"])
    out += ["", "Accuracy counts only invoices whose ground truth has a value for the field. "
                "*Precision* is the share of filled-in values that are correct; a missing value "
                "goes to human review, while a wrong one has to be caught by the reviewer."]
    return "\n".join(out) + "\n"
