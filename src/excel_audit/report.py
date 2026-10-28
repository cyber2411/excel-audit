"""Render an audit result as Markdown."""

from __future__ import annotations


def _issue_table(title, issues):
    lines = [f"## {title}", ""]
    if not issues:
        lines.append("None found. ✅")
        return lines
    lines.append("| Sheet | Cell | Problem | Formula |")
    lines.append("| ----- | ---- | ------- | ------- |")
    for issue in issues:
        formula = issue["formula"].replace("|", "\\|")
        lines.append(
            f"| {issue['sheet']} | {issue['cell']} | {issue['reason']} | `{formula}` |"
        )
    return lines


def render_markdown(path, result):
    """Build the full Markdown report for one audited workbook."""
    lines = [f"# Excel audit: {path}", ""]
    broken = result["broken_refs"]
    inconsistent = result["inconsistent_columns"]
    lines.append(
        f"**Summary:** {len(broken)} broken reference(s), "
        f"{len(inconsistent)} inconsistent column(s), "
        f"{result['volatile_total']} volatile function call(s)."
    )
    lines.append("")
    lines.extend(_issue_table("Broken references", broken))
    lines.append("")
    lines.extend(_issue_table("Inconsistent columns", inconsistent))
    lines.append("")
    lines.append("## Volatile functions")
    lines.append("")
    if not result["volatile_counts"]:
        lines.append("None found. ✅")
    else:
        lines.append("| Function | Count |")
        lines.append("| -------- | ----- |")
        for name, count in sorted(result["volatile_counts"].items()):
            lines.append(f"| {name} | {count} |")
    lines.append("")
    lines.append("## Sheet stats")
    lines.append("")
    lines.append("| Sheet | Rows used | Columns used | Formulas | Volatile calls |")
    lines.append("| ----- | --------- | ------------ | -------- | -------------- |")
    for stat in result["sheet_stats"]:
        lines.append(
            f"| {stat['sheet']} | {stat['max_row']} | {stat['max_column']} | "
            f"{stat['formula_count']} | {stat['volatile_count']} |"
        )
    lines.append("")
    return "\n".join(lines)
