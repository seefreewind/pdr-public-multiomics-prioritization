from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[2]
OUTPUTS = ROOT / "outputs"
REPORTS = ROOT / "reports"

MAIN = OUTPUTS / "jtm_manuscript_round7_submission_ready_pending_author_confirmation.docx"
FILES = [
    MAIN,
    OUTPUTS / "jtm_title_page_round7.docx",
    OUTPUTS / "jtm_declarations_round7_pending_author_confirmation.docx",
    OUTPUTS / "jtm_cover_letter_round7.docx",
    OUTPUTS / "jtm_author_confirmation_checklist_round7.docx",
]

REQUIRED_MAIN_STRINGS = {
    "author line": "Yi Zha1, Da Lin1, Ying Chen2, Yue Liu2, Yu Zhang1,*",
    "corresponding email": "zhangyu1@wzhealth.com",
    "corresponding ORCID": "0000-0001-8579-3692",
    "GitHub repository": "https://github.com/seefreewind/pdr-public-multiomics-prioritization",
    "Zenodo DOI": "https://doi.org/10.5281/zenodo.21404212",
    "data availability heading": "Availability of data and materials",
    "ethics heading": "Ethics approval and consent to participate",
    "competing interests heading": "Competing interests",
    "conservative conclusion": "do not establish causal genes, clinical biomarkers or treatment-ready interventions",
}

EXPECTED_AUTHOR_CONFIRMATIONS = {
    "individual contribution details": "please verify individual contribution details before submission",
    "funding/grant numbers": "insert grant numbers and funder names",
    "competing interests": "if accurate, replace this sentence",
}


def read_docx(path: Path) -> str:
    return "\n".join(p.text for p in Document(path).paragraphs)


def passfail(condition: bool) -> str:
    return "PASS" if condition else "FAIL"


def main() -> None:
    rows = []
    for path in FILES:
        rows.append(("file exists", path.name, passfail(path.exists())))

    main_text = read_docx(MAIN)
    for label, needle in REQUIRED_MAIN_STRINGS.items():
        rows.append(("main manuscript required content", label, passfail(needle in main_text)))

    for label, needle in EXPECTED_AUTHOR_CONFIRMATIONS.items():
        rows.append(("expected author-confirmation blocker", label, passfail(needle in main_text)))

    rows.append(
        (
            "main manuscript structure",
            "Declarations before References",
            passfail(main_text.find("Declarations") < main_text.find("References")),
        )
    )
    rows.append(
        (
            "main manuscript structure",
            "References before Figures",
            passfail(main_text.find("References") < main_text.find("Figures")),
        )
    )

    title_text = read_docx(OUTPUTS / "jtm_title_page_round7.docx")
    rows.append(("title page", "contains DOI", passfail("10.5281/zenodo.21404212" in title_text)))
    rows.append(("title page", "contains GitHub URL", passfail("github.com/seefreewind" in title_text)))

    cover_text = read_docx(OUTPUTS / "jtm_cover_letter_round7.docx")
    rows.append(("cover letter", "contains conservative scope statement", passfail("does not claim" in cover_text)))

    n_fail = sum(1 for _, _, status in rows if status == "FAIL")
    n_pass = sum(1 for _, _, status in rows if status == "PASS")

    lines = [
        "# Round7 Submission Readiness QC",
        "",
        f"Generated from `{Path(__file__).relative_to(ROOT)}`.",
        "",
        f"Summary: {n_pass} PASS, {n_fail} FAIL.",
        "",
        "| Category | Item | Status |",
        "|---|---|---:|",
    ]
    lines.extend(f"| {category} | {item} | {status} |" for category, item, status in rows)
    lines.extend(
        [
            "",
            "## Interpretation",
            "",
            "The remaining flagged items are intentional author-confirmation blockers rather than computational analysis blockers.",
            "Before journal submission, the corresponding author should confirm contribution wording, funding/grant numbers, competing interests, repository visibility and Zenodo license/access status.",
        ]
    )

    REPORTS.mkdir(exist_ok=True)
    out = REPORTS / "jtm_round7_submission_readiness_qc.md"
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(out)


if __name__ == "__main__":
    main()
