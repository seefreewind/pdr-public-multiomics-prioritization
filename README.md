# PDR public multi-omics candidate prioritization audit

This repository contains the reproducible manuscript-support package for a
public-data, secondary evidence audit and candidate prioritization study in
proliferative diabetic retinopathy (PDR).

The project is framed as a retrospective audit of a frozen, project-defined
430-gene candidate universe. It is not a same-patient multimodal fusion study
and does not claim independent molecular validation, clinical prediction or
treatment readiness.

## Main contents

- `outputs/jtm_manuscript_round6_discussion_enhanced.docx`: Round 6 scientific manuscript.
- `outputs/jtm_manuscript_round6_blinded.docx`: blinded manuscript version.
- `outputs/jtm_discussion_round6.docx`: standalone revised Discussion.
- `outputs/jtm_references_round6.docx`: standalone reconstructed references.
- `outputs/*.tsv`: reconstructed candidate universe, shortlist derivation and sensitivity outputs.
- `outputs/figures/`: main figure files in PDF, SVG and PNG formats.
- `reports/`: Round 6 discussion/reference audit, literature log, gene evidence map, method-reference audit and QC reports.
- `scripts/round6/build_round6_discussion_package.py`: script used to regenerate Round 6 manuscript-support files.
- `scripts/qc/check_round6_discussion_references.py`: Round 6 QC checks.

## Reproducibility boundary

The archived 430-gene universe can be reconstructed from frozen primary-ranking
outputs. The earlier upstream reduction from broader candidate exports to this
430-gene universe is not fully recoverable. The manuscript therefore describes
the work as a frozen-universe secondary evidence audit.

## Data availability

This repository contains scripts, processed summary tables, manuscript files,
figures and audit reports. It does not include restricted raw public-data
downloads, large single-cell objects or controlled-access resources.

Raw public datasets should be obtained from their original repositories using
the accessions reported in the manuscript and audit files.

## Citation and archival DOI

The manuscript-support archive corresponding to this repository is available
through Zenodo:

https://doi.org/10.5281/zenodo.21404212

Suggested data-availability wording:

`The analysis code and processed, non-restricted manuscript-support files are available at GitHub: https://github.com/seefreewind/pdr-public-multiomics-prioritization. The version corresponding to this manuscript has been archived at Zenodo: https://doi.org/10.5281/zenodo.21404212. Raw public datasets should be obtained from the original repositories using the accessions reported in the manuscript.`
