# Standalone respondent logic prototype

`qsf_logic.py` analyzes a QSF and produces proposed respondent instructions. It does not modify the QSF, import the Word converter, or change Word generation. It uses only the Python standard library. This is a separate development step: review the instruction report before we connect it to document generation.

```bash
python qsf_logic.py your-survey.qsf
python qsf_logic.py your-survey.qsf -o instructions.json
python qsf_logic.py your-survey.qsf --numbering survey -o instructions.json
```

The terminal shows readable instructions and a separate manual review list. The optional JSON contains instructions keyed by stable question, choice and block IDs, with skip destinations and the questions each route bypasses. Review notes are separate from respondent instructions.

## Initial behavior

- For a simple single-choice question controlling immediately following questions in the same block, instructions are attached to the relevant source response. If Q2 requires Yes to Q1, Q1's No response receives **Skip to Q3.**
- Many responses sharing a destination can become one question-level instruction, such as **If you selected anything other than “Less than one year”, skip the remaining questions in this section.** This also works as a starting point for dropdowns converted to write-in fields.
- Several consecutive questions can be bypassed together. The analyzer evaluates subsequent conditions before choosing a destination. If a destination depends on an unknown answer, it keeps **Answer only if…** instructions rather than guessing a route.
- At the end of a block, the instruction is **Skip the remaining questions in this section.** It does not jump across a block boundary or invert block eligibility.
- Survey-flow branch conditions produce **Answer this section only if…** at the affected block. Nested supported branches are combined with AND.
- Multiple-selection controllers and nonadjacent display conditions retain question-level **Answer only if…** instructions.
- Supported comparisons are ordinary multiple-choice selections and nonselections. Unambiguous AND/OR groups can be translated. Mixed operators without clear grouping, embedded data, text/numeric comparisons, matrix-cell conditions, loops, and randomization require review. Existing Qualtrics skip logic is reported for review in this first version.
- Trash questions are excluded. Blocks and questions absent from flow are retained for review, matching the converter's content inclusion policy; their automatic routing is disabled.

## Numbering and integration

Choose `--numbering sequential` (default), `survey`, or `off` to select instruction references. Sequential numbering excludes descriptive text questions. Survey mode uses export tags, falling back to sequential references. Off uses export tags, falling back to stable QIDs. References still need to be visible or identifiable in the eventual paper document; numbering off does not itself add tags to the original converter.

This prototype does **not** load `config.json` or survey configs, including per-question numbering overrides. Select the numbering mode explicitly. Formatting settings remain entirely with the unchanged converter. Manual instruction overrides and Word insertion can be added after reviewing the analysis approach.

An unresolved review item means the instructions are incomplete, not that a route has been validated. This tool reports those items without failing the whole survey or putting technical notes in respondent instructions. It does not determine a paper equivalent for randomization, survey termination, or external data.

Run checks with:

```bash
python tests/test_logic.py
```
