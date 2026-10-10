# FY27 MRE survey evaluation

Evaluated 2026-10-10 as an unchanged-code baseline using the uploaded QSF, the main `config.json`, and the separate logic adapter. No survey-specific overrides or manual Word/PDF edits were applied. The Python scripts and configuration files are unchanged.

## Output and layout

The generated form contains 44 response questions across seven blocks, plus two descriptive text questions. The LibreOffice PDF has 24 Letter-size pages. All pages were visually inspected; all 44 numbered prompts were also located exactly once in the PDF text. All 11 matrices, including the two 17-statement matrices with ten scale columns, fit on single pages. Their headers are rotated and their response rows are lightly banded; all matrix statement labels were found on the same page as their question prompt.

The default layout keeps questions together, but leaves substantial whitespace on pages 3, 6, 8, 11, 16, 23, and 24. In particular, Q10, Q26, and Q43 occupy mostly empty pages before a new block. This is a useful baseline for future pagination improvements. Q41 and Q42 are ordinal flexibility ratings, but automatic scale detection renders them vertically; a future survey config could force horizontal layout. Source wording and rich text are retained, including online-only instructions in Q6 about an auto-populated MOS designation list and clicking Continue. That wording needs a paper-specific replacement before field use.

## Generated logic

Paper references use sequential numbering.

| Location | Generated direction |
| --- | --- |
| After Q7 | If the answer is anything other than Less than one year, skip to Q9 |
| Q9, No response | Skip the remaining questions in the section |
| Before Q36 | Answer only if one of the three spoon-breakage Yes responses was selected in Q35 |
| Before Q38 | Answer only if one of the three pre-use spoon-damage Yes responses was selected in Q37 |

The file has no conditional survey-flow branches. Four of its six active question display conditions are translated or covered by the generated directions. The other two are reported for review; they are not silently converted into guessed directions.

## Logic gaps exposed by this example

Q15 (QID32) and Q20 (QID35) use separate `If` and `AndIf` condition groups. The parser currently rejects `AndIf` with “Unrecognized condition group type.” The source requires at least one response in Q14/Q19, respectively, AND the corresponding N/A - None of these choice not selected. No eligibility instruction for these two questions appears in this baseline document. Supporting these exported group forms is a focused next improvement for the separate logic module.

Individual choices or matrix statements also have display conditions in seven questions. These are preserved as printed items but are not translated into respondent instructions by the current module. They are also not comprehensively enumerated in its JSON review report; this table supplements that report.

| Paper question | Source QID | Conditional choices/statements |
| --- | --- | --- |
| Q5, rank grade | QID19789337 | 24 of 24 choices; default dropdown handling produces a write-in field |
| Q13, entree ratings | QID62 | 17 of 17 statements |
| Q14, other entrees tried | QID31 | 17 of 18 choices |
| Q15, other entree ratings | QID32 | 17 of 17 statements |
| Q18, item ratings | QID33 | 13 of 13 statements |
| Q19, other items tried | QID34 | 13 of 14 choices |
| Q20, other item ratings | QID35 | 13 of 13 statements |

The online form uses these conditions to show items appropriate to prior menu/item selections. Printing every item preserves content, but does not reproduce that filtering. The existing N/A instructions help with ratings; they do not fully replace the filtering for “other items” checklists. This baseline needs those paper instructions resolved before being treated as logic-equivalent to the online survey.

## Reproduce

From the repository root:

```bash
python qsf_logic_to_paper.py examples/FY27_MRE_IOTE.qsf -o examples/FY27_MRE_IOTE-paper-logic.docx --report examples/FY27_MRE_IOTE-logic.json
```

Export the generated Word document to PDF using Word or LibreOffice. The committed PDF was exported with LibreOffice; Microsoft Word may paginate slightly differently.

The archived QSF is byte-for-byte identical to the uploaded source. SHA-256: `eac5a5296cefeb14bca7d8e42c8ce1b0b708b15d04e0455fc5e673ecb54292a0`.
