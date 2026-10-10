# Fuel handler survey evaluation

Evaluated 2026-10-10 using the unchanged source QSF and original Word formatter, with the separate logic analyzer and Word adapter.

## Result

The reviewed form has 23 Letter-size pages, 79 response questions, and the source introduction. The response questions comprise 49 multiple-choice questions, 19 text-entry questions, and 11 matrices. Trash questions are excluded. All 25 active display conditions and both explicit end-of-block skip rules have paper instructions or are covered by generated skip routes. No active rule remains unsupported in this export.

The generated instructions consist of 11 choice-specific skips, one grouped skip after a write-in response, and six question eligibility instructions. There are no conditional survey-flow branches in this file, so it does not exercise block-level “Answer this section only if…” wording.

The source export is unchanged. `qsf_to_paper.py` and the main `config.json` are unchanged. The survey config overrides only QID85 and QID91 to render their overall like/dislike rating scales horizontally; the existing automatic scale detection rendered them vertically.

## Routing checks

References below use the paper form's sequential numbering.

| Source answer | Paper direction |
| --- | --- |
| Q8 response other than Less than one year | Skip Q9 and continue at Q10 |
| Q11 No | Skip the remainder of COVERALLS and go to HANDWEAR |
| Q25 No | Skip durability follow-ups Q26–Q28 and continue at Q29 |
| Q29 No | Skip POL follow-ups Q30–Q33 and continue at Q34 |
| Q34 No | Skip Q35 and continue at Q36 |
| Q44 No | Skip the remaining HANDWEAR questions |
| Q54 No | Skip durability follow-ups Q55–Q57 and continue at Q58 |
| Q58 No | Skip POL follow-ups Q59–Q62 and continue at Q63 |
| Q63 No | Skip Q64 and continue at Q65 |

Q21–Q24 have alternating dependencies on Q19/Q20, so their eligibility directions remain beside the affected questions. A single skip would bypass questions that some respondents still need. Q66 and Q68 require a positive response count and exclude the “None of these” choice, so they also retain eligibility wording.

Every inferred skip was checked against the exported display expressions: bypassed questions evaluate false and conditional destinations evaluate true for the controlling answer. The two explicit section exits were checked against the rules stored in block question entries.

## Source issue for review

The handwear modification follow-up Q78 (QID94, survey tag Q4.35) has an OR condition that repeats the minor-modifications answer, includes major modifications, and excludes “No, and it cannot be modified to meet my needs” from Q77. The equivalent coverall follow-up Q42 includes all three No answers. This may be intentional, but the mismatch should be reviewed. The generated form faithfully preserves it: the excluded handwear answer skips to Q79. No source rule was corrected automatically.

## Layout review and remaining refinements

All 23 rendered pages were visually checked. All 11 matrices fit on single pages; scale labels rotate where appropriate, body rows are banded, and no question prompt is separated from its response area. Circles and checkboxes render correctly, response-table paragraphs retain zero spacing, and skip directions stay in their response label cells where possible.

The layout still has room for refinement. The new-page preference for large matrices leaves substantial whitespace on pages 9 and 21. Long task lists remain vertical because their labels and accompanying Other write-ins do not suit the current short-label column rule. The eligibility wording for Q66/Q68 is accurate but verbose and would benefit from editable instruction overrides. Short write-in lines can wrap below a long prompt, as in Q9; that is readable but could be formatted more deliberately.

## Reproduce

```bash
python qsf_logic_to_paper.py examples/Extreme_Cold_Weather_Fuel_Handler_Handwear_Survey.qsf --survey-config examples/Extreme_Cold_Weather_Fuel_Handler_Handwear_Survey_config.json --report examples/Extreme_Cold_Weather_Fuel_Handler_Handwear_Survey-logic.json
```

This writes the Word form. Export it to PDF using Word or LibreOffice. The committed PDF was exported with LibreOffice and visually checked; pagination can vary slightly in Microsoft Word.
