# Qualtrics paper survey configuration

The converter runs locally with Python 3.8 or newer and `python-docx`. Existing commands still work. All defaults are built into the code. The supplied `config.json` also contains all defaults; keep it beside the Python script to customize your general preferences. JSON does not support comments.

```bash
pip install python-docx
python qsf_to_paper.py survey.qsf --survey-config surveyname_config.json
```

The default output is `survey-paper.docx`. Use `-o filename.docx` to change it. The script automatically reads `config.json` beside the script, regardless of the directory you run it from. Use `--config another_defaults.json` to choose a different main config. `--survey-config` selects a survey-specific override file; it is not automatically discovered by filename. Explicit paths are relative to your working directory.

If the main JSON is missing, built-in defaults still work. If an explicitly selected file is missing, or a config cannot be read, conversion continues with a console warning and the available lower-priority settings. Existing files containing invalid JSON or unknown settings produce a clear error so mistakes are not silently ignored.

## Priority

Settings apply in this order: built-in defaults, main config, survey config, general CLI options, question-specific overrides. Files merge recursively: a survey file needs only the settings it changes, and overrides for the same QID inherit unspecified properties from the main config. A question override wins over general CLI settings. Repeated `--set` assignments use the last value. `--set` also wins over CLI shortcuts, regardless of argument order.

## Quick CLI adjustments

```bash
python qsf_to_paper.py survey.qsf --survey-config surveyname_config.json --question-font-size 12 --response-font-size 10 --bold-questions --numbering survey
python qsf_to_paper.py survey.qsf --matrix-row-height 0.35 --no-matrix-banding --matrix-row-borders --matrix-row-border-weight 0.75
python qsf_to_paper.py survey.qsf --header-rotation never --blank-lines-before 1
```

Run `python qsf_to_paper.py --help` for shortcuts. Every setting is also available using dotted paths:

```bash
python qsf_to_paper.py survey.qsf --set matrix.header_max_height_inches=1.5 --set tables.line_spacing=1.15
python qsf_to_paper.py survey.qsf --set page.paper_size=a4 --set page.orientation=landscape
```

Values can be plain strings or JSON numbers, `true`, `false`, and `null`. Quote the whole assignment when it contains spaces, for example `--set "fonts.family=Times New Roman"`.

## Numbering and question IDs

`questions.numbering` accepts `off`, `sequential`, or `survey`. Survey numbering uses each question's `DataExportTag` exactly as exported, followed by the usual question separator. If a tag is missing, it falls back to the sequential number. `--use-export-tags` remains an alias for survey numbering.

Overrides use the stable question ID, such as `QID22`, rather than its displayed number or export tag. List the IDs, export tags, and prompts with:

```bash
python qsf_to_paper.py survey.qsf --list-questions
python qsf_logic_to_paper.py survey.qsf --list-questions
```

Both commands print the same tab-separated list: QID, export tag, and question text. Use the first column for JSON overrides. Listing exits without generating a Word document or logic report and does not load formatting configs. This list includes all question records in the export, including unused records; the converter still excludes trash questions from the document. An override referring to an absent question produces a console warning.

## Per-question overrides

Example survey-specific override file (general settings and per-question changes can be combined):

```json
{
  "questions": {"numbering": "survey"},
  "matrix": {"banded_rows": true},
  "question_overrides": {
    "QID22": {
      "questions": {"bold": true, "page_break_before": true},
      "matrix": {
        "header_rotation": "never",
        "statement_font_size": 10,
        "preferred_row_height_inches": 0.35,
        "banded_rows": false,
        "row_bottom_border": {
          "enabled": true,
          "weight_pt": 0.75,
          "color": "#999999"
        }
      }
    }
  }
}
```

You can also change an override from the CLI:

```bash
python qsf_to_paper.py survey.qsf --set question_overrides.QID22.matrix.header_rotation=always
```

Question overrides support `questions`, `responses`, `tables`, `matrix`, `write_in`, and question-level `pagination` settings. Under `fonts`, they support `family` and `preserve_rich_text`. Paper dimensions, title/section sizes, `sections`, and `new_page_per_block` are document-wide settings.

## Block sections

Block names and page breaks before later blocks are independent options, both off by default:

```json
{
  "sections": {"show_block_names": false},
  "pagination": {"new_page_per_block": false}
}
```

Set either value to `true` in your main or survey configuration. Block names use the QSF block description; empty names and Qualtrics' generic "Default Question Block" label remain suppressed. Enabling block page breaks starts each later nonempty block on a fresh page, with or without a heading; the first block does not receive an extra page break. Empty blocks and repeated flow references do not add headings or pages.

```bash
python qsf_to_paper.py survey.qsf --show-block-names --new-page-per-block
python qsf_to_paper.py survey.qsf --no-show-block-names --no-new-page-per-block
python qsf_logic_to_paper.py survey.qsf --set sections.show_block_names=true --set pagination.new_page_per_block=true
```

Both options also support `--set` and apply to the respondent-logic adapter. When headings are hidden, its cross-block skip directions refer to the next question rather than an invisible section name. Question-level page breaks and automatic large-matrix page breaks continue to apply independently.

## Setting conventions

| Group | Behavior |
|---|---|
| `fonts` | Font family, title and section sizes, and rich text preservation. Font sizes are points. |
| `sections` | Show QSF block descriptions as section headings; off by default. |
| `questions` | Prompt size, numbering, bold, blank lines before, line spacing, space after in points, and a forced page break before a particular question. Bold off preserves emphasis already present in the survey; bold on makes the entire prompt bold. |
| `responses` | Label and symbol sizes, symbols, choice layout, long-list thresholds, dropdown mode/write-in threshold/column limit, and minimum heights for response/field rows. `layout` is `auto`, `vertical`, or `horizontal`. Write-in choices stay vertical even if horizontal is requested. |
| `tables` | General response-table spacing, padding, banding, full grid borders, and row bottom borders. |
| `matrix` | Matrix-specific body spacing, padding, header size/rotation, statement size/width, row-height preferences, header fill, banding, and row bottom borders. Scale columns always have equal widths. |
| `write_in` | Inline short answers, underscore count, large-field lines/spacing, long-answer line count, and handwriting row height. Qualtrics input size can increase the line count. |
| `pagination` | New pages for later blocks, keeping whole questions together where possible, and thresholds for starting large matrices on fresh pages. QSF page breaks are always ignored. |
| `page` | Letter, A4, or Legal; portrait or landscape; margins in inches. |

Line spacing values are multiples, such as `1`, `1.15`, or `1.5`. Table paragraphs always have zero space before and after. Heights, widths, margins, and padding use inches; border weights use points. Colors use six-digit hex strings such as `#F7F7F7`.

Matrix banding and matrix row borders are independent: enable either, both, or neither. The `tables` banding/bottom-border settings control other response tables separately. `tables.borders` controls full grid borders for all tables. Handwriting fields retain their write-in rules even when decorative row borders are disabled.

Multiple-choice tables use `responses.indent_inches` (default `0.5`) from the left text margin. Their total width is reduced by the same amount so the right edge stays within the page. Set it to `0` for flush-left tables. This applies to both vertical lists and horizontal rating scales; matrix and form tables stay at the text margin. A dropdown converted to an inline write-in has no response table to indent.

Long multiple-choice lists have automatic row banding by default through `responses.large_list_banded_rows`. “Long” uses the existing `responses.multiple_columns_above` threshold: more than 10 choices by default. Banding applies whether a long list uses one vertical list or several columns, with the color from `tables.band_color`; a horizontal multiple-choice scale above the same threshold is also banded. Set `responses.large_list_banded_rows` to `false` to disable automatic long-list banding; `tables.banded_rows` can still enable general table banding independently. Yes/No choices use vertical lists in automatic layout, including Qualtrics predefined Yes/No scales; an explicit `responses.layout` override still takes priority.

Text-entry FORM tables grow toward `responses.form_preferred_row_height_inches` (default `0.45`) to provide handwriting space. If the compact form fits in the estimated remaining space on the current page, added height is limited to that space. If it already needs a fresh page, growth uses a full-page budget instead. The existing `responses.field_row_height_inches` (`0.34`) remains the minimum and continues to control compact rank and other field tables. Wrapped labels can increase the necessary height. These calculations estimate paragraph wrapping and table heights and reserve room for pagination differences; Word decides the final page placement.

```bash
python qsf_to_paper.py survey.qsf --set responses.indent_inches=0.5 --set responses.large_list_banded_rows=true --set responses.form_preferred_row_height_inches=0.5
```

All three settings support main JSON, survey JSON, CLI `--set`, and per-question overrides.

## Dropdown lists

Dropdowns use the same editable response tables, circles, indentation, label styling, and long-list banding as other single-choice questions. Short dropdowns stay vertical; genuine rating scales can still be horizontal. Long lists of short labels use multiple columns, filled down each column and then across, preserving the QSF choice order.

```json
{
  "responses": {
    "dropdown_mode": "auto",
    "dropdown_write_in_at": 50,
    "dropdown_numeric_write_in_at": 12,
    "dropdown_max_columns": 3
  }
}
```

In `auto` mode, numeric sequences such as age, height with units, or years/months of service become write-ins at `dropdown_numeric_write_in_at` choices (12 by default). Short numeric lists remain choices. Categorical responses such as military grade codes E-1/W-1/O-1 stay as choices. Other lists, including name selectors, become write-ins at `dropdown_write_in_at` choices (50 by default, replacing 12). Numeric/grade recognition requires at least 80% of labels to form an ascending or descending sequence, allowing a few special endpoints such as “Less than one year” or “Prefer not to answer.” Labels and source order are preserved; the converter does not sort them. This detection is a heuristic; unusual categories can be set explicitly.

`dropdown_mode` accepts `auto`, `choices` (always show the options), and `write_in` (always provide a write-in). It is independent of `responses.layout`, which controls the arrangement of retained choices. `dropdown_max_columns` defaults to 3 for dropdowns only; other multiple-choice lists retain `responses.max_columns` (2). The existing `multiple_columns_above` and `short_label_max_chars` settings control when columns are appropriate. Long labels and choices with accompanying text fields remain a single vertical list. Setting either column limit to 1 prevents splitting for that question type.

```bash
python qsf_to_paper.py survey.qsf --set responses.dropdown_write_in_at=75
python qsf_to_paper.py survey.qsf --set responses.dropdown_numeric_write_in_at=20
python qsf_to_paper.py survey.qsf --set responses.dropdown_max_columns=2
python qsf_to_paper.py survey.qsf --set question_overrides.QID12.responses.dropdown_mode=choices
python qsf_logic_to_paper.py survey.qsf --set question_overrides.QID20.responses.dropdown_mode=write_in
```

All dropdown settings support the main config, survey config, CLI `--set`, and per-question overrides. An existing config that explicitly sets `dropdown_write_in_at` to 12 retains that threshold for ordinary categorical lists; update it to 50 for the new default. Recognized grade sequences stay as choices. Numeric ranges use their separate threshold. To force a particular question to be a write-in, set its `dropdown_mode` to `write_in`; use `choices` to display any numeric or name list explicitly. Extremely long retained lists may require more than one printed page; conversion still preserves every choice.

## Choice-table rows and column breaks

Rank-grade lists are grouped automatically by default, with separate columns for enlisted (`E-`), warrant (`W-`, `WO-`, `CWO-`), and officer (`O-`) grades. This uses the response labels, independently of the question's QID or wording; your 24-grade list therefore has columns of 9, 5, and 10 options without any question override. Supported grade codes may also appear in longer labels, and source wording/order is retained. All labels must contain recognized grade codes, and each family must be contiguous; otherwise ordinary layout applies. `responses.group_rank_grades` defaults to `true`; set it to `false` globally or per question to disable automatic grouping. Manual `column_breaks_before` always takes priority. An explicit horizontal layout also disables automatic grouping.

`responses.max_rows` sets the maximum number of response options in each vertical column. Its default is `null` (automatic layout with no explicit row cap). Use a positive integer such as 10 in the main config, a survey config, a question override, or CLI `--set`. This controls option rows, not wrapped text lines or matrix statement rows. It applies when the response is a choice table; numeric/name dropdowns converted to write-ins keep their normal behavior unless `dropdown_mode` is set to `choices`. Horizontal rating scales retain their scale layout.

`responses.column_breaks_before` defaults to `[]` (automatic breaks). Each entry starts a new column **before** that response. Use a 1-based option position (integer) or an exact plain-text response label (string). Positions follow QSF choice order, not choice IDs or recode values. Strings such as `"10"` mean a label, not position 10. Labels must identify exactly one response. Breaks are applied in survey order; duplicate break positions are harmless. Clear inherited manual breaks with `[]`.

For E-1 through E-9, five warrant grades, and O-1 through O-10, automatic grouping already produces the requested layout. To specify the same layout manually, use this survey override:

```json
{
  "question_overrides": {
    "QID19789337": {
      "responses": {
        "dropdown_mode": "choices",
        "max_rows": 10,
        "column_breaks_before": [10, 15]
      }
    }
  }
}
```

This creates three top-aligned groups containing 9, 5, and 10 options. The overall table has 10 rows; shorter columns leave blank cells at the bottom with no response circles. A label-based equivalent is `"column_breaks_before": ["WO-1", "O-1"]` **if those labels exactly match the survey**. The supplied MRE survey uses `W-1`, so its equivalent is `["W-1", "O-1"]`. The [MRE survey config](../examples/FY27_MRE_IOTE_config.json) demonstrates a manual override using positions. Use `--list-questions` to find your own survey's QID when applying an override; automatic rank grouping needs none. Clearing manual breaks with `[]` restores automatic grouping unless `group_rank_grades` is also disabled.

```bash
python qsf_logic_to_paper.py examples/FY27_MRE_IOTE.qsf --survey-config examples/FY27_MRE_IOTE_config.json
python qsf_logic_to_paper.py survey.qsf --set question_overrides.QID12.responses.max_rows=10 --set 'question_overrides.QID12.responses.column_breaks_before=[10,15]'
```

Manual breaks take priority over automatic balancing and the usual column/label-length limits. They also make that question a vertical list if it would otherwise use a horizontal rating scale. A row cap may add columns beyond `max_columns` or `dropdown_max_columns`; if an explicit group exceeds the cap, it is split further. Column widths still respect the printable page area. When the requested columns cannot fit side by side, the converter uses stacked table panels rather than squeezing labels or dropping responses. Font size stays consistent across all options in the question. Missing, ambiguous, first-option, or out-of-range break targets produce a console warning and are ignored; the remaining choices and valid breaks are preserved. Formatting pressure does not stop generation.

## Matrix sizing

Header rotation accepts `auto`, `always`, or `never`; automatic rotation uses `rotate_at_columns`. Rotated labels are flattened to remove source line breaks, then wrapped only when needed to fit. Turning rotation off on a matrix with many narrow columns may require smaller labels or a wider page.

`null` is allowed only where it appears in the default config. It requests adaptive sizing, such as the matrix statement size, preferred row height, column width, or header height limit. A null question line spacing uses the document's normal spacing.

Matrix widths, row heights, and header height limits are preferences, not clipping limits. The converter may adjust widths, reduce fonts down to the specified minimums, increase heights, or allow an oversized matrix to continue with a repeated header. Very wide matrices may become multiple column panels. Formatting pressure should not stop conversion; any relaxed sizing is reported in the console.

## Generate a fresh default config

```bash
python qsf_to_paper.py --write-default-config config.json
```

Unknown keys, invalid values, incompatible minimum font sizes, and margins leaving insufficient space produce a clear error before conversion. Logic display is not part of this configuration version.
