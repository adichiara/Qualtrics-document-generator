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
```

This list includes all question records in the export, including unused records; the converter still excludes trash questions from the document. An override referring to an absent question produces a console warning.

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

Question overrides support `questions`, `responses`, `tables`, `matrix`, `write_in`, and question-level `pagination` settings. Under `fonts`, they support `family` and `preserve_rich_text`. Paper dimensions, title/section sizes, and `new_page_per_block` are document-wide settings.

## Setting conventions

| Group | Behavior |
|---|---|
| `fonts` | Font family, title and section sizes, and rich text preservation. Font sizes are points. |
| `questions` | Prompt size, numbering, bold, blank lines before, line spacing, space after in points, and a forced page break before a particular question. Bold off preserves emphasis already present in the survey; bold on makes the entire prompt bold. |
| `responses` | Label and symbol sizes, symbols, choice layout, long-list thresholds, dropdown write-in threshold, and minimum heights for response/field rows. `layout` is `auto`, `vertical`, or `horizontal`. Write-in choices stay vertical even if horizontal is requested. |
| `tables` | General response-table spacing, padding, banding, full grid borders, and row bottom borders. |
| `matrix` | Matrix-specific body spacing, padding, header size/rotation, statement size/width, row-height preferences, header fill, banding, and row bottom borders. Scale columns always have equal widths. |
| `write_in` | Inline short answers, underscore count, large-field lines/spacing, long-answer line count, and handwriting row height. Qualtrics input size can increase the line count. |
| `pagination` | New pages for later blocks, keeping whole questions together where possible, and thresholds for starting large matrices on fresh pages. QSF page breaks are always ignored. |
| `page` | Letter, A4, or Legal; portrait or landscape; margins in inches. |

Line spacing values are multiples, such as `1`, `1.15`, or `1.5`. Table paragraphs always have zero space before and after. Heights, widths, margins, and padding use inches; border weights use points. Colors use six-digit hex strings such as `#F7F7F7`.

Matrix banding and matrix row borders are independent: enable either, both, or neither. The `tables` banding/bottom-border settings control other response tables separately. `tables.borders` controls full grid borders for all tables. Handwriting fields retain their write-in rules even when decorative row borders are disabled.

Header rotation accepts `auto`, `always`, or `never`; automatic rotation uses `rotate_at_columns`. Rotated labels are flattened to remove source line breaks, then wrapped only when needed to fit. Turning rotation off on a matrix with many narrow columns may require smaller labels or a wider page.

`null` is allowed only where it appears in the default config. It requests adaptive sizing, such as the matrix statement size, preferred row height, column width, or header height limit. A null question line spacing uses the document's normal spacing.

Matrix widths, row heights, and header height limits are preferences, not clipping limits. The converter may adjust widths, reduce fonts down to the specified minimums, increase heights, or allow an oversized matrix to continue with a repeated header. Very wide matrices may become multiple column panels. Formatting pressure should not stop conversion; any relaxed sizing is reported in the console.

## Generate a fresh default config

```bash
python qsf_to_paper.py --write-default-config config.json
```

Unknown keys, invalid values, incompatible minimum font sizes, and margins leaving insufficient space produce a clear error before conversion. Logic display is not part of this configuration version.
