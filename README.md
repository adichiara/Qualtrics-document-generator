# Qualtrics document generator

Convert a Qualtrics `.qsf` survey export into an editable Word document for collecting responses on paper.

The converter preserves survey wording and supported rich text formatting, creates editable response tables, and adapts pagination for print. Matrix questions use statement rows, equal-width scale columns, and optional rotated headers. Online page breaks are ignored. Block headings and fresh pages for later blocks are independent options, both off by default.

Dropdowns retain categorical choices such as military rank grades as multiple-choice response tables. Long lists of short labels can use up to three columns, read down each column in survey order. Numeric ranges (such as age, height, and years of service) become write-ins at 12 choices by default; unordered lists such as names become write-ins at 50 choices. Configure either threshold, the column limit, or `responses.dropdown_mode` (`auto`, `choices`, or `write_in`), including per-question overrides.

## Run locally

Requires Python 3.8 or newer.

[Download the current code and configuration](downloads/Qualtrics-document-generator-survey-config-export-20261028.zip).

```bash
pip install -r requirements.txt
python qsf_to_paper.py your-survey.qsf
```

This creates `your-survey-paper.docx` beside the input. Open it in Word to review, adjust, print, or export to PDF.

To choose an output location:

```bash
python qsf_to_paper.py your-survey.qsf -o paper-survey.docx
```

## Configure formatting

All defaults are built into the Python code. The optional `config.json` beside the script supplies your general preferences. A separate survey config can override those preferences and individual questions by their stable QID.

For a complete copy-and-paste reference, see [all options with comments](examples/config_reference.jsonc) and the matching [runnable JSON config](examples/config_reference.json). The commented reference includes allowed values, units, defaults, and a complete per-question override template. The converters require standard JSON, so use the `.json` file directly; comments must be removed before using content copied from `.jsonc`.

Export a ready-to-edit settings file for every question in your source survey, using the current defaults, config files, and CLI settings:

```bash
python qsf_logic_to_paper.py your-survey.qsf --export-survey-config
# Edit your-survey_config.json, then regenerate:
python qsf_logic_to_paper.py your-survey.qsf --survey-config your-survey_config.json
```

Both converters support `--export-survey-config [PATH]`. The JSON file includes every supported per-question setting, with `_comment` labels identifying the source questions; document-wide settings appear once at the top. Exporting does not generate Word/PDF or run the logic analyzer. Explicit question values override global values, so edit a question's value (or remove it to inherit) when experimenting. See the [exported example](examples/Example_survey_explicit_config.json).

```bash
python qsf_to_paper.py your-survey.qsf --survey-config your-survey_config.json
python qsf_to_paper.py your-survey.qsf --response-font-size 10 --bold-questions
python qsf_to_paper.py your-survey.qsf --set matrix.row_bottom_border.enabled=true
python qsf_to_paper.py your-survey.qsf --show-block-names --new-page-per-block
```

Settings apply in this order: built-in defaults, main config, survey config, general CLI options, then per-question overrides. Missing config files do not prevent conversion. See [the configuration guide](docs/configuration.md) for all settings and examples.

Rank-grade lists automatically use separate columns for enlisted, warrant, and officer grades, regardless of QID (`responses.group_rank_grades`, default `true`). Use `responses.max_rows` to cap the number of options in each vertical column and `responses.column_breaks_before` to choose exactly where a new column begins. For example, `[10, 15]` starts new columns before options 10 and 15. Manual breaks override automatic grouping; see the [MRE rank-grade config](examples/FY27_MRE_IOTE_config.json). Row caps and manual breaks default to automatic (`null` and `[]`).

Piped-text codes are shown as readable placeholders by default: `${e://Field/Issued_Trouser_Size}` becomes `{Issued_Trouser_Size}`. Use `piped_text.show: false` in JSON, `--no-show-piped-text`, or `--set piped_text.show=false` to remove them from question/response text and generated respondent instructions. Both converter CLIs support the switch; per-question overrides are also supported. This identifies piped text without attempting to resolve online values.

```bash
python qsf_to_paper.py --help
python qsf_to_paper.py your-survey.qsf --list-questions
python qsf_logic_to_paper.py your-survey.qsf --list-questions
```

## Included examples

- Latest response-layout update: [source ZIP](downloads/Qualtrics-document-generator-4b6916e.zip), [MRE Word](examples/FY27_MRE_IOTE-paper-logic-4b6916e.docx), and [MRE PDF](examples/FY27_MRE_IOTE-paper-logic-4b6916e.pdf). These versioned downloads include vertical Yes/No choices, more form writing space where the page permits, banded lists above 10 choices, and 0.5-inch multiple-choice indentation.
- [Example survey QSF](examples/Example_survey.qsf)
- [Example Word document](examples/Example_survey-paper.docx) and [PDF](examples/Example_survey-paper.pdf), using the default print layout
- [Example with respondent logic in Word](examples/Example_survey-paper-logic.docx) and [PDF](examples/Example_survey-paper-logic.pdf), with the analyzer's years-of-service skip direction added to a separate generated copy; [analysis report](examples/Example_survey-logic.json)
- [Survey-specific configuration](examples/Example_survey_config.json), demonstrating survey numbering, bold prompts, and a QID-specific matrix border style
- [Qualtrics product survey demo](examples/Demo_product_survey.qsf)
- [Real fuel handler survey](examples/Extreme_Cold_Weather_Fuel_Handler_Handwear_Survey.qsf), with [Word](examples/Extreme_Cold_Weather_Fuel_Handler_Handwear_Survey-paper-logic.docx), [PDF](examples/Extreme_Cold_Weather_Fuel_Handler_Handwear_Survey-paper-logic.pdf), [survey config](examples/Extreme_Cold_Weather_Fuel_Handler_Handwear_Survey_config.json), and [routing/layout evaluation](docs/handwear-evaluation.md)
- [FY27 MRE survey](examples/FY27_MRE_IOTE.qsf), with default-settings [Word](examples/FY27_MRE_IOTE-paper-logic.docx), [PDF](examples/FY27_MRE_IOTE-paper-logic.pdf), [logic report](examples/FY27_MRE_IOTE-logic.json), and [evaluation of layout and unsupported logic](docs/mre-evaluation.md)
- [Manually revised layout reference](examples/references/Example_survey-paper_revised.docx), retained as a formatting reference

Generate examples into an ignored local output directory:

```bash
mkdir output
python qsf_to_paper.py examples/Example_survey.qsf -o output/example.docx
python qsf_to_paper.py examples/Example_survey.qsf --survey-config examples/Example_survey_config.json -o output/example-custom.docx
python qsf_to_paper.py examples/Demo_product_survey.qsf -o output/demo.docx
```

The supplied Word/PDF example shows the reviewed default layout. The survey-specific configuration changes its appearance when applied.

## Validation and scope

Run the configuration regression checks from the repository root:

```bash
python tests/test_configuration.py
python tests/test_config_layers.py
python tests/test_logic.py
python tests/test_logic_document.py
python tests/test_response_layout.py
python tests/test_block_sections.py
python tests/test_dropdown_layout.py
python tests/test_choice_columns.py
python tests/test_piped_text.py
python tests/test_survey_config_export.py
```

The formatter preserves all non-trash questions it finds, including those absent from survey flow. Question types without a full paper equivalent produce a console warning for manual review. Very large matrices may span pages or become multiple column panels to retain their content.

The Word converter currently omits logic. A separate [respondent logic prototype](docs/logic.md) proposes skip directions for questions and “Answer this section only if…” directions for conditional blocks, without changing the converter:

```bash
python qsf_logic.py your-survey.qsf -o instructions.json
python qsf_logic_to_paper.py your-survey.qsf --report instructions.json
```

The optional adapter inserts respondent instructions into a separate generated Word copy; the original converter is unchanged. These scripts generate Word documents; the included PDFs are exported examples.
