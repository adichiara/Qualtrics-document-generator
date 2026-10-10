# Qualtrics document generator

Convert a Qualtrics `.qsf` survey export into an editable Word document for collecting responses on paper.

The converter preserves survey wording and supported rich text formatting, creates editable response tables, and adapts pagination for print. Matrix questions use statement rows, equal-width scale columns, and optional rotated headers. Online page breaks are ignored; later survey blocks start on new pages by default.

## Run locally

Requires Python 3.8 or newer.

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

```bash
python qsf_to_paper.py your-survey.qsf --survey-config your-survey_config.json
python qsf_to_paper.py your-survey.qsf --response-font-size 10 --bold-questions
python qsf_to_paper.py your-survey.qsf --set matrix.row_bottom_border.enabled=true
```

Settings apply in this order: built-in defaults, main config, survey config, general CLI options, then per-question overrides. Missing config files do not prevent conversion. See [the configuration guide](docs/configuration.md) for all settings and examples.

```bash
python qsf_to_paper.py --help
python qsf_to_paper.py your-survey.qsf --list-questions
```

## Included examples

- [Example survey QSF](examples/Example_survey.qsf)
- [Example Word document](examples/Example_survey-paper.docx) and [PDF](examples/Example_survey-paper.pdf), using the default print layout
- [Example with respondent logic in Word](examples/Example_survey-paper-logic.docx) and [PDF](examples/Example_survey-paper-logic.pdf), with the analyzer's years-of-service skip direction added to a separate generated copy; [analysis report](examples/Example_survey-logic.json)
- [Survey-specific configuration](examples/Example_survey_config.json), demonstrating survey numbering, bold prompts, and a QID-specific matrix border style
- [Qualtrics product survey demo](examples/Demo_product_survey.qsf)
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
```

The formatter preserves all non-trash questions it finds, including those absent from survey flow. Question types without a full paper equivalent produce a console warning for manual review. Very large matrices may span pages or become multiple column panels to retain their content.

The Word converter currently omits logic. A separate [respondent logic prototype](docs/logic.md) proposes skip directions for questions and “Answer this section only if…” directions for conditional blocks, without changing the converter:

```bash
python qsf_logic.py your-survey.qsf -o instructions.json
python tests/test_logic.py
```

The script currently generates Word documents; the included PDF is an exported example.
