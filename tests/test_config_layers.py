"""Check config layering and fallback from a standalone checkout."""
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import qsf_to_paper as m


def test_layers():
    with tempfile.TemporaryDirectory() as temp:
        temp = Path(temp)
        main = temp / 'config.json'
        survey = temp / 'survey_config.json'
        main.write_text(json.dumps({'questions': {'font_size': 13, 'bold': True},
                                   'question_overrides': {'QID2': {'questions': {'font_size': 15, 'blank_lines_before': 1}}}}))
        survey.write_text(json.dumps({'questions': {'font_size': 14},
                                     'question_overrides': {'QID2': {'questions': {'font_size': 16}}}}))
        merged, warnings = m.load_config_layers(main, survey)
        assert not warnings
        resolved = m.resolve_config(merged, {'questions': {'font_size': 12}})
        assert resolved['questions']['font_size'] == 12 and resolved['questions']['bold']
        assert resolved['question_overrides']['QID2']['questions'] == {'font_size': 16, 'blank_lines_before': 1}
        merged, warnings = m.load_config_layers(temp / 'absent.json', survey)
        assert len(warnings) == 1 and merged['questions']['font_size'] == 14

        script = temp / 'qsf_to_paper.py'
        shutil.copyfile(ROOT / 'qsf_to_paper.py', script)
        elsewhere = temp / 'other-directory'
        elsewhere.mkdir()
        out = temp / 'output.docx'

        def run(*args):
            result = subprocess.run([sys.executable, str(script), str(ROOT / 'examples/Example_survey.qsf'),
                                     '-o', str(out), *args], cwd=elsewhere, capture_output=True, text=True)
            assert result.returncode == 0, result.stderr
            doc = Document(out)
            paragraph = next(p for p in doc.paragraphs if 'Please rate your satisfaction' in p.text)
            return paragraph, result

        p, result = run('--survey-config', str(survey), '--question-font-size', '12')
        assert p.runs[1].font.size.pt == 16 and p.runs[1].bold
        main.unlink()
        survey.unlink()
        p, result = run()
        assert p.runs[1].font.size.pt == 12 and not p.runs[1].bold and not result.stderr
        p, result = run('--config', str(main), '--survey-config', str(survey))
        assert p.runs[1].font.size.pt == 12 and 'not found' in result.stderr
        main.write_text('{broken')
        result = subprocess.run([sys.executable, str(script), str(ROOT / 'examples/Example_survey.qsf'),
                                 '-o', str(temp / 'invalid.docx')], capture_output=True, text=True)
        assert result.returncode != 0 and 'Config file' in result.stderr and 'Traceback' not in result.stderr
    print('Layered configuration checks passed.')


if __name__ == '__main__':
    test_layers()
