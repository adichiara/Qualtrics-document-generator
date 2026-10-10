"""Routing checks; synthetic conditions use the exported QSF expression shape."""
import copy
import itertools
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import qsf_logic as logic


def choice(qid="QID1", cid="1", operator="Selected"):
    return {"Type": "Expression", "LogicType": "Question", "QuestionID": qid,
            "ChoiceLocator": "q://{}/SelectableChoice/{}".format(qid, cid), "Operator": operator}


def condition(expr):
    return {"Type": "BooleanExpression", "0": {"Type": "If", "0": expr}}


def fixture(count=4):
    questions = []
    for number in range(1, count + 1):
        questions.append({"Element": "SQ", "Payload": {
            "QuestionID": "QID{}".format(number), "DataExportTag": "Q{}".format(number * 10),
            "QuestionType": "MC", "Selector": "SAVR", "QuestionText": "Question {}".format(number),
            "Choices": {"1": {"Display": "<b>Yes</b>"}, "2": {"Display": "No"}}}})
    questions[1]["Payload"]["DisplayLogic"] = condition(choice())
    return {"SurveyEntry": {"SurveyName": "Logic test"}, "SurveyElements": questions + [
        {"Element": "BL", "Payload": [{"ID": "BL1", "Description": "First section", "Type": "Default",
         "BlockElements": [{"Type": "Question", "QuestionID": "QID{}".format(i)} for i in range(1, count + 1)]}]},
        {"Element": "FL", "Payload": {"Flow": [{"Type": "Block", "ID": "BL1"}]}}]}


def question(data, index):
    return data["SurveyElements"][index - 1]["Payload"]


class LogicTests(unittest.TestCase):
    def test_simple_skip_and_pure_input(self):
        data = fixture()
        original = copy.deepcopy(data)
        report = logic.analyze_survey(data)
        self.assertEqual(data, original)
        self.assertFalse(report["review_required"])
        self.assertEqual(len(report["instructions"]), 1)
        skip = report["instructions"][0]
        self.assertEqual((skip["question_id"], skip["choice_id"], skip["target_question_id"]), ("QID1", "2", "QID3"))
        self.assertEqual(skip["text"], "Skip to Q3.")

    def test_skip_group_and_check_destination(self):
        data = fixture(5)
        question(data, 3)["DisplayLogic"] = condition(choice())
        question(data, 4)["DisplayLogic"] = condition(choice(cid="2"))
        report = logic.analyze_survey(data)
        skip = next(i for i in report["instructions"] if i["kind"] == "skip")
        self.assertEqual(skip["target_question_id"], "QID4")
        self.assertEqual(skip["skipped_question_ids"], ["QID2", "QID3"])
        # Q4 still needs eligibility: the Yes route reaches intervening questions first.
        self.assertTrue(any(i.get("question_id") == "QID4" and i["kind"] == "question_eligibility"
                            for i in report["instructions"]))

    def test_unknown_destination_does_not_guess(self):
        data = fixture()
        question(data, 3)["DisplayLogic"] = condition(choice("QID2"))
        report = logic.analyze_survey(data)
        self.assertFalse(any(i["kind"] == "skip" and i["question_id"] == "QID1" for i in report["instructions"]))
        self.assertTrue(any(i["kind"] == "question_eligibility" and i["question_id"] == "QID2" for i in report["instructions"]))

    def test_multiple_choice_no_inversion(self):
        data = fixture()
        question(data, 1)["Selector"] = "MAVR"
        report = logic.analyze_survey(data)
        self.assertEqual(report["instructions"][0]["kind"], "question_eligibility")

    def test_negative_predicate(self):
        data = fixture()
        question(data, 2)["DisplayLogic"] = condition(choice(operator="NotSelected"))
        skip = logic.analyze_survey(data)["instructions"][0]
        self.assertEqual(skip["choice_id"], "1")

    def test_compound_one_controller(self):
        data = fixture()
        question(data, 1)["Choices"]["3"] = {"Display": "Unsure"}
        expr = condition(choice())
        expr["0"]["1"] = dict(choice(cid="2"), Conjunction="Or")
        question(data, 2)["DisplayLogic"] = expr
        report = logic.analyze_survey(data)
        self.assertEqual(report["instructions"][0]["choice_id"], "3")
        self.assertIsNone(logic.evaluate(logic.parse_condition(expr), {}))

    def test_all_false_routes_need_coverage_before_removing_eligibility(self):
        data = fixture(5)
        question(data, 1)["Choices"]["3"] = {"Display": "Unsure"}
        question(data, 3)["DisplayLogic"] = condition(choice(cid="2"))
        question(data, 4)["DisplayLogic"] = condition(choice())
        report = logic.analyze_survey(data)
        self.assertTrue(any(i["kind"] == "question_eligibility" and i["question_id"] == "QID4"
                            for i in report["instructions"]))

    def test_routes_and_fallbacks_for_all_short_single_controller_patterns(self):
        # Exhaust all eligibility sets for three consecutive questions and three answers.
        # Check paper routes against the intended eligibility sets, independently of the parser.
        sets = [set(c) for n in range(1, 4) for c in itertools.combinations(["1", "2", "3"], n)]
        for pattern in itertools.product(sets, repeat=3):
            data = fixture(5)
            question(data, 1)["Choices"]["3"] = {"Display": "Unsure"}
            for index, allowed in enumerate(pattern, 2):
                expr = condition(choice(cid=sorted(allowed)[0]))
                for n, cid in enumerate(sorted(allowed)[1:], 1):
                    expr["0"][str(n)] = dict(choice(cid=cid), Conjunction="Or")
                question(data, index)["DisplayLogic"] = expr
            report = logic.analyze_survey(data)
            routes = {i["choice_id"]: i for i in report["instructions"] if i["kind"] == "skip"}
            eligibility = {i["question_id"] for i in report["instructions"] if i["kind"] == "question_eligibility"}
            for cid in ["1", "2", "3"]:
                skipped = routes.get(cid, {}).get("skipped_question_ids", [])
                for index, allowed in enumerate(pattern, 2):
                    qid = "QID{}".format(index)
                    if qid in skipped:
                        self.assertNotIn(cid, allowed)
                    elif cid not in allowed:
                        self.assertIn(qid, eligibility)
                if cid in routes:
                    target_index = int(routes[cid]["target_question_id"][3:])
                    if target_index < 5:
                        self.assertIn(cid, pattern[target_index - 2])

    def test_unsupported_and_missing_sources(self):
        for expr in [dict(choice(), Operator="GreaterThan"), choice("MISSING"),
                     {"Type": "Expression", "LogicType": "EmbeddedField", "Operator": "EqualTo"},
                     choice(cid="missing")]:
            data = fixture()
            question(data, 2)["DisplayLogic"] = condition(expr)
            report = logic.analyze_survey(data)
            self.assertFalse(report["instructions"])
            self.assertTrue(report["review_required"])

    def test_block_branch_uses_answer_only_if(self):
        data = fixture()
        blocks = data["SurveyElements"][-2]["Payload"]
        blocks[0]["BlockElements"] = [{"Type": "Question", "QuestionID": "QID1"}]
        blocks.append({"ID": "BL2", "Description": "Conditional section", "Type": "Standard",
                       "BlockElements": [{"Type": "Question", "QuestionID": "QID{}".format(i)} for i in range(2, 5)]})
        question(data, 2).pop("DisplayLogic")
        data["SurveyElements"][-1]["Payload"]["Flow"].append({
            "Type": "Branch", "FlowID": "FL3", "BranchLogic": condition(choice()),
            "Flow": [{"Type": "Standard", "ID": "BL2"}]})
        report = logic.analyze_survey(data)
        self.assertEqual(report["instructions"], [{"kind": "block_eligibility", "block_id": "BL2",
                                                  "text": "Answer this section only if you selected “Yes” in Q1."}])

    def test_end_of_section_does_not_skip_next_block(self):
        data = fixture(2)
        report = logic.analyze_survey(data)
        self.assertIsNone(report["instructions"][0]["target_question_id"])
        self.assertEqual(report["instructions"][0]["text"], "Skip the remaining questions in this section.")

    def test_randomized_order_disables_skips(self):
        data = fixture()
        data["SurveyElements"][-2]["Payload"][0]["Options"] = {"RandomizeQuestions": True}
        report = logic.analyze_survey(data)
        self.assertFalse(any(i["kind"] == "skip" for i in report["instructions"]))
        self.assertTrue(report["review_required"])

    def test_unsupported_branch_never_becomes_partial_block_instruction(self):
        data = fixture()
        data["SurveyElements"][-1]["Payload"]["Flow"] = [{"Type": "Branch", "FlowID": "FL2",
            "BranchLogic": condition({"Type": "Expression", "LogicType": "EmbeddedField"}),
            "Flow": [{"Type": "Block", "ID": "BL1"}]}]
        report = logic.analyze_survey(data)
        self.assertFalse(any(i["kind"] in {"skip", "block_eligibility"} for i in report["instructions"]))
        self.assertTrue(report["review_required"])

    def test_future_question_and_ambiguous_groups_require_review(self):
        data = fixture()
        question(data, 2)["DisplayLogic"] = condition(choice("QID3"))
        report = logic.analyze_survey(data)
        self.assertFalse(report["instructions"])
        self.assertTrue(report["review_required"])
        expr = condition(choice())
        expr["0"]["1"] = choice(cid="2")  # No connector: do not assume AND or OR.
        with self.assertRaises(logic.UnsupportedLogic):
            logic.parse_condition(expr)

    def test_repeated_blocks_and_duplicate_tags_are_flagged(self):
        data = fixture()
        data["SurveyElements"][-1]["Payload"]["Flow"].append({"Type": "Block", "ID": "BL1"})
        report = logic.analyze_survey(data)
        self.assertFalse(any(i["kind"] == "skip" for i in report["instructions"]))
        self.assertTrue(report["review_required"])
        data = fixture()
        question(data, 2)["DataExportTag"] = "Q30"
        report = logic.analyze_survey(data, "survey")
        self.assertEqual(report["question_references"]["QID3"], "Q30 [QID3]")
        self.assertTrue(report["review_required"])

    def test_source_numbering_and_db(self):
        data = fixture()
        question(data, 4)["QuestionType"] = "DB"
        self.assertEqual(logic.analyze_survey(data, "survey")["instructions"][0]["text"], "Skip to Q30.")
        self.assertNotIn("QID4", logic.analyze_survey(data)["question_references"])
        self.assertEqual(logic.analyze_survey(data, "off")["instructions"][0]["text"], "Skip to Q30.")
        question(data, 1)["DataExportTag"] = ""
        self.assertEqual(logic.analyze_survey(data, "off")["question_references"]["QID1"], "QID1")

    def test_long_choice_lists_share_one_direction(self):
        data = fixture()
        for cid in range(3, 8):
            question(data, 1)["Choices"][str(cid)] = {"Display": "Option {}".format(cid)}
        report = logic.analyze_survey(data)
        self.assertEqual(len(report["instructions"]), 1)
        item = report["instructions"][0]
        self.assertEqual(item["kind"], "conditional_skip")
        self.assertEqual(set(item["choice_ids"]), {str(cid) for cid in range(2, 8)})
        self.assertEqual(item["text"], "If you selected anything other than “Yes”, skip to Q3.")

    def test_db_destination_and_existing_skip_use_fallback(self):
        for change in ["DB", "skip"]:
            data = fixture()
            if change == "DB":
                question(data, 3)["QuestionType"] = "DB"
            else:
                question(data, 1)["SkipLogic"] = {"destination": "QID4"}
            report = logic.analyze_survey(data)
            self.assertFalse(any(i["kind"] == "skip" for i in report["instructions"]))
            self.assertTrue(any(i["kind"] == "question_eligibility" for i in report["instructions"]))

    def test_supplied_surveys_and_cli(self):
        report = logic.analyze_file(ROOT / "examples/Example_survey.qsf")
        self.assertEqual(len(report["question_references"]), 13)
        self.assertNotIn("QID19795634", report["question_references"])
        self.assertFalse(report["review_required"])
        self.assertEqual(report["instructions"][0]["question_id"], "QID19795631")
        self.assertEqual(len(logic.analyze_file(ROOT / "examples/Demo_product_survey.qsf")["question_references"]), 11)
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "instructions.json"
            result = subprocess.run([sys.executable, str(ROOT / "qsf_logic.py"), str(ROOT / "examples/Example_survey.qsf"),
                                     "-o", str(output)], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(output.read_text(encoding="utf-8")), report)


if __name__ == "__main__":
    unittest.main()
