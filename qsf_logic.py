"""Standalone, conservative QSF-to-paper logic analysis (no Word dependencies).

Run: python qsf_logic.py survey.qsf [-o instructions.json] [--numbering survey]
The existing qsf_to_paper.py is neither imported nor modified by this module.
"""
import argparse
import json
import re
import sys
from html import unescape
from pathlib import Path


class UnsupportedLogic(ValueError):
    pass


def text(value):
    """Flatten rich text for short respondent directions, never use raw HTML."""
    if isinstance(value, dict):
        value = value.get("Display", "")
    return " ".join(unescape(re.sub(r"<[^>]*>", " ", str(value or ""))).split())


def numbered_items(value):
    return [value[k] for k in sorted((k for k in value if str(k).isdigit()), key=lambda k: int(k))]


def parse_condition(value):
    """Parse a deliberately small subset; unfamiliar forms require review.

    A group child's If/And/Or Type or Conjunction joins it to its predecessor.
    Mixed operators are not guessed. Question predicates must identify an
    ordinary SelectableChoice, rather than a matrix cell or text comparison.
    """
    if not isinstance(value, dict):
        raise UnsupportedLogic("Unrecognized condition structure.")
    if value.get("Type") == "Expression":
        if value.get("LogicType") == "BooleanValue":
            truth = str(value.get("Value")).lower()
            if truth not in {"true", "false"}:
                raise UnsupportedLogic("Unrecognized Boolean value.")
            return ("constant", truth == "true")
        if value.get("LogicType") != "Question":
            raise UnsupportedLogic("Condition uses information other than a question response.")
        operator = value.get("Operator")
        locator = value.get("ChoiceLocator") or value.get("LeftOperand", "")
        count = re.fullmatch(r"q://([^/]+)/SelectedChoicesCount", str(locator))
        if count:
            if operator != "GreaterThanOrEqual" or not str(value.get("RightOperand", "")).isdigit():
                raise UnsupportedLogic("Unrecognized response-count comparison.")
            if value.get("QuestionIsInLoop", "no") not in {"no", False}:
                raise UnsupportedLogic("Response-count condition refers to a looped question.")
            minimum = int(value["RightOperand"])
            if minimum < 1 or value.get("QuestionID", count[1]) != count[1]:
                raise UnsupportedLogic("Invalid response-count comparison.")
            return ("count", count[1], minimum)
        if operator not in {"Selected", "NotSelected"}:
            raise UnsupportedLogic("Only Selected and NotSelected comparisons are supported initially.")
        match = re.fullmatch(r"q://([^/]+)/SelectableChoice/([^/]+)", str(locator))
        if not match or value.get("QuestionIsInLoop", "no") not in {"no", False}:
            raise UnsupportedLogic("Condition is not an ordinary choice selection.")
        qid = value.get("QuestionID") or match[1]
        if qid != match[1]:
            raise UnsupportedLogic("Condition question ID disagrees with its choice locator.")
        return ("choice", qid, match[2], operator == "Selected")
    if value.get("Type") not in {"BooleanExpression", "If", "And", "Or"}:
        raise UnsupportedLogic("Unrecognized condition group type.")
    children = numbered_items(value)
    if not children:
        raise UnsupportedLogic("Condition has no recognized expressions.")
    parsed = [parse_condition(child) for child in children]
    if len(parsed) == 1:
        return parsed[0]
    joins = []
    for child in children[1:]:
        join = child.get("Conjunction") or child.get("Conjuction") or child.get("Type")
        if join not in {"And", "Or"}:
            raise UnsupportedLogic("Condition conjunction is ambiguous.")
        joins.append(join.lower())
    if len(set(joins)) != 1:
        raise UnsupportedLogic("Mixed AND/OR groups require manual review.")
    return (joins[0], tuple(parsed))


def dependencies(condition):
    if condition[0] in {"choice", "count"}:
        return {condition[1]}
    if condition[0] == "constant":
        return set()
    return set().union(*(dependencies(child) for child in condition[1]))


def evaluate(condition, answers):
    """Return True, False or None (unknown); unknown answers never justify skips."""
    kind = condition[0]
    if kind == "constant":
        return condition[1]
    if kind == "count":
        return len(answers[condition[1]]) >= condition[2] if condition[1] in answers else None
    if kind == "choice":
        _, qid, choice, positive = condition
        if qid not in answers:
            return None
        selected = choice in answers[qid]
        return selected if positive else not selected
    results = [evaluate(child, answers) for child in condition[1]]
    if kind == "and":
        return False if False in results else (None if None in results else True)
    return True if True in results else (None if None in results else False)


def describe(condition, questions, references):
    kind = condition[0]
    if kind == "constant":
        raise UnsupportedLogic("Constant visibility is not a respondent instruction.")
    if kind == "count":
        _, qid, minimum = condition
        if qid not in references or questions[qid].get("QuestionType") != "MC":
            raise UnsupportedLogic("Response-count condition has no multiple-choice source on paper.")
        amount = "one response" if minimum == 1 else "{} responses".format(minimum)
        return "you selected at least {} in {}".format(amount, references[qid])
    if kind == "choice":
        _, qid, cid, positive = condition
        if qid not in references:
            raise UnsupportedLogic("Condition refers to a question absent from the paper survey.")
        q = questions[qid]
        if q.get("QuestionType") != "MC":
            raise UnsupportedLogic("Condition refers to a non-multiple-choice question.")
        choices = q.get("Choices", {})
        if cid not in choices or not text(choices[cid]):
            raise UnsupportedLogic("Condition refers to an absent or empty response label.")
        verb = "selected" if positive else "did not select"
        return 'you {} “{}” in {}'.format(verb, text(choices[cid]), references[qid])
    separator = " and " if kind == "and" else " or "
    phrases = [describe(child, questions, references) for child in condition[1]]
    return separator.join("({})".format(p) if child[0] in {"and", "or"} else p
                          for child, p in zip(condition[1], phrases))


def analyze_survey(data, numbering="sequential", question_numbering=None):
    """Return JSON-compatible instructions, source QIDs and review issues.

    Uses the converter's flow/block/orphan order, excluding trash and counting
    only response questions. Block branches become eligibility instructions;
    automatic skips never cross block boundaries. No config files are read.
    """
    if numbering not in {"sequential", "survey", "off"}:
        raise ValueError("numbering must be sequential, survey or off")
    elements = data.get("SurveyElements")
    if not isinstance(elements, list):
        raise ValueError("No SurveyElements array found; is this a QSF export?")
    questions, blocks, flows, skip_rules = {}, {}, [], {}
    for element in elements:
        payload = element.get("Payload", {})
        if element.get("Element") == "SQ" and isinstance(payload, dict):
            questions[payload.get("QuestionID")] = payload
            if payload.get("SkipLogic"):
                skip_rules.setdefault(payload.get("QuestionID"), []).append(payload["SkipLogic"])
        elif element.get("Element") == "BL":
            for block in payload if isinstance(payload, list) else [payload]:
                if isinstance(block, dict) and block.get("ID"):
                    blocks[block["ID"]] = block
                    for item in block.get("BlockElements", []):
                        if item.get("Type") == "Question" and item.get("SkipLogic"):
                            skip_rules.setdefault(item.get("QuestionID"), []).append(item["SkipLogic"])
        elif element.get("Element") == "FL":
            flows.append(payload)
    issues, sections, order = [], [], []
    visited, printed = set(), set()

    def review(scope, source_id, reason):
        issue = {"scope": scope, "source_id": source_id, "reason": reason}
        if issue not in issues:
            issues.append(issue)

    def trash(block):
        return str(block.get("Type", "")).lower() == "trash" or "trash" in text(block.get("Description")).lower()

    trash_qids = {e.get("QuestionID") for b in blocks.values() if trash(b)
                  for e in b.get("BlockElements", []) if e.get("Type") == "Question"}

    def emit(bid, conditions=(), uncertain=False):
        if bid in visited:
            review("block", bid, "Repeated block placement requires manual routing review.")
            # Instructions at an earlier occurrence cannot safely describe this reuse.
            for section in sections:
                if section["block_id"] == bid:
                    section["uncertain"] = True
            return
        visited.add(bid)
        if bid not in blocks:
            review("block", bid, "Flow references a missing block.")
            return
        block = blocks[bid]
        if trash(block):
            return
        ids = [e.get("QuestionID") for e in block.get("BlockElements", [])
               if e.get("Type") == "Question" and e.get("QuestionID") in questions
               and e.get("QuestionID") not in printed]
        if not ids:
            return
        options = block.get("Options") or {}
        random = options.get("RandomizeQuestions", False)
        if str(random).lower() not in {"false", "none", "0", ""} or options.get("Randomization"):
            uncertain = True
            review("block", bid, "Question randomization needs a chosen paper order before routing.")
        sections.append({"block_id": bid, "title": text(block.get("Description")),
                         "question_ids": ids, "conditions": conditions, "uncertain": uncertain})
        order.extend(ids)
        printed.update(ids)

    def walk(nodes, conditions=(), uncertain=False):
        if isinstance(nodes, dict):
            nodes = nodes.get("Flow", [])
        for node in nodes if isinstance(nodes, list) else []:
            kind = node.get("Type")
            if kind in {"Block", "Standard"}:
                emit(node.get("ID"), conditions, uncertain)
            elif kind == "Branch":
                try:
                    condition = parse_condition(node.get("BranchLogic") or node.get("Logic"))
                    inherited = conditions + (condition,)
                    bad = uncertain
                except UnsupportedLogic as exc:
                    review("flow", node.get("FlowID"), str(exc))
                    inherited, bad = conditions, True
                walk(node.get("Flow", []), inherited, bad)
            elif kind == "Group":
                walk(node.get("Flow", []), conditions, uncertain)
            elif kind in {"Randomizer", "LoopAndMerge"} or node.get("Flow"):
                review("flow", node.get("FlowID"), "Flow order or eligibility requires manual review: {}.".format(kind))
                walk(node.get("Flow", []), conditions, True)
            elif kind not in {"EmbeddedData"}:
                review("flow", node.get("FlowID"), "Untranslated survey flow action: {}.".format(kind))

    for flow in flows:
        walk(flow)
    for bid, block in blocks.items():
        if bid not in visited and not trash(block):
            if flows:
                review("block", bid, "Block is absent from survey flow; eligibility is unknown.")
            emit(bid, uncertain=bool(flows))
    orphaned = [qid for qid in questions if qid not in printed and qid not in trash_qids]
    if orphaned:
        sections.append({"block_id": None, "title": "", "question_ids": orphaned,
                         "conditions": (), "uncertain": True})
        order.extend(orphaned)
        review("survey", None, "Questions outside blocks require manual routing review.")

    references, number = {}, 0
    for qid in order:
        if questions[qid].get("QuestionType") != "DB":
            number += 1
            tag = text(questions[qid].get("DataExportTag"))
            mode = (question_numbering or {}).get(qid, numbering)
            if mode not in {"sequential", "survey", "off"}:
                raise ValueError("Invalid numbering mode for {}".format(qid))
            references[qid] = ("Q{}".format(number) if mode == "sequential" else
                               tag or ("Q{}".format(number) if mode == "survey" else qid))
    if len(set(references.values())) != len(references):
        # Distinguish ambiguous export tags rather than silently pointing at the wrong question.
        for qid, ref in list(references.items()):
            if list(references.values()).count(ref) > 1:
                review("question", qid, "Duplicate question tag; routing references require review.")
        duplicate_refs = {ref for ref in references.values() if list(references.values()).count(ref) > 1}
        references = {qid: ("{} [{}]".format(ref, qid) if ref in duplicate_refs else ref)
                      for qid, ref in references.items()}
    instructions, conditions, fallback = [], {}, {}

    for qid in order:
        q = questions[qid]
        if q.get("DisplayLogic"):
            try:
                condition = parse_condition(q["DisplayLogic"])
                phrase = describe(condition, questions, references)
                if any(order.index(dep) >= order.index(qid) for dep in dependencies(condition)):
                    raise UnsupportedLogic("Condition does not refer exclusively to earlier questions.")
                conditions[qid] = condition
                fallback[qid] = {"kind": "question_eligibility", "question_id": qid,
                                 "text": "Answer only if {}.".format(phrase)}
            except UnsupportedLogic as exc:
                conditions[qid] = None
                review("question", qid, str(exc))

    for section_index, section in enumerate(sections):
        bid, ids = section["block_id"], section["question_ids"]
        if section["conditions"] and not section["uncertain"]:
            try:
                joined = ("and", section["conditions"])
                first_index = order.index(ids[0])
                phrase = describe(joined, questions, references)
                if any(order.index(dep) >= first_index for dep in dependencies(joined)):
                    raise UnsupportedLogic("Block condition does not exclusively reference preceding questions.")
                instructions.append({"kind": "block_eligibility", "block_id": bid,
                                     "text": "Answer this section only if {}.".format(phrase)})
            except UnsupportedLogic as exc:
                review("block", bid, str(exc))
        if section["uncertain"]:
            review("block", bid, "Automatic routing disabled because block order or eligibility is unresolved.")
            continue
        # QSF exports can store skip rules on BlockElements rather than SQ payloads.
        # Initially translate only ordinary single-choice exits to this block's end.
        for index, qid in enumerate(ids):
            q = questions[qid]
            if qid not in skip_rules:
                continue
            proposals, destinations, seen, bad = [], {}, set(), False
            try:
                if q.get("QuestionType") != "MC" or q.get("Selector") not in {"SAVR", "SAHR", "SACOL", "DL", "SB"}:
                    raise UnsupportedLogic("Skip rule uses an unsupported response type.")
                for bundle in skip_rules[qid]:
                    if not isinstance(bundle, list):
                        raise UnsupportedLogic("Unrecognized skip-rule structure.")
                    for rule in bundle:
                        signature = json.dumps(rule, sort_keys=True)
                        if signature in seen:
                            continue
                        seen.add(signature)
                        match = re.fullmatch(r"q://([^/]+)/SelectableChoice/([^/]+)",
                                             str(rule.get("ChoiceLocator") or rule.get("Locator", "")))
                        if not match or match[1] != qid or rule.get("Condition") != "Selected":
                            raise UnsupportedLogic("Skip rule is not a simple selected-choice condition.")
                        cid = match[2]
                        if cid not in q.get("Choices", {}) or rule.get("SkipToDestination") != "ENDOFBLOCK":
                            raise UnsupportedLogic("Only selected-choice skips to end of block are supported initially.")
                        next_section = sections[section_index + 1] if section_index + 1 < len(sections) else None
                        direction = ("Skip to the {} section.".format(next_section["title"])
                                     if next_section and next_section["title"] not in {"", "Default Question Block"}
                                     else "Skip the remaining questions in this section.")
                        if cid in destinations:
                            continue
                        destinations[cid] = "ENDOFBLOCK"
                        proposals.append({"kind": "skip", "origin": "explicit_skip", "question_id": qid,
                                          "choice_id": cid, "choice_label": text(q["Choices"][cid]),
                                          "text": direction, "target_question_id": None,
                                          "target_block_id": next_section["block_id"] if next_section else None,
                                          "block_id": bid, "skipped_question_ids": ids[index + 1:]})
            except (UnsupportedLogic, AttributeError) as exc:
                bad = True
                review("question", qid, str(exc))
            if not bad:
                instructions.extend(proposals)
        for index, qid in enumerate(ids[:-1]):
            q = questions[qid]
            single = q.get("QuestionType") == "MC" and q.get("Selector") in {"SAVR", "SAHR", "SACOL", "DL", "SB"}
            if not single or qid in skip_rules or q.get("ChoiceDisplayLogic"):
                continue
            next_id = ids[index + 1]
            if next_id not in conditions or conditions[next_id] is None or dependencies(conditions[next_id]) != {qid}:
                continue
            choices = q.get("Choices", {})
            if not isinstance(choices, dict) or not choices:
                continue
            if any(isinstance(c, dict) and (c.get("DisplayLogic") or c.get("ExclusiveAnswer")) for c in choices.values()):
                continue
            proposals, covered, reliable = [], set(), True
            for choice_id, choice in choices.items():
                answers = {qid: {str(choice_id)}}
                skipped, destination = [], index + 1
                while destination < len(ids):
                    target = ids[destination]
                    if target not in conditions:
                        break
                    condition = conditions[target]
                    state = evaluate(condition, answers) if condition is not None else None
                    if state is None:
                        reliable = False
                        break
                    if state:
                        break
                    skipped.append(target)
                    destination += 1
                if not reliable:
                    break
                if not skipped:
                    continue
                target_id = ids[destination] if destination < len(ids) else None
                if target_id is not None and target_id not in references:
                    reliable = False  # Do not jump over a source instruction (DB) to find a number.
                    break
                direction = ("Skip to {}.".format(references[target_id]) if target_id else
                             "Skip the remaining questions in this section.")
                proposals.append({"kind": "skip", "question_id": qid, "choice_id": str(choice_id),
                                  "choice_label": text(choice), "text": direction,
                                  "target_question_id": target_id, "block_id": bid,
                                  "skipped_question_ids": skipped})
                covered.update(skipped)
            if reliable and proposals:
                instructions.extend(proposals)
                # Remove eligibility only when every false case is covered here.
                # Other controllers must keep their own conditions visible.
                skipped_by_choice = {p["choice_id"]: set(p["skipped_question_ids"]) for p in proposals}
                for target in covered:
                    all_false_routes_covered = all(
                        evaluate(conditions[target], {qid: {str(cid)}}) is not False
                        or target in skipped_by_choice.get(str(cid), set()) for cid in choices)
                    if dependencies(conditions[target]) == {qid} and all_false_routes_covered:
                        fallback.pop(target, None)
    # Long dropdowns may become write-in fields on paper. A single question-level
    # direction is more useful than dozens of identical choice-level directions.
    grouped, compacted = {}, []
    for instruction in instructions:
        if instruction["kind"] != "skip":
            compacted.append(instruction)
            continue
        key = (instruction["question_id"], instruction["target_question_id"],
               tuple(instruction["skipped_question_ids"]))
        grouped.setdefault(key, []).append(instruction)
    for group in grouped.values():
        first = group[0]
        choices = questions[first["question_id"]].get("Choices", {})
        affected = {item["choice_id"] for item in group}
        excluded = [cid for cid in choices if str(cid) not in affected]
        if len(group) >= 4 and len(excluded) == 1:
            compacted.append({"kind": "conditional_skip", "question_id": first["question_id"],
                              "choice_ids": [item["choice_id"] for item in group],
                              "excluded_choice_ids": [str(cid) for cid in excluded],
                              "target_question_id": first["target_question_id"],
                              "block_id": first["block_id"],
                              "skipped_question_ids": first["skipped_question_ids"],
                              "text": 'If you selected anything other than “{}”, {}{}'.format(
                                  text(choices[excluded[0]]), first["text"][0].lower(), first["text"][1:])})
        else:
            compacted.extend(group)
    instructions = compacted + list(fallback.values())
    return {"schema_version": 1, "survey_name": text(data.get("SurveyEntry", {}).get("SurveyName")),
            "numbering": numbering, "question_references": references,
            "blocks": [{k: s[k] for k in ("block_id", "title", "question_ids")} for s in sections],
            "instructions": instructions, "review_required": issues}


def analyze_file(path, numbering="sequential"):
    with open(path, encoding="utf-8-sig") as source:
        return analyze_survey(json.load(source), numbering)


def format_report(report):
    lines = [report["survey_name"] or "Survey logic instructions", ""]
    for instruction in report["instructions"]:
        kind = instruction["kind"]
        if kind == "block_eligibility":
            location = "Section {}".format(instruction["block_id"])
        else:
            qid = instruction["question_id"]
            location = report["question_references"].get(qid, qid)
            if kind == "skip":
                location += ' — “{}”'.format(instruction["choice_label"])
        lines.append("{}: {}".format(location, instruction["text"]))
    if not report["instructions"]:
        lines.append("No respondent instructions generated.")
    if report["review_required"]:
        lines.extend(["", "Manual review (these notes are not respondent instructions):"])
        lines.extend("{}: {}".format(item["source_id"] or "Survey", item["reason"])
                     for item in report["review_required"])
    return "\n".join(lines) + "\n"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("qsf", type=Path)
    parser.add_argument("-o", "--output", type=Path, help="Write a JSON instruction report")
    parser.add_argument("--numbering", choices=["sequential", "survey", "off"], default="sequential")
    args = parser.parse_args()
    try:
        report = analyze_file(args.qsf, args.numbering)
        if args.output:
            args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(format_report(report), end="")
        return 0
    except (OSError, ValueError, TypeError) as exc:
        print("Error: {}".format(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
