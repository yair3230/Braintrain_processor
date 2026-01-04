import os.path

# import docx.text.font
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
import re
import json
import hashlib

HEBREW_LETTERS = r'\u0590-\u05FF'
DOCS_FOLDER = '.\\doc_files'
FILE_NAME = 'nueroanatomy_tashpag.docx'
DOCX_FILE = os.path.join(DOCS_FOLDER, FILE_NAME)
EXPECTED_RESULTS_FILE = 'expected_results.json'
EXPECTED_RESULTS_FODLER = 'expected_results'
STRATEGY_JSON = os.path.join(DOCS_FOLDER, 'strategy.json')


def is_numbered(para):
    line = para.text.strip()
    if not line:
        return False

    p_xml = para._element

    # First get <w:pPr>
    pPr = p_xml.find(qn('w:pPr'))
    if pPr is not None:
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            numId = numPr.find(qn('w:numId')).get(qn('w:val'))
            ilvl = numPr.find(qn('w:ilvl')).get(qn('w:val'))
            # print(numId, ilvl)
            return True

    if line[0].isdigit():
        return True
    return False


def process_answer(para, line, questions, a_strat, correct_strat, blanks):
    # Remove the letters and the dot at the beginning of the question
    # Ex. "a. Why are u gay?" turns into "Why are u gay?"
    if a_strat == 'numbered':
        answer = line[2:]
    else:
        answer = line
    if blanks:
        answer = answer.strip()
        answer_set = set(answer)
        if len(answer_set) == 1 and answer_set.pop() == blanks:
            return

    questions[-1]['answers']['text'].append(answer)

    if correct_strat:
        is_correct = False
        if correct_strat == 'bold':
            is_correct = any(run.bold for run in para.runs if run.bold)
        elif correct_strat == 'highlight':
            is_correct = any(
                run.font.highlight_color for run in para.runs if run.font.highlight_color != WD_COLOR_INDEX.WHITE)

        # I cant believe this shit exists fr fr
        elif correct_strat == 'heart':
            is_correct = '❤' in line
            questions[-1]['answers']['text'][-1] = answer.replace('❤', '')
        if is_correct:
            questions[-1]['answers']['checkbox'].append('1')
        else:
            questions[-1]['answers']['checkbox'].append('0')
    else:
        questions[-1]['answers']['checkbox'].append('0')


def process_file(file_name, save_file=False):
    file_path = os.path.join(DOCS_FOLDER, file_name)
    doc = Document(file_path)
    correct_strat = None
    blanks = None
    with open(STRATEGY_JSON, 'r') as openfile:
        settings = json.load(openfile)
        if file_name not in settings:
            print(f'No strategy for file {file_name}')
            exit(-1)
        file_settings = settings[file_name]
        q_strat = file_settings['q']
        a_strat = file_settings['a']
        if 'correct' in file_settings:
            correct_strat = file_settings['correct']
        if 'blanks' in file_settings:
            blanks = file_settings['blanks']

    # each item in this list dict
    # {'question': 'the sky is', 'answers': ['green', 'blue', ...], 'checkbox': ['0', '1', ...]
    questions = []
    found_first_line = False
    next_line_is_question = False
    counter = 0

    for para in doc.paragraphs:
        if 'header_lines' in file_settings and counter < file_settings['header_lines']:
            counter += 1
            continue
        line = para.text.strip()
        line = line.replace('"', '\\"')
        line = line.replace('”', '\\"')
        # print(line)

        # Skip header
        if 'header_lines' not in file_settings and not found_first_line:
            if is_numbered(para):
                found_first_line = True
            else:
                continue

        # Skip empty lines
        if not line:
            if q_strat in ['after_whitespace']:
                next_line_is_question = True
            continue

        if q_strat in ["numbered", "native_numbering"] and is_numbered(para):
            q = line

            # Remove the digits and the dot at the beginning of the question
            # Ex. "1. Why are u gay?" turns into "Why are u gay?"
            if q_strat == 'numbered':
                regex = r'(?:\d+)\. (.*)'
                q = re.match(regex, q).group(1)
            elif q_strat == 'native_numbering':
                q = line
            else:
                print('ERROR: unknown question strategy')
                exit(-1)
            q_dict = {'question': q, 'answers': {'text': [], 'checkbox': []}}
            questions.append(q_dict)
        elif q_strat in ['header']:
            if next_line_is_question:
                next_line_is_question = False
                q_dict = {'question': line, 'answers': {'text': [], 'checkbox': []}}
                questions.append(q_dict)
                continue
            else:
                regex = r'.*(\d:)'
                if re.match(regex, line):
                    next_line_is_question = True
                    continue
                else:
                    process_answer(para, line, questions, a_strat, correct_strat, blanks)
        elif q_strat in ['single_line_header']:
            regex = 'שאלה' + ' \d+(.*)'
            if re.match(regex, line):
                q = re.match(regex, line).group(1)

                # Remove first char
                if q[0] in [':', '-']:
                    q = q[1:]
                q = q.strip()
                q_dict = {'question': q, 'answers': {'text': [], 'checkbox': []}}
                questions.append(q_dict)
            else:
                process_answer(para, line, questions, a_strat, correct_strat, blanks)
        elif q_strat in ['after_whitespace'] and next_line_is_question:
            next_line_is_question = False
            q_dict = {'question': line, 'answers': {'text': [], 'checkbox': []}}
            questions.append(q_dict)
        else:
            process_answer(para, line, questions, a_strat, correct_strat, blanks)

    # Turn single answers to correct answers
    for item in questions:
        if len(item['answers']['text']) == 1:
            item['answers']['checkbox'] = ['1']
        # print(item)

    json_obj = json.dumps(questions)
    json_hash = hashlib.md5(json_obj.encode('utf-8')).hexdigest()
    mismatch = False
    with open(EXPECTED_RESULTS_FILE, 'r+') as openfile:
        expected_json = json.load(openfile)
        if file_name not in expected_json:
            expected_json[file_name] = json_hash
            openfile.seek(0)
            json.dump(expected_json, openfile, indent=4)
        elif expected_json[file_name] != json_hash:
            mismatch = True
            print(f"Hash mismatch: {file_name}")
    if save_file and not mismatch:
        save_path = os.path.join(EXPECTED_RESULTS_FODLER, file_name.strip('docx'))
        save_path += 'json'
        with open(save_path, 'w') as openfile:
            json.dump(questions, openfile, indent=4)



# for file_name in os.listdir(DOCS_FOLDER):
#     if not file_name.endswith('docx'):
#         continue
#     print(f'processing {file_name}')
#     process_file(file_name)

process_file("Psycho_2025_a.docx", True)
