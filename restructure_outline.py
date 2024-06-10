"""
Naming convention

Group: bakery_(Module)_[a | b | primer]
bakery_(Module)_primer_[read | quiz | code]
bakery_(Module)_(Lesson)_read
bakery_(Module)_(Lesson)_quiz
bakery_(Module)_(Lesson)_code_(Problem)

Unit (month/chunk of course)
Chapter (Module)
Section (Primer/A/B)
Lesson
Assignment (Read | Quiz | Code)

bakery_logic

1-intro
    A/
    B/
    Primer/

Modules,
intro,
functions,
if,
structures,
for,
sequences,
nesting,
time,
"""
import json
import os
from io import StringIO
from pprint import pprint
from markdown import Markdown
import frontmatter
from frontmatter.default_handlers import YAMLHandler
from ruamel.yaml import YAML

yaml = YAML()
yaml.default_flow_style = False
yaml.allow_unicode=True


class RuamelYamlHandler(YAMLHandler):
    def load(self, fm, **kwargs):
        return yaml.load(StringIO(fm))

class MarkdownFile:
    def __init__(self, filename):
        self.filename = filename
    def __enter__(self):
        with open(self.filename) as existing_file:
            raw = existing_file.read()
        self.metadata, self.waltz, self.content = extract_front_matter(raw)
        return self
    def __exit__(self, type, value, traceback):
        if type is None:
            self.save()
            return True
        else:
            # Failed to save, rethrow exception
            return False
    def save(self):
        new_version = add_to_front_matter(self.content, self.waltz)
        with open(self.filename, 'w') as existing_file:
            existing_file.write(new_version)

def extract_front_matter(text):
    data = frontmatter.loads(text, handler=RuamelYamlHandler())
    regular_metadata = data.metadata
    front_matter_metadata = regular_metadata.pop('waltz', {})
    return regular_metadata, front_matter_metadata, data.content

def add_to_front_matter(markdown, yaml):
    if markdown.startswith("---"):
        data = frontmatter.loads(markdown, handler=yaml)
        regular_metadata = data.metadata
        if 'waltz' not in regular_metadata:
            regular_metadata['waltz'] = {}
        regular_metadata['waltz'].update(yaml)
        yaml = regular_metadata
        markdown = data.content
    else:
        yaml = {'waltz': yaml}
    return inject_yaml(markdown, yaml)


def inject_yaml(markdown, yaml_data):
    stream = StringIO()
    yaml.dump(yaml_data, stream)
    return "---\n{}---\n{}".format(stream.getvalue(), markdown)

class Tracker:
    def __init__(self):
        self.value = None
        self.index = 0
        self.previous = None
        self.is_new = False
    def update(self, value):
        if self.value != value:
            self.previous = self.value
            self.value = value
            self.index += 1
            self.is_new = True
        else:
            self.is_new = False
    def restart(self):
        self.index = 0
        self.value = None
        self.previous = None
        self.is_new = False

def assert_path(path):
    if not os.path.exists(path):
        print("Missing path:", repr(path))
        return False
    return True

STARTER = """---
waltz:
  display title: '{title}'
  title: {url}
  resource: problem
  type: {resource_type}
---"""

def merge_dictionaries(a: dict, b: dict) -> dict:
    """
    Merges b into a, recursively
    https://stackoverflow.com/a/60353078/1718155
    """
    for key in b:
        if isinstance(a.get(key), dict) or isinstance(b.get(key), dict):
            merge_dictionaries(a[key], b[key])
        else:
            a[key] = b[key]
    return a

EMPTY = """---\n---\n"""

def basic_data(existing_data):
    result = {
        'identity': {
            'course id': 12,
            'created': 'June 28 2022, 1500',
            'modified': 'June 28 2022, 1500',
            'owner email': 'acbart@udel.edu',
            'owner id': 1,
            'version downloaded': 1,
        },
        'resource': 'problem'
    }
    if existing_data.get('identity', {}).get('created') is not None:
        del result['identity']['created']
    if existing_data.get('identity', {}).get('version downloaded') is not None:
        del result['identity']['version downloaded']
    return result

def create_if_needed(path, contents=EMPTY):
    if not os.path.exists(path):
        with open(path, 'w') as out:
            out.write(contents)


def make_group(module_index, module_title, part_index, part_title):
    if part_title.lower() == 'primer':
        group_title = f"{module_index}) {module_title} {part_title.title()}"
    else:
        group_title = f"{module_index}{part_title.title()}) {module_title}"
    group_url = f"bakery_{new_module}_{part_title.lower()}"
    return {
        "_schema_version": 2,
        "name": group_title,
        "url": group_url,
        "owner_id": 1,
        "owner_id__email": "acbart@udel.edu",
        "course_id": 12,
        "position": (module_index-1) * 3 + part_index,
    }


with open('../modules/outline.csv') as outline_file:
    modules = [[piece.strip() for piece in line.split(',')
                if piece.strip()]
                for line in outline_file][1:]

MODULE_FOLDER = "../modules/{index}-{name}/"

module = Tracker()
part = Tracker()
lesson = Tracker()
groups = []
all_resources, all_coding = [], []
for new_module, module_title, new_part, new_lesson, lesson_title, *problems in modules:
    module.update(new_module)
    part.update(new_part)
    if module.is_new:
        print("Starting module:", new_module)
    if module.is_new or part.is_new:
        lesson.restart()
    lesson.update(new_lesson)
    
    # Confirm module/primer/part/lesson folder
    module_folder = MODULE_FOLDER.format(index=module.index, name=module.value)
    primer_folder = module_folder + "primer/"
    part_folder = module_folder + f"{new_part}/"
    lesson_folder = part_folder + f"{lesson.index}-{lesson.value}/"
    os.makedirs(lesson_folder, exist_ok=True)
    if not all(assert_path(path) for path in
                [module_folder, primer_folder, part_folder, lesson_folder]):
        print("Skipping:", new_module, new_part, new_lesson)
        continue
    # Make primer
    if module.is_new:
        primer_path = primer_folder + f"bakery_{new_module}_primer_read.md"
        create_if_needed(primer_path)
    # Make quiz and reading
    assignment_lead = f"bakery_{new_module}_{new_lesson}"
    reading_path = lesson_folder + assignment_lead + "_read.md"
    quiz_path = lesson_folder + assignment_lead + "_quiz.md"
    all_resources.extend([assignment_lead+"_read", assignment_lead+"_quiz"])
    create_if_needed(reading_path)
    create_if_needed(quiz_path)
    # Make any coding problems
    for problem_index, problem in enumerate(problems, 1):
        coding_path = lesson_folder + assignment_lead + f"_code_{problem}/"
        os.makedirs(coding_path, exist_ok=True)
        create_if_needed(coding_path+'index.md')
        with MarkdownFile(coding_path+"index.md") as coding:
            merge_dictionaries(coding.waltz, basic_data(coding.waltz))
            coding.waltz['type'] = 'blockpy'
            problem_title_guess = problem.replace("_", " ").title()
            if 'display title' not in coding.waltz:
                coding.waltz['display title'] = f"{module.index}{part.value.upper()}{lesson.index}.{problem_index}) {problem_title_guess}"
            coding.waltz['title'] = assignment_lead+f"_code_{problem}"
            if 'files' not in coding.waltz:
                coding.waltz['files'] = {'path': assignment_lead+f"_code_{problem}"}
            if 'visibility' not in coding.waltz:
                quiz.waltz['visibility'] = { 'publicly indexed': True }
            if 'additional settings' not in coding.waltz:
                if module.index == 1:
                    coding.waltz['additional settings'] = {'start_view': 'split'}
                else:
                    coding.waltz['additional settings'] = {'start_view': 'text'}
        create_if_needed(coding_path+"on_run.py", "from pedal import *\n")
        create_if_needed(coding_path+"starting_code.py", "\n")
        all_coding.append(assignment_lead + f"_code_{problem}/")
    # Fill out contents
    lesson_title = f"{module.index}{part.value.upper()}{lesson.index}) {lesson_title}"
    with MarkdownFile(reading_path) as reading:
        merge_dictionaries(reading.waltz, basic_data(reading.waltz))
        reading.waltz['type'] = 'reading'
        reading.waltz['display title'] = lesson_title + " Reading"
        reading.waltz['title'] = assignment_lead + "_read"
        reading.waltz['visibility'] = {
            'subordinate': True,
            'publicly indexed': True,
        }
    with MarkdownFile(quiz_path) as quiz:
        merge_dictionaries(quiz.waltz, basic_data(quiz.waltz))
        quiz.waltz['type'] = 'quiz'
        quiz.waltz['display title'] = lesson_title
        quiz.waltz['title'] = assignment_lead + "_quiz"
        quiz.waltz['visibility'] = {
            'publicly indexed': True
        }
    # Group Formation
    if module.is_new:
        groups.append(make_group(module.index, module_title, 0, "primer"))
    if part.is_new:
        groups.append(make_group(module.index, module_title, part.index, new_part))
    # Wrap-up
    print(f"  Finished lesson {new_part}/{new_lesson}, {len(problems)} problems")

with open('../modules/bakery_groups.json', 'w') as group_file:
    json.dump({"groups": groups}, group_file, indent=2)

with open('../modules/push_script.bat', 'w') as push_script:
    for resource in all_resources:
        print(f"waltz push blockpy problem {resource} --combine", file=push_script)
with open('../modules/push_coding.bat', 'w') as push_coding:
    for resource in all_coding:
        print(f"waltz push blockpy problem {resource}", file=push_coding)