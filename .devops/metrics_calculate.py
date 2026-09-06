#!/usr/bin/env python3

# pyright: strict, reportUnknownMemberType=false, reportMissingTypeStubs=false

import argparse
import csv
import hashlib
import json
import math
import os
import re
import subprocess
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, Optional, Union

import yaml
from lxml import etree
from tabulate import tabulate

parser = argparse.ArgumentParser()
parser.add_argument('config_yaml', type=Path)
parser.add_argument('output_directory', type=Path)
parser.add_argument('-s', '--skip-generation', action='store_true', help='Skip generating metrics.xml and duplicates.csv')
args = parser.parse_args()

SCRIPT_DIR = Path(__file__).parent.resolve()  # .devops
DOTNET_TOOLS_DIR = SCRIPT_DIR / 'dotnet-tools'
# `dotnet tool` launchers carry a .exe suffix on Windows but not on macOS/Linux.
TOOL_SUFFIX = '.exe' if os.name == 'nt' else ''
OUTPUT_DIR: Path = Path(args.output_directory).resolve()
METRICS_XML_PATH = OUTPUT_DIR / 'metrics.xml'
DUPLICATES_CSV_PATH = OUTPUT_DIR / 'duplicates.csv'

Value = Union[str, int, float]
Color = Literal['red', 'green', 'yellow']
ThresholdType = Literal['<=', '>=']
Severity = Literal['info', 'minor', 'major', 'critical', 'blocker']


class Threshold:
    def __init__(self,
                 name: str,
                 badge_id: str,
                 threshold_type: ThresholdType,
                 good_threshold: int,
                 acceptable_threshold: Optional[int] = None,
                 suffix: str = '',
                 ignore: Optional[list[str]] = None,
                 ignore_classes: Optional[list[str]] = None,
                 ignore_methods: Optional[list[str]] = None,
                 ignore_projects: Optional[list[str]] = None,
                 ignore_directories: Optional[list[str]] = None,
                 fail_build: bool = True,
                 ):
        self.name = name
        self.badge_id = badge_id
        self.threshold_type = threshold_type
        self.good_threshold = good_threshold
        self.suffix = suffix
        # When False, a violation of this threshold is reported but does not fail the build.
        self.fail_build = fail_build
        # Support legacy 'ignore' field as ignore_classes fallback
        self.ignore_classes: list[str] = ignore_classes or ignore or []
        self.ignore_methods: list[str] = ignore_methods or []
        self.ignore_projects: list[str] = ignore_projects or []
        self.ignore_directories: list[str] = ignore_directories or []
        if acceptable_threshold is not None:
            self.acceptable_threshold = acceptable_threshold
        else:
            self.acceptable_threshold = good_threshold

    def _leq_color(self, value: Value) -> Color:
        '''Return green if metric <= good_threshold, else yellow if metric <= acceptable_threshold, else red.'''
        try:
            value = float(value)
        except:
            return 'red'

        if value <= self.good_threshold:
            return 'green'
        elif value <= self.acceptable_threshold:
            return 'yellow'
        return 'red'

    def _geq_color(self, value: Value) -> Color:
        '''Return green if metric >= good_threshold, else yellow if metric >= acceptable_threshold, else red.'''
        try:
            value = float(value)
        except:
            return 'red'

        if value >= self.good_threshold:
            return 'green'
        elif value >= self.acceptable_threshold:
            return 'yellow'
        return 'red'

    def get_color(self, value: Value) -> Color:
        if self.threshold_type == '<=':
            return self._leq_color(value)
        elif self.threshold_type == '>=':
            return self._geq_color(value)
        else:
            raise ValueError(f'Unknown threshold type: {self.threshold_type}')

    def is_violation(self, value: Value) -> bool:
        return self.get_color(value) == 'red'


class ConfigParser:
    def __init__(self, config_path: Path):
        with open(config_path, 'r') as f:
            self._config: dict[str, Any] = yaml.safe_load(f)

        self.solution: str = self._try_get('solution', '')
        self.ignore_projects: list[str] = self._try_get('ignore_projects', [])
        self.ignore_directories: list[str] = self._try_get('ignore_directories', [])
        self.ignore_classes: list[str] = self._try_get('ignore_classes', [])
        self.ignore_methods: list[str] = self._try_get('ignore_methods', [])

        self.duplicate_code = self._parse_threshold('duplicate_code')
        if self.duplicate_code is not None:
            self.duplicate_min_tokens: int = self._config['duplicate_code']['min_tokens']
            self.duplicate_block_split: str = self._config['duplicate_code']['block_split']

        self.todos = self._parse_threshold('todos')
        self.not_implemented = self._parse_threshold('not_implemented')
        self.maintainability = self._parse_threshold('maintainability')
        self.depth_of_inheritance = self._parse_threshold('depth_of_inheritance')
        self.cyclomatic_complexity = self._parse_threshold('cyclomatic_complexity')
        self.class_coupling = self._parse_threshold('class_coupling')
        self.class_coupling_exclude_types: list[str] = self._config.get('class_coupling', {}).get('exclude_types', [])
        self.class_num_fields = self._parse_threshold('class_num_fields')
        self.class_num_public_methods = self._parse_threshold('class_num_public_methods')
        self.class_cohesion = self._parse_threshold('class_cohesion')
        if self.class_cohesion is not None:
            self.class_cohesion_min_fields: int = self._config['class_cohesion']['min_fields']
            self.class_cohesion_min_methods: int = self._config['class_cohesion']['min_methods']

    def _parse_threshold(self, key: str) -> Optional[Threshold]:
        if key not in self._config:
            return None
        t = self._config[key]
        if 'enabled' in t and str(t['enabled']).lower() == 'false':
            return None
        return Threshold(
            name=t['name'],
            badge_id=key,
            threshold_type=t['threshold_type'],
            good_threshold=t['good'],
            acceptable_threshold=t.get('acceptable'),
            suffix=t.get('suffix', ''),
            ignore=t.get('ignore'),
            ignore_classes=t.get('ignore_classes'),
            ignore_methods=t.get('ignore_methods'),
            ignore_projects=t.get('ignore_projects'),
            ignore_directories=t.get('ignore_directories'),
            fail_build=t.get('fail_build', True),
        )

    def _try_get(self, key: str, default: Any = None) -> Any:
        return self._config.get(key, default)


config = ConfigParser(args.config_yaml)

SOLUTION_PATH: Path = (args.config_yaml.parent / config.solution).resolve()
SOLUTION_DIR = SOLUTION_PATH.parent.resolve()


# The slt-csharp-metrics analyzer synthesizes a <Program> type with these phantom
# entry-point methods (one per source file) even in library projects — filter them out.
TOP_LEVEL_ENTRY_POINT_NAME = '<top-level-statements-entry-point>'


# Helper functions
def sln_relative_path(path: Union[str, Path]) -> str:
    x: Union[str, Path]
    try:
        x = Path(path).relative_to(SOLUTION_DIR)
    except ValueError:
        x = path
    return str(x).replace('\\', '/')


def ignore_class(class_name: str, badge: Optional[Threshold] = None) -> bool:
    for pattern in config.ignore_classes:
        if re.match(pattern, class_name):
            return True
    if badge is not None:
        for pattern in badge.ignore_classes:
            if re.match(pattern, class_name):
                return True
    return False


def ignore_method(method_name: str, badge: Optional[Threshold] = None) -> bool:
    for pattern in config.ignore_methods:
        if re.match(pattern, method_name):
            return True
    if badge is not None:
        for pattern in badge.ignore_methods:
            if re.match(pattern, method_name):
                return True
    return False


def ignore_project(project_name: str, badge: Optional[Threshold] = None) -> bool:
    for pattern in config.ignore_projects:
        if re.match(pattern, project_name):
            return True
    if badge is not None:
        for pattern in badge.ignore_projects:
            if re.match(pattern, project_name):
                return True
    return False


def ignore_directory(file_path: str, badge: Optional[Threshold] = None) -> bool:
    for dir_name in config.ignore_directories:
        if f'/{dir_name}/' in file_path or file_path.startswith(f'{dir_name}/'):
            return True
    if badge is not None:
        for dir_name in badge.ignore_directories:
            if f'/{dir_name}/' in file_path or file_path.startswith(f'{dir_name}/'):
                return True
    return False


def ignore_class_fields(class_name: str) -> bool:
    if config.class_num_fields is None:
        return False
    return ignore_class(class_name, config.class_num_fields)


def ignore_class_methods(class_name: str) -> bool:
    if config.class_num_public_methods is None:
        return False
    return ignore_class(class_name, config.class_num_public_methods)


@dataclass
class Metrics:
    _metrics_table: list[list[str]]
    lines_unshared: dict[str, int]

    all_pass: bool = True

    lines_shared: int = 0
    shared_percent: float = 0

    todos: int = 0
    not_implemented: int = 0
    duplicate_methods: float = 0

    maintainability: int = 0
    max_depth: int = 0
    max_complexity: int = 0
    max_coupling: int = 0
    min_cohesion: float = 100.0
    max_fields: int = 0
    max_methods: int = 0

    def _add_metric_to_table(self, value: Value, badge: Optional[Threshold]) -> None:
        if badge is None:
            return

        color_name = badge.get_color(value)
        threshold_str = f'{badge.threshold_type} {badge.acceptable_threshold}{badge.suffix}'

        if color_name == 'green':
            status = 'GOOD'
            is_pass = True
        elif color_name == 'yellow':
            status = 'ACCEPTABLE'
            is_pass = True
        else:
            status = 'FAIL' if badge.fail_build else 'FAIL (non-blocking)'
            is_pass = False

        # A violation only fails the build for blocking metrics; non-blocking ones are
        # reported (and still published to GitLab) but do not affect the exit code.
        if not is_pass and badge.fail_build:
            self.all_pass = False

        self._metrics_table.append([badge.name, str(round(float(value) * 100) / 100), threshold_str, status])

    def metric_badge_values(self) -> list[tuple[Threshold, Value]]:
        '''Map each configured metric threshold to its computed value.'''
        pairs: list[tuple[Optional[Threshold], Value]] = [
            (config.maintainability, self.maintainability),
            (config.todos, self.todos),
            (config.not_implemented, self.not_implemented),
            (config.duplicate_code, self.duplicate_methods),
            (config.depth_of_inheritance, self.max_depth),
            (config.cyclomatic_complexity, self.max_complexity),
            (config.class_coupling, self.max_coupling),
            (config.class_cohesion, self.min_cohesion),
            (config.class_num_fields, self.max_fields),
            (config.class_num_public_methods, self.max_methods),
        ]
        return [(badge, value) for badge, value in pairs if badge is not None]

    def create_metrics_table(self) -> str:
        for badge, value in self.metric_badge_values():
            self._add_metric_to_table(value, badge)
        return tabulate(
            self._metrics_table,
            ['Metrics', 'Value', 'Target', 'Status'],
            floatfmt='.0f')


class MetricsCalculator:
    def __init__(self, total_methods: Optional[int] = None):
        self._metrics = Metrics(_metrics_table=[], lines_unshared={})
        self._total_methods = total_methods

        # Load metrics xml
        self._tree = etree.parse(METRICS_XML_PATH)
        self._root = self._tree.getroot()

        assemblies = self._root.xpath('Targets/Target/Assembly')
        self._assemblies: dict[str, Any] = {}
        for assembly in assemblies:
            project_name = assembly.getparent().get('Name')
            self._assemblies[project_name] = assembly

        badges = [
            config.maintainability, config.depth_of_inheritance, config.cyclomatic_complexity,
            config.class_coupling, config.class_cohesion, config.class_num_fields, config.class_num_public_methods,
            config.todos, config.not_implemented
            ]
        self._violations: dict[str, list[tuple[str, Value, str, int]]] = {}
        for badge in badges:
            if badge is not None:
                self._violations[badge.badge_id] = []

        self._code_quality: list[dict[str, Any]] = []

        self._coupling_coupled_classes: dict[tuple[str, int], list[str]] = {}
        self._cohesion_field_method_counts: dict[tuple[str, int], tuple[int, int]] = {}

    def _get_filtered_xpaths(self, node_type: str = 'NamedType') -> str:
        '''Build XPath selecting node_type only from non-ignored projects,
        excluding the analyzer's synthesized top-level-statements entry points.'''
        if node_type == 'Method':
            predicate = f'[@Name!="{TOP_LEVEL_ENTRY_POINT_NAME}"]'
        elif node_type == 'NamedType':
            predicate = f'[not(Members/Method/@Name="{TOP_LEVEL_ENTRY_POINT_NAME}")]'
        else:
            predicate = ''
        xpaths: list[str] = []
        for project_name in self._assemblies:
            if not ignore_project(project_name):
                xpaths.append(f'Targets/Target[@Name="{project_name}"]/Assembly//{node_type}{predicate}')
        return ' | '.join(xpaths)

    def passed_all(self) -> bool:
        return self._metrics.all_pass

    def save_metrics(self) -> None:
        report_path = OUTPUT_DIR / 'metrics.txt'
        with open(report_path, mode='w', encoding='utf-8') as f:
            if config.todos is not None:
                f.write(f'todos {self._metrics.todos}\n')
            if config.not_implemented is not None:
                f.write(f'not_implemented_exceptions {self._metrics.not_implemented}\n')
            if config.duplicate_code is not None:
                f.write(f'duplicate_methods_percent {self._metrics.duplicate_methods:.2f}\n')
            if config.maintainability is not None:
                f.write(f'maintainability {self._metrics.maintainability}\n')
            if config.depth_of_inheritance is not None:
                f.write(f'depth_of_inhertance_max {self._metrics.max_depth}\n')
            if config.cyclomatic_complexity is not None:
                f.write(f'cyclomatic_complexity_max {self._metrics.max_complexity}\n')
            if config.class_coupling is not None:
                f.write(f'class_coupling_max {self._metrics.max_coupling}\n')
            if config.class_cohesion is not None:
                f.write(f'class_cohesion_min {self._metrics.min_cohesion}\n')

    def _save_violations_csv(self,
                              badge: Optional[Threshold],
                              csv_filename: str,
                              header: list[str],
                              sort_descending: bool,
                              extra_columns_fn: Optional[Callable[[str, Value, str, int], list[Any]]] = None,
                              value_formatter: Optional[Callable[[Value], Any]] = None) -> None:
        if badge is None:
            return
        violations = self._violations[badge.badge_id]
        csv_path = OUTPUT_DIR / csv_filename
        with open(csv_path, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(header)
            sorted_violations = sorted(violations, key=lambda x: float(x[1]), reverse=sort_descending)
            for name, value, file_name, line_num in sorted_violations:
                formatted_value = value_formatter(value) if value_formatter is not None else value
                row: list[Any] = [name, formatted_value, file_name, line_num]
                if extra_columns_fn is not None:
                    extras = extra_columns_fn(name, value, file_name, int(line_num))
                    # Insert extra columns before File/Line for readability
                    row = [name, formatted_value, *extras, file_name, line_num]
                writer.writerow(row)

    def save_complexity_csv(self) -> None:
        self._save_violations_csv(
            config.cyclomatic_complexity,
            'cyclomatic_complexity.csv',
            ['Method Name', 'Cyclomatic Complexity', 'File', 'Line'],
            sort_descending=True,
        )

    def save_cohesion_csv(self) -> None:
        if config.class_cohesion is None:
            return

        def extras(_name: str, _value: Value, file_name: str, line_num: int) -> list[Any]:
            fields, methods = self._cohesion_field_method_counts.get((file_name, line_num), (0, 0))
            return [fields, methods]

        self._save_violations_csv(
            config.class_cohesion,
            'class_cohesion.csv',
            ['Class Name', 'Cohesion (%)', 'Fields', 'Methods', 'File', 'Line'],
            sort_descending=False,
            extra_columns_fn=extras,
            value_formatter=lambda v: f'{float(v):.2f}',
        )

    def save_coupling_csv(self) -> None:
        if config.class_coupling is None:
            return
        violations = self._violations[config.class_coupling.badge_id]
        sorted_violations = sorted(violations, key=lambda x: float(x[1]), reverse=True)
        csv_path = OUTPUT_DIR / 'class_coupling.csv'
        with open(csv_path, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Class Name', 'Coupling', 'File', 'Line', 'Coupled Class'])
            for name, value, file_name, line_num in sorted_violations:
                coupled = sorted(self._coupling_coupled_classes.get((file_name, int(line_num)), []))
                if not coupled:
                    writer.writerow([name, value, file_name, line_num, ''])
                    continue
                for coupled_class in coupled:
                    writer.writerow([name, value, file_name, line_num, coupled_class])

    def save_code_quality(self) -> None:
        self._add_code_quality(config.depth_of_inheritance, 'major',
            "Depth of inheritance for class '{name}' is {value}, but should be {threshold_type} {threshold}")

        self._add_code_quality(config.maintainability, 'major',
            "Maintainability index for class '{name}' is {value}, but should be {threshold_type} {threshold}")

        self._add_code_quality(config.cyclomatic_complexity, 'major',
            "Cyclomatic complexity for method '{name}' is {value}, but should be {threshold_type} {threshold}")

        self._add_code_quality(config.class_coupling, 'major',
            "Class coupling for class '{name}' is {value}, but should be {threshold_type} {threshold}")

        self._add_code_quality(config.class_cohesion, 'major',
            "Class cohesion for class '{name}' is {value}, but should be {threshold_type} {threshold}")

        self._add_code_quality(config.class_num_fields, 'major',
            "Class '{name}' has {value} non-constant fields, but it should have {threshold_type} {threshold}")

        self._add_code_quality(config.class_num_public_methods, 'major',
            "Class '{name}' has {value} public methods, but it should have {threshold_type} {threshold}")

        self._add_code_quality(config.todos, 'minor', "Unresolved {name}")
        self._add_code_quality(config.not_implemented, 'minor', "Unresolved {name}")

        report_path = OUTPUT_DIR / 'code_quality.json'
        with open(report_path, mode='w', encoding='utf-8') as f:
            json.dump(self._code_quality, f, indent=2)

    def calc_metrics(self) -> None:
        self._calc_todos()
        self._calc_all_class_metrics()
        self._calc_complexity()
        self._calc_core_class_metrics()
        self._parse_duplicate_methods()

    def print_metrics(self) -> None:
        print('Code Metrics')
        print(self._metrics.create_metrics_table())

        self._print_violations(
           config.depth_of_inheritance,
           'Depth of Inheritance Violations',
           'Class Name',
           'Depth')

        self._print_violations(
           config.maintainability,
           'Maintainability Violations',
           'Class Name',
           'Maintainability')

        self._print_violations(
           config.cyclomatic_complexity,
           'Cyclomatic Complexity Violations',
           'Method Name',
           'Complexity')

        self._print_violations(
            config.class_coupling,
            'Class Coupling Violations',
            'Class Name',
            'Coupling')

        self._print_violations(
            config.class_cohesion,
            'Class Cohesion Violations',
            'Class Name',
            'Class Cohesion (%)')

        self._print_violations(
            config.class_num_fields,
            'Number of Fields in Class Violations',
            'Class Name',
            '# Fields')

        self._print_violations(
            config.class_num_public_methods,
            'Number of Public Methods in Class Violations',
            'Class Name',
            '# Public Methods')

    def _add_violation(self, badge: Optional[Threshold], name: str, value: Value, file_name: str, line_num: int) -> None:
        if badge is not None:
            self._violations[badge.badge_id].append((name, value, file_name, line_num))

    def _add_if_violation(self, badge: Optional[Threshold], name: str, value: Value, file_name: str, line_num: int) -> None:
        if badge is not None and badge.is_violation(value):
            self._add_violation(badge, name, value, file_name, line_num)

    def _print_violations(self, badge: Optional[Threshold], title: str, left_col: str, right_col: str) -> None:
        if badge is None:
            return

        scores = self._violations[badge.badge_id]
        if len(scores) > 0:
            print('\n\n' + title)
            scores.sort(key=lambda x: x[1], reverse=badge.threshold_type == '<=')
            print(tabulate(
                scores,
                [left_col, right_col, 'File', 'Line']))

    def _add_code_quality(self, badge: Optional[Threshold], severity: Severity, description_fmt: str) -> None:
        if badge is None:
            return

        for name, value, file_path, line_num in self._violations[badge.badge_id]:
            check_name = f'slt-metrics-{badge.badge_id}'
            fingerprint = f'{check_name}-{file_path}-{line_num}'
            threshold = badge.acceptable_threshold
            threshold_type = badge.threshold_type
            self._code_quality.append({
                'description': description_fmt.format(
                    name=name,
                    value=value,
                    threshold=threshold,
                    threshold_type=threshold_type),
                'check_name': check_name,
                'fingerprint': hashlib.md5(fingerprint.encode('utf-8')).hexdigest(),
                'severity': severity,
                'location': {
                    'path': file_path,
                    'lines': {
                        'begin': int(line_num)
                    }
                }
            })

    def _parse_duplicate_methods(self) -> None:
        if config.duplicate_code is None:
            return

        check_name = 'slt-metrics-duplicate-method'
        duplicate_count = 0
        with open(DUPLICATES_CSV_PATH, 'r') as f:
            reader = csv.DictReader(f)
            for line in reader:
                class1, class2 = line['Class1'], line['Class2']
                if ignore_class(class1, config.duplicate_code) or ignore_class(class2, config.duplicate_code):
                    continue

                method1, method2 = line['Method1'], line['Method2']
                if ignore_method(method1, config.duplicate_code) or ignore_method(method2, config.duplicate_code):
                    continue

                file1_rel = sln_relative_path(line['File1'])
                file2_rel = sln_relative_path(line['File2'])
                if ignore_directory(file1_rel, config.duplicate_code) or ignore_directory(file2_rel, config.duplicate_code):
                    continue

                project1 = file1_rel.split('/')[0]
                project2 = file2_rel.split('/')[0]
                if ignore_project(project1, config.duplicate_code) or ignore_project(project2, config.duplicate_code):
                    continue

                duplicate_count += 1
                similarity = float(line['JaccardSimilarity'])
                violation1 = (
                    line['File1'],
                    class1 + '.' + method1,
                    line['LineNumber1']
                )
                violation2 = (
                    line['File2'],
                    class2 + '.' + method2,
                    line['LineNumber2']
                )

                for v1, v2 in [(violation1, violation2), (violation2, violation1)]:
                    file1, method_a, line1 = v1
                    file2, method_b, line2 = v2
                    fingerprint = f'{check_name}-{file1}-{method_a}-{line1}-{file2}-{method_b}-{line2}'
                    description = f'Duplicate method: {method_a} is {similarity:.2%} similar to {method_b} in {file2}:{line2}'
                    self._code_quality.append({
                        'description': description,
                        'check_name': check_name,
                        'fingerprint': hashlib.md5(fingerprint.encode('utf-8')).hexdigest(),
                        'severity': 'minor',
                        'location': {
                            'path': file1,
                            'lines': {
                                'begin': int(line1)
                            }
                        }
                    })

        if self._total_methods is not None and self._total_methods > 0:
            self._metrics.duplicate_methods = duplicate_count / self._total_methods * 100
        else:
            self._metrics.duplicate_methods = duplicate_count

    def _get_xml_metric(self, xml_node: Any, metric_name: str, dtype: Callable[..., Any] = int) -> Any:
        xpath = f'Metrics/Metric[@Name="{metric_name}"]/@Value'
        return dtype(xml_node.xpath(xpath)[0])

    def _calc_todos(self) -> None:
        if config.todos is None and config.not_implemented is None:
            return

        cs_files = SOLUTION_DIR.rglob('*.cs')
        for cs_file in cs_files:
            repo_path = sln_relative_path(cs_file)
            if '/obj/Debug' in repo_path:
                continue

            class_name = os.path.splitext(os.path.basename(cs_file))[0]
            if ignore_class(class_name):
                continue

            project_name = repo_path.split('/')[0]
            if ignore_project(project_name):
                continue

            if ignore_directory(repo_path):
                continue

            with open(cs_file, 'r', encoding='utf8') as f:
                line_num = 1
                for line in f:
                    if 'TODO' in line:
                        if not ignore_class(class_name, config.todos) and not ignore_project(project_name, config.todos) and not ignore_directory(repo_path, config.todos):
                            self._metrics.todos += 1
                            todo_msg = line[line.index('TODO'):].strip()
                            self._add_violation(config.todos, todo_msg, line_num, str(cs_file), line_num)

                    if 'NotImplementedException' in line:
                        if not ignore_class(class_name, config.not_implemented) and not ignore_project(project_name, config.not_implemented) and not ignore_directory(repo_path, config.not_implemented):
                            self._metrics.not_implemented += 1
                            self._add_violation(config.not_implemented, 'NotImplementedException', line_num, str(cs_file), line_num)

                    line_num += 1

    def _calc_complexity(self) -> None:
        if config.cyclomatic_complexity is None:
            return

        methods = self._root.xpath(self._get_filtered_xpaths('Method'))
        for method in methods:
            class_name = method.getparent().getparent().get('Name')
            if ignore_class(class_name, config.cyclomatic_complexity):
                continue

            file_name = sln_relative_path(method.get('File'))
            if ignore_directory(file_name, config.cyclomatic_complexity):
                continue

            method_signature = method.get('Name')
            method_name = method_signature.split('(')[0].split(' ')[-1]
            line_num = method.get('Line')

            if ignore_method(method_name, config.cyclomatic_complexity):
                continue

            complexity = self._get_xml_metric(method, 'CyclomaticComplexity')
            self._metrics.max_complexity = max(complexity, self._metrics.max_complexity)
            self._add_if_violation(config.cyclomatic_complexity, method_name, complexity, file_name, line_num)

    def _calc_all_class_metrics(self) -> None:
        '''Calculate Maintainability and Depth of Inheritance'''
        if config.maintainability is None and config.depth_of_inheritance is None:
            return

        self._metrics.maintainability = 100
        for project_name in self._assemblies:
            if ignore_project(project_name):
                continue
            assembly = self._assemblies[project_name]
            self._metrics.maintainability = min(
                self._metrics.maintainability,
                self._get_xml_metric(assembly, 'MaintainabilityIndex'))

        classes = self._root.xpath(self._get_filtered_xpaths('NamedType'))
        for c in classes:
            class_name = c.get('Name')
            if ignore_class(class_name):
                continue
            file_name = sln_relative_path(c.get('File'))
            if ignore_directory(file_name):
                continue
            line_num = c.get('Line')

            if config.depth_of_inheritance is not None:
                if not ignore_class(class_name, config.depth_of_inheritance) and not ignore_directory(file_name, config.depth_of_inheritance):
                    depth = self._get_xml_metric(c, 'DepthOfInheritance')
                    self._metrics.max_depth = max(depth, self._metrics.max_depth)
                    self._add_if_violation(config.depth_of_inheritance, class_name, depth, file_name, line_num)

            if config.maintainability is not None:
                if not ignore_class(class_name, config.maintainability) and not ignore_directory(file_name, config.maintainability):
                    maintainability = self._get_xml_metric(c, 'MaintainabilityIndex')
                    self._add_if_violation(config.maintainability, class_name, maintainability, file_name, line_num)

    def _get_filtered_coupled_classes(self, class_node: Any) -> list[str]:
        '''Get the list of coupled classes, excluding types matching exclude_types patterns.'''
        coupling_list_values = class_node.xpath('Metrics/Metric[@Name="ProjectClassCouplingList"]/@Value')
        if not coupling_list_values:
            return []
        coupled_classes = coupling_list_values[0].split('%')
        return [c for c in coupled_classes if not any(re.match(p, c) for p in config.class_coupling_exclude_types)]

    def _calc_coupling(self, class_node: Any) -> int:
        '''Calculate class coupling, excluding types matching exclude_types patterns.'''
        if not config.class_coupling_exclude_types:
            return self._get_xml_metric(class_node, 'ProjectClassCoupling')
        return len(self._get_filtered_coupled_classes(class_node))

    def _calc_core_class_metrics(self) -> None:
        '''Calculate class coupling, number of fields, and number of public methods'''
        if config.class_coupling is None and config.class_num_fields is None and config.class_num_public_methods is None and config.class_cohesion is None:
            return

        classes = self._root.xpath(self._get_filtered_xpaths('NamedType'))
        for c in classes:
            class_name = c.get('Name')
            if ignore_class(class_name):
                continue
            file_name = sln_relative_path(c.get('File'))
            if ignore_directory(file_name):
                continue
            line_num = c.get('Line')

            coupling = self._calc_coupling(c)
            if not ignore_class(class_name, config.class_coupling) and not ignore_directory(file_name, config.class_coupling):
                self._metrics.max_coupling = max(coupling, self._metrics.max_coupling)
                self._add_if_violation(config.class_coupling, class_name, coupling, file_name, line_num)
                if config.class_coupling is not None and config.class_coupling.is_violation(coupling):
                    self._coupling_coupled_classes[(file_name, int(line_num))] = self._get_filtered_coupled_classes(c)
            if config.class_coupling is not None and not ignore_class(class_name, config.class_coupling) and config.class_coupling.is_violation(coupling):
                coupled_classes = self._get_filtered_coupled_classes(c)
                print(f'Class coupling of {coupling} for {class_name}:')
                for coupled_class in coupled_classes:
                    print('  ' + coupled_class)
                print('')

            if config.class_cohesion is not None and not ignore_class(class_name, config.class_cohesion) and not ignore_directory(file_name, config.class_cohesion):
                fields = len(c.xpath('Members/Field[@Constant="false"]'))
                methods = len(c.xpath('Members/Method'))
                lcom = self._get_xml_metric(c, 'LackOfCohesionOfMethods', float)
                if not math.isnan(lcom):
                    cohesion = (1 - lcom) * 100
                    if fields >= config.class_cohesion_min_fields and methods >= config.class_cohesion_min_methods:
                        self._metrics.min_cohesion = min(cohesion, self._metrics.min_cohesion)
                        self._add_if_violation(config.class_cohesion, class_name, cohesion, file_name, line_num)
                        if config.class_cohesion.is_violation(cohesion):
                            self._cohesion_field_method_counts[(file_name, int(line_num))] = (fields, methods)

            if not ignore_class_fields(class_name) and not ignore_directory(file_name, config.class_num_fields):
                fields = len(c.xpath('Members/Field[@Constant="false"]'))
                self._metrics.max_fields = max(fields, self._metrics.max_fields)
                self._add_if_violation(config.class_num_fields, class_name, fields, file_name, line_num)

            if not ignore_class_methods(class_name) and not ignore_directory(file_name, config.class_num_public_methods):
                methods = len(c.xpath('Members/Method[@Private="false"]'))
                self._metrics.max_methods = max(methods, self._metrics.max_methods)
                self._add_if_violation(config.class_num_public_methods, class_name, methods, file_name, line_num)


BUFFER_LINE = '#' * 20
def print_buffer(txt: str) -> None:
    txt_len = len(txt) + 2
    full_buffer = BUFFER_LINE + ('#' * txt_len) + BUFFER_LINE
    print('')
    print(full_buffer)
    print(BUFFER_LINE + ' ' + txt + ' ' + BUFFER_LINE)
    print(full_buffer)
    print('')


def run_os_command(*cmd_args: str, capture_output: bool = False) -> str:
    cmd = ' '.join(cmd_args)
    print(f'$ {cmd}')
    result = subprocess.run(cmd_args, shell=False, capture_output=capture_output, text=True)
    err = result.returncode
    if err != 0:
        print(f'ERROR: Command {cmd_args[0]} failed with return code {err}')
        exit(err)
    return result.stdout if capture_output else ''


if __name__ == '__main__':
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    duplicate_stdout = ''
    if not args.skip_generation:
        print_buffer('ANALYZING CODE METRICS')
        print('Generating metrics.xml...')
        run_os_command(
            str(DOTNET_TOOLS_DIR / f'slt-csharp-metrics{TOOL_SUFFIX}'),
            f'/solution:{SOLUTION_PATH}',
            f'/out:{METRICS_XML_PATH}'
        )

        if config.duplicate_code is not None:
            print_buffer('ANALYZING DUPLICATE CODE')
            exclude_dirs = ','.join(config.ignore_directories)
            duplicate_stdout = run_os_command(
                str(DOTNET_TOOLS_DIR / f'csharp-duplicate-detector{TOOL_SUFFIX}'),
                f'--min-tokens={config.duplicate_min_tokens}',
                f'--block-split={config.duplicate_block_split}',
                f'--exclude={exclude_dirs}',
                str(SOLUTION_DIR),
                str(OUTPUT_DIR),
                capture_output=True
            )

    print_buffer('RESULTS')
    total_methods: Optional[int] = None
    if config.duplicate_code is not None and not args.skip_generation:
        match = re.search(r'Found \d+ of (\d+) methods have minimum of \d+ tokens', duplicate_stdout)
        if match:
            total_methods = int(match.group(1))
    metrics = MetricsCalculator(total_methods=total_methods)
    metrics.calc_metrics()
    metrics.print_metrics()
    metrics.save_metrics()
    metrics.save_complexity_csv()
    metrics.save_coupling_csv()
    metrics.save_cohesion_csv()
    metrics.save_code_quality()
    if not metrics.passed_all():
        exit(1)
