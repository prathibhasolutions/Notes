import random
import string
from functools import reduce


def generate_id(prefix: str = 'ID'):
    return f"{prefix}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=6))}"


def safe_number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def calculate_average(values):
    if not values:
        return 0
    return reduce(lambda total, current: total + current, values) / len(values)


def map_names(persons):
    return list(map(lambda person: person.name, persons))


def filter_failed_students(students):
    return list(filter(lambda student: getattr(student, 'get_average_mark', lambda: 0)() < 40, students))


def to_dict(obj):
    if hasattr(obj, '__dict__'):
        return obj.__dict__
    return obj


def nested_menu_example():
    matrix = [
        ['C', 'O', 'R', 'E'],
        ['P', 'Y', 'T', 'H'],
        ['O', 'N', 'O', 'O'],
    ]

    result = []
    for row in matrix:
        inner = []
        for char in row:
            inner.append(char)
        result.append(inner)
    return result
