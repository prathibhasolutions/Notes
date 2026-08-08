# Python OOP + Functional Programming School Management System

This project is a complete, multi-file Python application that demonstrates both object-oriented programming (OOP) and functional programming (FP) ideas in one connected business-style system.

## Project Goal

The system manages a small school or training academy. It supports:

- Students
- Teachers
- Administrators
- Departments
- Courses
- Enrollments
- Reports
- File-based JSON storage acting as a lightweight database

## Architectural Summary

The project is divided into multiple Python modules so each concept has a clear responsibility:

- `main.py` starts the interactive command-line interface.
- `src/models.py` contains the core classes and OOP design building blocks.
- `src/storage.py` handles saving and loading data from JSON files.
- `src/services.py` provides business logic and functional transformations.
- `src/utils.py` contains reusable helper utilities.

## Core OOP Concepts Implemented

- Abstraction via the `Person` base class and `SchoolReport` service contracts.
- Encapsulation through private attributes and property accessors.
- Inheritance through `Student`, `Teacher`, and `Administrator` classes.
- Multiple inheritance through `Administrator(Employee, AuditLogMixin)`.
- Composition through a `School` object owning departments, student records, teacher records, and course records.
- Aggregation through course and student references being managed in collections.
- Magic methods like `__str__`, `__repr__`, `__iter__`, `__contains__`, `__len__`, and `__getitem__`.

## Functional Programming Ideas

The project uses:

- List comprehensions
- Lambda expressions
- `map()` and `filter()`
- `reduce()` use cases
- Sorted reports and aggregation logic
- Functional transformations for generating dashboards and reports

## File Handling as a Database

Data is stored in JSON files under the `data/` folder. All basic CRUD operations are implemented through file reads and writes.

## Features

- Register students
- Register teachers
- Register administrators
- Create courses
- Enroll students in courses
- Add marks and compute grade summaries
- Generate dashboards
- Save and reload all data from the JSON database
- Command-line menu for all major operations

## How to Run

```bash
python main.py
```

## Notes

This is a simple educational project that shows how to build a more realistic layered Python application.
