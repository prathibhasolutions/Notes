from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any


class Person:
    """Abstract base for all human-role entities."""

    def __init__(self, person_id: str, name: str, email: str, age: int):
        self._person_id = person_id
        self._name = name
        self._email = email
        self._age = age

    @property
    def person_id(self):
        return self._person_id

    @property
    def name(self):
        return self._name

    @property
    def email(self):
        return self._email

    @property
    def age(self):
        return self._age

    def get_contact_info(self) -> str:
        return f"{self._name} <{self._email}>"

    def __str__(self) -> str:
        return f"{self.__class__.__name__}(id={self._person_id}, name={self._name})"

    def __repr__(self) -> str:
        return self.__str__()


class Employee(Person):
    """Base class for staff and faculty."""

    def __init__(self, person_id: str, name: str, email: str, age: int, employee_id: str, department: str):
        super().__init__(person_id, name, email, age)
        self._employee_id = employee_id
        self._department = department

    @property
    def employee_id(self):
        return self._employee_id

    @property
    def department(self):
        return self._department

    def work(self):
        return "Working on learning operations"


class Student(Person):
    """Represents a learner. Inherits Person."""

    def __init__(self, person_id: str, name: str, email: str, age: int, student_id: str, course_ids=None):
        super().__init__(person_id, name, email, age)
        self._student_id = student_id
        self._course_ids = course_ids or []
        self._marks = {}

    @property
    def student_id(self):
        return self._student_id

    @property
    def course_ids(self):
        return list(self._course_ids)

    def add_course(self, course_id: str):
        if course_id not in self._course_ids:
            self._course_ids.append(course_id)

    def add_mark(self, course_id: str, mark: float):
        self._marks[course_id] = mark

    def get_average_mark(self):
        if not self._marks:
            return 0
        return sum(self._marks.values()) / len(self._marks)

    def __len__(self):
        return len(self._course_ids)

    def __contains__(self, course_id):
        return course_id in self._course_ids

    def __iter__(self):
        return iter(self._course_ids)


class Teacher(Employee):
    """Represents a teacher. Inherits Employee."""

    def __init__(self, person_id: str, name: str, email: str, age: int, employee_id: str, department: str,
                 teacher_id: str, subjects: List[str] = None):
        super().__init__(person_id, name, email, age, employee_id, department)
        self._teacher_id = teacher_id
        self._subjects = subjects or []

    @property
    def teacher_id(self):
        return self._teacher_id

    @property
    def subjects(self):
        return list(self._subjects)

    def teach(self):
        return f"{self.name} is conducting a class"


class AuditLogMixin:
    """Mixin used to show object-oriented mixin behavior."""

    def log_event(self, message: str):
        self._audit_log.append(message)


class Administrator(Employee, AuditLogMixin):
    """Multiple inheritance example: Administrator inherits Employee and AuditLogMixin."""

    def __init__(self, person_id: str, name: str, email: str, age: int, employee_id: str, department: str,
                 admin_id: str):
        Employee.__init__(self, person_id, name, email, age, employee_id, department)
        self._admin_id = admin_id
        self._audit_log = []

    @property
    def admin_id(self):
        return self._admin_id

    def add_audit_event(self, event: str):
        self._audit_log.append(event)

    def get_audit_log(self):
        return list(self._audit_log)


class Department:
    """Department composition member."""

    def __init__(self, department_id: str, name: str):
        self.department_id = department_id
        self.name = name

    def __str__(self):
        return f"Department({self.name})"


class Course:
    """Course model."""

    def __init__(self, course_id: str, title: str, duration_weeks: int, credits: int, teacher_id: str = None):
        self._course_id = course_id
        self._title = title
        self._duration_weeks = duration_weeks
        self._credits = credits
        self._teacher_id = teacher_id

    @property
    def course_id(self):
        return self._course_id

    @property
    def title(self):
        return self._title

    @property
    def duration_weeks(self):
        return self._duration_weeks

    @property
    def credits(self):
        return self._credits

    @property
    def teacher_id(self):
        return self._teacher_id

    def __repr__(self):
        return f"Course(course_id={self._course_id}, title={self._title})"


class Enrollment:
    """Relationship object connecting a student to a course."""

    def __init__(self, enrollment_id: str, student_id: str, course_id: str, status: str = 'active'):
        self.enrollment_id = enrollment_id
        self.student_id = student_id
        self.course_id = course_id
        self.status = status

    def __str__(self):
        return f"Enrollment({self.enrollment_id}: {self.student_id}->{self.course_id})"


class School:
    """Aggregate root. It composes departments, courses, and persons."""

    def __init__(self, school_id: str, name: str):
        self.school_id = school_id
        self.name = name
        self.departments: Dict[str, Department] = {}
        self.students: Dict[str, Student] = {}
        self.teachers: Dict[str, Teacher] = {}
        self.administrators: Dict[str, Administrator] = {}
        self.courses: Dict[str, Course] = {}
        self.enrollments: Dict[str, Enrollment] = {}

    def add_department(self, department: Department):
        self.departments[department.department_id] = department

    def add_student(self, student: Student):
        self.students[student.student_id] = student

    def add_teacher(self, teacher: Teacher):
        self.teachers[teacher.teacher_id] = teacher

    def add_administrator(self, administrator: Administrator):
        self.administrators[administrator.admin_id] = administrator

    def add_course(self, course: Course):
        self.courses[course.course_id] = course

    def add_enrollment(self, enrollment: Enrollment):
        self.enrollments[enrollment.enrollment_id] = enrollment

    def get_student(self, student_id):
        return self.students.get(student_id)

    def get_course(self, course_id):
        return self.courses.get(course_id)

    def __iter__(self):
        return iter(self.students.values())

    def __len__(self):
        return len(self.students)

    def __contains__(self, student_id):
        return student_id in self.students

    def __getitem__(self, index):
        return list(self.students.values())[index]

